import { useState, useCallback, useEffect } from "react";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, BorderStyle, WidthType, VerticalAlign, ShadingType, UnderlineType, PageNumber, Footer } from "docx";
import { saveAs } from "file-saver";

// ── Datos empresas ────────────────────────────────────────────────────────────
const EMP = {
  "Malvarrosa SpA": {
    nombre: "MALVARROSA SPA", rut: "76.475.913-3",
    rep: "MAGALY DEL CARMEN FIGUEROA ASTUDILLO", repRut: "6.974.618-7",
    dom: "Av. Ortuzar 250, Ñuñoa, Santiago",
    dir: "AV ORTUZAR #250", comuna: "Ñuñoa",
    giro: "PASTELERIA SALON DE TE - MINIMARKET"
  },
  "Eric Hansen EIRL": {
    nombre: "VENTA DE ALIMENTOS ERIC FELIPE HANSEN FIGUEROA EIRL", rut: "77.840.316-1",
    rep: "ERIC FELIPE HANSEN FIGUEROA", repRut: "19.077.228-4",
    dom: "Av. Ossa 1624, Ñuñoa, Santiago",
    dir: "AV OSSA 1624", comuna: "Ñuñoa",
    giro: "ACTIVIDADES DE RESTAURANTES Y DE SERVICIO MOVIL DE COMIDAS"
  }
};
const SUELDOS = { 42: 539000, 30: 385000, 20: 256667 };
const DIAS_C = ["Lun","Mar","Mie","Jue","Vie","Sab","Dom"];
const DIAS_F = ["Lunes","Martes","Miercoles","Jueves","Viernes","Sabado","Domingo"];
const WEBHOOK = "https://script.google.com/macros/s/AKfycbyXuD--zD-R40H-Lin09aotxP0342f2SvvT3DO9tzXtHgaAxTn1NXfzy2aTUffTHiYjcw/exec";

// ── API Trabajadores ──────────────────────────────────────────────────────────
async function apiGuardar(trabajador) {
  try {
    await fetch(WEBHOOK, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({accion:"guardar",...trabajador}) });
  } catch(e) { console.error("Error guardando trabajador:", e); }
}
async function apiDesactivar(id) {
  try {
    await fetch(WEBHOOK, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({accion:"desactivar", id}) });
  } catch(e) { console.error("Error desactivando trabajador:", e); }
}
async function apiListar(empresa) {
  try {
    const r = await fetch(WEBHOOK+"?empresa="+encodeURIComponent(empresa));
    const d = await r.json();
    return d.trabajadores || [];
  } catch(e) { return []; }
}

const FN = {
  garzon: "El trabajador se compromete a desempeniar las funciones de Garzon - Barista, las cuales comprenderan, de forma declarativa y no limitativa, las siguientes responsabilidades especificas tanto para el area de cafeteria como para el sector de tienda del local: Gestion de Barra, Servicio y Hospitalidad: Recibir y atender cordialmente a los clientes en salon, barra, formato take away o canales de delivery. Sera responsable de la toma de pedidos, sugerencias comerciales cruzadas (cafeteria y articulos de tienda) y de la preparacion de cafe especializado, tes, bebidas frias, calientes y postres, siguiendo estrictamente las recetas, gramajes y estandares esteticos de la empresa. Asimismo, debera calibrar diariamente molinos y maquinarias, resguardar el correcto uso de la vajilla e informar fallas tecnicas a la jefatura. Operacion de Caja y Custodia de Valores: Ejecutar los procesos de apertura, operacion y cuadratura diaria de los sistemas de cajas registradoras, terminales de tarjetas, efectivo o transferencias, procesando de manera unificada tanto el consumo de cafeteria como las ventas de la tienda de alimentos y regalos. El trabajador actua como custodio directo de los fondos y documentos recibidos en su turno, debiendo reportar cualquier descuadre de forma inmediata a la administracion. Protocolos de Apertura, Cierre y Mantencion Operativa: Ejecutar los procedimientos de apertura o cierre fisico del local segun el turno asignado (encendido/apagado de equipos, alarmas y accesos). El turno de apertura ejecutara el aseo, orden y sanitizacion profunda inicial de todas las areas (salon, barra y gondolas) junto con la instalacion de carteles y pizarras publicas; mientras que el de cierre resguardara dichos elementos, dejando el local impecable y vaciado de residuos. Asimismo, durante toda la jornada, el trabajador debera monitorear y mantener activamente la limpieza y orden de los espacios, quedando facultado y obligado a realizar tareas inmediatas de aseo, ordenamiento o sanitizacion de emergencia ante cualquier contingencia, derrame o requerimiento sanitario que se presente en el momento. Control de Stock, Exhibicion y Rotacion: Recibir, registrar y verificar la calidad, temperatura y volumen de los insumos de proveedores. Es responsabilidad del trabajador mantener el stock de barra y la reposicion constante de las gondolas y vitrinas de la tienda, asegurando una exhibicion limpia y atractiva. Debera realizar los inventarios periodicos solicitados y controlar rigurosamente las fechas de caducidad aplicando el sistema de rotacion FIFO / PEPS (Primero en entrar, primero en salir) para evitar mermas de stock.",
  cocina: "El trabajador se compromete a desempeniar las funciones de Personal de Cocina y Produccion, atendiendo de manera unificada las lineas de salon, despacho mayorista y servicios de catering o eventos: Produccion Gastronomica e Insumos: Elaborar desde cero masas (pan, pizzas, waffles), pasteleria, almuerzos, postres y platos de la carta, cumpliendo rigurosamente los gramajes y fichas tecnicas de la empresa. Tendra la obligacion de confeccionar y documentar las recetas para el recetario institucional cuando una preparacion no este estandarizada. Asimismo, sera responsable de recibir y verificar tecnicamente la calidad, temperatura y estado de las materias primas entregadas por proveedores. Ante requerimientos de la operacion o de la jefatura, estara facultado y obligado a apoyar activamente en la barra y en el montaje de postres. Control de Produccion, Logistica e Higiene: Monitorear stock, reportar mermas y aplicar estrictamente la rotacion FIFO / PEPS. Es obligacion estricta del trabajador registrar de forma diaria el volumen de produccion ejecutado, informar oportunamente a la jefatura la disponibilidad de stock terminado para despacho, y revisar constantemente que los productos elaborados y almacenados no se encuentren caducados. Debera fraccionar, porcionar, rotular y embalar tecnicamente los productos destinados a cafeteria, clientes externos o eventos. Mantendra la limpieza profunda y sanitizacion de su estacion, mesones, artefactos y maquinarias (incluyendo el lavado constante de utensilios, platos sucios, herramientas y guardado de loza), asegurando el uso eficiente y apagado seguro de los servicios y equipos al cierre del turno.",
  operario: "El trabajador se compromete a desempeniar las funciones de Asistente Multifuncional de Tienda y Alimentos, actuando como el motor operativo responsable de la continuidad, resguardo comercial y salud general del establecimiento: Operacion de Apertura, Cierre y Salud del Local: Ejecutar los procesos de apertura fisica o cierre seguro del local segun el turno asignado. Es obligacion directa del trabajador el encendido oportuno, supervision continua y apagado seguro de toda la maquinaria (maquina de cafe, microondas, freidoras, vitrinas y refrigeradores). Asimismo, velara permanentemente por la salud y correcta presentacion del local, asegurando un ambiente limpio, ordenado, seguro y con la senaletica de precios siempre actualizada. Preparara y servira cafes o bebidas, manteniendo las vitrinas de alimentos atractivas e higienizadas. Gestion de Caja, Custodia de Efectivo y Pedidos: Atender cordialmente al publico y operar los sistemas de venta. El trabajador sera responsable directo de la custodia estricta del dinero en efectivo recaudado durante su jornada, debiendo guardarlo y asegurar su resguardo en los lugares definidos por la empresa, ejecutando una cuadratura transparente de caja al termino de su turno. Monitoreara constantemente el inventario en exhibicion, gestionando y alertando de forma anticipada los pedidos e insumos necesarios para evitar quiebres de stock. Repondra mercancias en gondolas aplicando el sistema FIFO / PEPS y mantendra la sanitizacion profunda de su area de trabajo."
};
const RIOSH_TABLA_ATRASOS = "Tabla multas segun RIOSH Titulo XVII Art.63: 5 a 15 min = 20% del maximo; 15 a 25 min = 50% del maximo; 25 a 45 min = 80% del maximo; mas de 45 min = 100% del maximo. El maximo aplicable equivale al 25% de la remuneracion diaria.";

// ── Helpers ───────────────────────────────────────────────────────────────────
function fmtF(s) {
  if (!s) return "";
  try {
    const d = new Date(s + "T12:00:00");
    const m = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
    return d.getDate() + " de " + m[d.getMonth()] + " de " + d.getFullYear();
  } catch(e) { return s; }
}
function numL(n) {
  const x = parseInt(n); if (isNaN(x)) return "";
  const u = ["","un","dos","tres","cuatro","cinco","seis","siete","ocho","nueve","diez","once","doce","trece","catorce","quince","dieciseis","diecisiete","dieciocho","diecinueve","veinte","veintiun","veintidos","veintitres","veinticuatro","veinticinco","veintiseis","veintisiete","veintiocho","veintinueve"];
  const d2 = ["","","veinte","treinta","cuarenta","cincuenta","sesenta","setenta","ochenta","noventa"];
  const c = ["","ciento","doscientos","trescientos","cuatrocientos","quinientos","seiscientos","setecientos","ochocientos","novecientos"];
  if (x===0) return "cero"; if (x<30) return u[x];
  if (x<100) { const dd=Math.floor(x/10),uu=x%10; return uu===0?d2[dd]:d2[dd]+" y "+u[uu]; }
  if (x===100) return "cien";
  if (x<1000) { const cc=Math.floor(x/100),r=x%100; return c[cc]+(r?" "+numL(r):""); }
  if (x<1000000) { const mm=Math.floor(x/1000),r=x%1000; return (mm===1?"mil":numL(mm)+" mil")+(r?" "+numL(r):""); }
  return x.toLocaleString("es-CL");
}
function calcBreak(e, s, col) {
  if (!e||!s) return "00:00";
  const ep=(e+":0").split(":").map(Number), sp=(s+":0").split(":").map(Number);
  const eM=ep[0]*60+(ep[1]||0), sM=sp[0]*60+(sp[1]||0);
  if (sM<=eM) return "00:00";
  const mid=Math.round((eM+sM)/2)-Math.floor(col/2);
  return String(Math.floor(mid/60)).padStart(2,"0")+":"+String(mid%60).padStart(2,"0");
}
function fmtRut(rut) {
  // Limpia todo excepto números y k/K
  let clean = rut.replace(/[^0-9kK]/g, "");
  if (clean.length < 2) return clean;
  const dv = clean.slice(-1);
  const num = clean.slice(0, -1);
  // Agrega puntos cada 3 dígitos
  let formatted = "";
  for (let i = 0; i < num.length; i++) {
    if (i > 0 && (num.length - i) % 3 === 0) formatted += ".";
    formatted += num[i];
  }
  return formatted + "-" + dv;
}

function getLugar(emp) {
  return emp==="Malvarrosa SpA" ? "Casa Turquesa, Av. Ortuzar 250, Ñuñoa, Santiago" : "Av. Ossa 1624, Ñuñoa, Santiago";
}

// ── DOCX helpers ──────────────────────────────────────────────────────────────
const FONT = "Times New Roman";
const SZ = 22; const SZ_T = 32;
const brd = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top:brd, bottom:brd, left:brd, right:brd };
const noBorders = { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} };

function P(runs, align, spacingAfter) {
  return new Paragraph({ children: runs, alignment: align||AlignmentType.JUSTIFIED, spacing: { after: spacingAfter!==undefined?spacingAfter:200 } });
}
function T(text, opts={}) {
  return new TextRun({ text, font:FONT, size:SZ, bold:opts.bold||false, underline:opts.underline?{type:UnderlineType.SINGLE}:undefined });
}
function TT(text, opts={}) {
  return new TextRun({ text, font:FONT, size:SZ_T, bold:opts.bold||false, underline:opts.underline?{type:UnderlineType.SINGLE}:undefined });
}
function titulo(text) {
  return P([TT(text, {bold:true, underline:true})], AlignmentType.CENTER, 120);
}
function art(label, content) {
  return P([T(label+" ", {bold:true}), T(content)], AlignmentType.JUSTIFIED, 80);
}
function makeFooter(empNombre, trabNombre) {
  return new Footer({ children: [
    new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({text:"Contrato de trabajo pagina ", font:FONT, size:16}),
      new TextRun({children:[PageNumber.CURRENT], font:FONT, size:16}),
      new TextRun({text:"/", font:FONT, size:16}),
      new TextRun({children:[PageNumber.TOTAL_PAGES], font:FONT, size:16}),
      new TextRun({text:"  —  "+empNombre+" y "+trabNombre, font:FONT, size:16}),
    ]})
  ]});
}
function cellT(text, width, opts={}) {
  return new TableCell({
    borders, verticalAlign: VerticalAlign.CENTER,
    width: { size: width, type: WidthType.DXA },
    shading: opts.header ? { fill:"D9D9D9", type:ShadingType.CLEAR } : opts.section ? { fill:"BFBFBF", type:ShadingType.CLEAR } : undefined,
    margins: { top:50, bottom:50, left:80, right:80 },
    children: [new Paragraph({ alignment:AlignmentType.CENTER, children:[new TextRun({text, font:FONT, size:16, bold:opts.bold||opts.header||opts.section||false})] })]
  });
}

// ── Generar Contrato PF ───────────────────────────────────────────────────────
async function generarContratoPF(data) {
  const emp = EMP[data.empresa];
  const dom = [data.calle, data.num, data.comuna, data.ciudad].filter(Boolean).join(", ");
  const hoy = fmtF(new Date().toISOString().split("T")[0]);
  const colW = [1800,1000,1000,1000,1000,1000,1000,1000];
  const dias = ["Seccion","Lunes","Martes","Miercoles","Jueves","Viernes","Sabado","Domingo"];

  function makeRow(vals, isHeader, isSection) {
    return new TableRow({ children: vals.map((v,i) => new TableCell({
      borders, verticalAlign:VerticalAlign.CENTER,
      width:{ size:colW[i], type:WidthType.DXA },
      shading: isHeader?{fill:"D9D9D9",type:ShadingType.CLEAR}:isSection?{fill:"BFBFBF",type:ShadingType.CLEAR}:undefined,
      margins:{top:50,bottom:50,left:80,right:80},
      children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:v,font:FONT,size:16,bold:isHeader||isSection||(i===0&&!isHeader)})]  })]
    }))});
  }
  function sectionRow(label) {
    return new TableRow({ children:[new TableCell({
      borders, columnSpan:8,
      width:{size:colW.reduce((a,b)=>a+b,0), type:WidthType.DXA},
      shading:{fill:"BFBFBF",type:ShadingType.CLEAR},
      margins:{top:50,bottom:50,left:80,right:80},
      children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:label,font:FONT,size:16,bold:true})]})]
    })]});
  }

  const turnoRows = [];
  const entradas = DIAS_F.map((d,i) => { const r=data.turno[i]; return r.on&&r.e ? r.e : "00:00"; });
  const salidas1 = DIAS_F.map((d,i) => { const r=data.turno[i]; if (!r.on||!r.e||!r.s) return "00:00"; return calcBreak(r.e,r.s,data.colacion); });
  const breaks  = DIAS_F.map((d,i) => { const r=data.turno[i]; if (!r.on) return "00:00"; return String(Math.floor(data.colacion/60)).padStart(2,"0")+":"+String(data.colacion%60).padStart(2,"0"); });
  const entradas2= DIAS_F.map((d,i) => { const r=data.turno[i]; if (!r.on||!r.e||!r.s) return "00:00"; const [s1h,s1m]=salidas1[i].split(":").map(Number); const t=s1h*60+s1m+data.colacion; return String(Math.floor(t/60)).padStart(2,"0")+":"+String(t%60).padStart(2,"0"); });
  const salidas2 = DIAS_F.map((d,i) => { const r=data.turno[i]; return r.on&&r.s ? r.s : "00:00"; });

  const tablaJornada = new Table({
    width:{size:colW.reduce((a,b)=>a+b,0),type:WidthType.DXA}, columnWidths:colW,
    rows:[
      makeRow(dias,true,false),
      sectionRow("Primera mitad jornada"),
      makeRow(["Entrada",...entradas],false,false),
      makeRow(["Salida",...salidas1],false,false),
      sectionRow("Tiempo Descanso"),
      makeRow(["Libre disposicion",...breaks],false,false),
      sectionRow("Segunda mitad jornada"),
      makeRow(["Entrada",...entradas2],false,false),
      makeRow(["Salida",...salidas2],false,false),
    ]
  });

  const tablaFirmas = new Table({
    width:{size:9360,type:WidthType.DXA}, columnWidths:[4680,4680],
    rows:[
      new TableRow({children:[
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:300,bottom:80,left:120,right:120},children:[P([T("_________________________________")],AlignmentType.CENTER,0)]}),
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:300,bottom:80,left:120,right:120},children:[P([T("_________________________________")],AlignmentType.CENTER,0)]}),
      ]}),
      new TableRow({children:[
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:0,bottom:0,left:120,right:120},children:[
          P([T(emp.nombre,{bold:true})],AlignmentType.CENTER,0),
          P([T(emp.rut)],AlignmentType.CENTER,0),
          P([T("Empleador")],AlignmentType.CENTER,0),
        ]}),
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:0,bottom:0,left:120,right:120},children:[
          P([T(data.nombre+" "+data.apellido,{bold:true})],AlignmentType.CENTER,0),
          P([T(data.rut)],AlignmentType.CENTER,0),
          P([T("Trabajador")],AlignmentType.CENTER,0),
        ]}),
      ]}),
    ]
  });

  const doc = new Document({ sections:[{
    properties:{ page:{ size:{width:11906,height:16838}, margin:{top:1134,right:1134,bottom:1134,left:1134} } },
    footers:{ default: makeFooter(emp.nombre, data.apellido+" "+data.nombre) },
    children:[
      titulo("CONTRATO DE TRABAJO"),
      P([]),
      P([
        T("En Ñuñoa, a "), T(hoy,{bold:true}), T(", entre "), T(emp.nombre,{bold:true}),
        T(", R.U.T. "+emp.rut+", con domicilio en "+emp.dom+", representado legalmente por don(a) "),
        T(emp.rep,{bold:true}), T(", cedula de identidad "+emp.repRut+", en adelante el \"Empleador\", y don(a) "),
        T(data.nombre+" "+data.apellido,{bold:true}),
        T(", de nacionalidad "+data.nac+", nacido(a) el "+fmtF(data.nacF)+", domiciliado(a) en "+dom+", RUT "+data.rut+", estado civil "+data.civil+", email: "+data.email+", telefono: "+data.tel+", cuenta bancaria "+data.banco+", "+data.tipoCuenta+" N "+data.cuenta+", en adelante tambien denominado \"TRABAJADOR\", se ha convenido el siguiente Contrato Individual de Trabajo:"),
      ], AlignmentType.JUSTIFIED, 120),
      art("PRIMERO :", data.funciones+" El trabajador se desempenara en "+getLugar(data.empresa)+", pudiendo ser trasladado a otro Departamento o Seccion de la Oficina Principal o de cualquiera de las Agencias del Empleador, a condicion que se trate de labores similares, en la misma ciudad, y sin que ello importe menoscabo para el trabajador, todo ello sujeto a las necesidades operativas de la Empresa."),
      P([T("SEGUNDO : ",{bold:true}), T("JORNADA DE TRABAJO. El trabajador cumplira una jornada semanal ordinaria de "), T(data.horas+" horas",{bold:true}), T(", de acuerdo a la siguiente distribucion semanal:")], AlignmentType.JUSTIFIED, 60),
      tablaJornada,
      P([T("La asignacion de este turno sera fija durante toda la vigencia del contrato, salvo acuerdo expreso entre las partes mediante anexo de contrato. El tiempo destinado a la colacion es de "), T(data.colacion+" minutos",{bold:true}), T(" y es de cargo del trabajador.")], AlignmentType.JUSTIFIED, 80),
      art("TERCERO :", "Cuando por necesidades de funcionamiento de la Empresa sea necesario pactar trabajo en tiempo extraordinario, el Empleado que lo acuerde se obligara a cumplir el horario que al efecto determine la Empleadora, dentro de los limites legales. Dicho acuerdo constara por escrito y se firmara por ambas partes, previamente a la realizacion del trabajo. El tiempo extraordinario se remunerara con el recargo legal del 50% sobre el sueldo convenido para la jornada ordinaria y se liquidara y pagara conjuntamente con la remuneracion del respectivo periodo."),
      art("CUARTO :", "El empleado percibira un sueldo de $"+parseInt(data.sueldo).toLocaleString("es-CL")+" ("+numL(data.sueldo)+" pesos) mensuales, pagaderos por meses vencidos. Las deducciones que la Empleadora podra practicar a las remuneraciones son todas aquellas que dispone el articulo 58 del Codigo del Trabajo."),
      art("QUINTO :", "El trabajador acepta y autoriza al Empleador para que haga las deducciones que establecen las leyes vigentes y para que le descuente las horas y el tiempo no trabajado debido a atrasos, inasistencias o permisos y, ademas, la rebaja del monto de las multas establecidas en el Reglamento Interno de Orden, Higiene y Seguridad, en caso que procedieren."),
      art("SEXTO :", "La Empresa se obliga a pagar al empleado una gratificacion anual equivalente al 25% del total de las remuneraciones mensuales que este hubiere percibido en el ano, con tope de 4,75 Ingresos Minimos Mensuales. Esta gratificacion se calculara, liquidara y anticipara mensualmente en forma coetanea con la remuneracion del mes respectivo, siendo cada abono equivalente a la doceava parte de la gratificacion anual. La gratificacion convenida sustituye a la que resulte de la aplicacion de los articulos 47 y siguientes del Codigo del Trabajo, siempre que esta ultima fuere igual o inferior a aquella. Los valores anticipados mensualmente se reajustaran en conformidad con lo dispuesto en el articulo 63 del Codigo del Trabajo."),
      art("SEPTIMO :", "El trabajador se obliga y compromete expresamente a cumplir las instrucciones que le sean impartidas por su jefe inmediato o por la Gerencia de la empresa y a acatar en todas sus partes las disposiciones establecidas en el Reglamento de Orden, Higiene y Seguridad, las que declara conocer y que se consideran parte integrante del presente contrato, reglamento del cual el trabajador recibe un ejemplar en este acto."),
      art("OCTAVO :", "Las partes acuerdan que los atrasos reiterados, sin causa justificada, se consideraran incumplimiento grave de las obligaciones que impone el presente contrato y daran lugar a la aplicacion de la caducidad del contrato, contemplada en el Art. 160 N7 del Codigo del Trabajo. Se entendera por atraso reiterado el llegar despues de la hora de ingreso durante 7 dias maximos, seguidos o no, en cada mes calendario."),
      art("NOVENO :", "El presente contrato comenzara el "+fmtF(data.inicio)+" y tendra duracion hasta el dia "+fmtF(data.termino)+", y cualquiera de las partes podra ponerle termino."),
      art("DECIMO :", "Para todas las cuestiones a que eventualmente pueda dar origen este contrato, las partes fijan domicilio en la comuna de Ñuñoa."),
      art("DECIMO PRIMERO :", "Todas las recetas, preparaciones, fichas tecnicas, procesos productivos, formulas y creaciones desarrolladas por el trabajador en el ejercicio de sus funciones seran de propiedad exclusiva del Empleador. El trabajador se obliga a no divulgar, reproducir ni utilizar dicha informacion fuera de la empresa, ni durante ni despues de la vigencia del presente contrato."),
      art("DECIMO SEGUNDO :", "Se deja constancia que el Empleado ingreso al servicio de la Empresa con fecha "+fmtF(data.inicio)+". El presente contrato se firma en dos ejemplares del mismo tenor y fecha, quedando uno en poder de cada parte."),
      P([T("")], AlignmentType.JUSTIFIED, 200),
      tablaFirmas,
    ]
  }]});

  const blob = await Packer.toBlob(doc);
  saveAs(blob, "Contrato_"+data.apellido+"_"+data.nombre+"_"+data.inicio+".docx");
}

// ── Generar Anexo ─────────────────────────────────────────────────────────────
async function generarAnexo(data) {
  const emp = EMP[data.empresa];
  const hoy = fmtF(new Date().toISOString().split("T")[0]);
  const tiposMap = { plazo:"Prorroga de plazo fijo", indefinido:"Conversion a contrato indefinido (Art. 159 N4 CT)", sueldo:"Cambio de sueldo", cargo:"Cambio de cargo", horario:"Cambio de horario", otro:"Modificacion contractual" };

  const tablaFirmas = new Table({
    width:{size:9360,type:WidthType.DXA}, columnWidths:[4680,4680],
    rows:[
      new TableRow({children:[
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:300,bottom:80,left:120,right:120},children:[P([T("_________________________________")],AlignmentType.CENTER,0)]}),
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:300,bottom:80,left:120,right:120},children:[P([T("_________________________________")],AlignmentType.CENTER,0)]}),
      ]}),
      new TableRow({children:[
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:0,bottom:0,left:120,right:120},children:[P([T(emp.nombre,{bold:true})],AlignmentType.CENTER,0),P([T(emp.rut)],AlignmentType.CENTER,0),P([T("Empleador")],AlignmentType.CENTER,0)]}),
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:0,bottom:0,left:120,right:120},children:[P([T(data.tNombre,{bold:true})],AlignmentType.CENTER,0),P([T(data.tRut)],AlignmentType.CENTER,0),P([T("Trabajador")],AlignmentType.CENTER,0)]}),
      ]}),
    ]
  });

  const doc = new Document({ sections:[{
    properties:{ page:{ size:{width:11906,height:16838}, margin:{top:1134,right:1134,bottom:1134,left:1134} } },
    children:[
      titulo("ANEXO DE CONTRATO DE TRABAJO"),
      P([]),
      P([T("En Ñuñoa, a "), T(hoy,{bold:true}), T(", entre "), T(emp.nombre,{bold:true}), T(", R.U.T. "+emp.rut+", representado por "+emp.rep+", cedula "+emp.repRut+", en adelante el Empleador, y don(a) "), T(data.tNombre,{bold:true}), T(", RUT "+data.tRut+", cargo: "+data.tCargo+", en adelante el Trabajador, se ha convenido el siguiente Anexo al Contrato de Trabajo:")], AlignmentType.JUSTIFIED, 120),
      art("ARTICULO PRIMERO :", "Las partes acuerdan modificar el contrato de trabajo en los siguientes terminos: "+tiposMap[data.anTipo]+". "+data.anDesc),
      art("ARTICULO SEGUNDO :", "Nueva condicion contractual: "+data.anCond+". Esta modificacion rige a partir de "+fmtF(data.anFecha)+"."),
      art("ARTICULO TERCERO :", "En todo lo demas, las clausulas del contrato original permanecen vigentes e inalteradas, siendo este Anexo parte integrante del mismo."),
      art("ARTICULO CUARTO :", "El presente Anexo se firma en dos ejemplares del mismo tenor y fecha, quedando uno en poder de cada contratante."),
      P([T("")], AlignmentType.JUSTIFIED, 400),
      tablaFirmas,
    ]
  }]});

  const blob = await Packer.toBlob(doc);
  saveAs(blob, "Anexo_"+data.tNombre.replace(" ","_")+"_"+data.anFecha+".docx");
}

// ── Generar Amonestacion Inasistencias ────────────────────────────────────────
async function generarAmonInasistencia(data) {
  const emp = EMP[data.empresa];
  const doc = new Document({ sections:[{
    properties:{ page:{ size:{width:11906,height:16838}, margin:{top:1134,right:1134,bottom:1134,left:1134} } },
    children:[
      titulo("CARTA DE AMONESTACION"),
      P([T("Por Inasistencias Injustificadas",{bold:true})], AlignmentType.CENTER, 200),
      P([T("Ñuñoa, "+fmtF(data.aiFecha))], AlignmentType.JUSTIFIED, 120),
      P([T("Senor(a):")], AlignmentType.JUSTIFIED, 40),
      P([T(data.tNombre,{bold:true})], AlignmentType.JUSTIFIED, 40),
      P([T("Cargo: "+data.tCargo)], AlignmentType.JUSTIFIED, 40),
      P([T("RUT: "+data.tRut)], AlignmentType.JUSTIFIED, 120),
      P([T("Por medio de la presente, "+emp.nombre+", RUT "+emp.rut+", representada por "+emp.rep+", viene en amonestarlo(a) formalmente por las siguientes inasistencias:")], AlignmentType.JUSTIFIED, 80),
      P([T("Fechas de inasistencia: "+data.aiDias,{bold:true})], AlignmentType.JUSTIFIED, 80),
      P([T("Motivo o justificacion entregada por el trabajador: "+(data.aiMotivo||"Sin justificacion")+".")], AlignmentType.JUSTIFIED, 80),
      P([T("Las inasistencias descritas constituyen una infraccion a las obligaciones establecidas en el Titulo VII, Articulo 27 del Reglamento Interno de Orden, Higiene y Seguridad (RIOSH) de la empresa, que dispone expresamente la obligacion de asistencia puntual y la comunicacion de toda inasistencia dentro de las primeras 24 horas.")], AlignmentType.JUSTIFIED, 80),
      P([T(data.aiReinc ? "Dado que usted presenta reincidencia en este tipo de faltas, la presente amonestacion se emite con copia a la Inspeccion del Trabajo correspondiente, conforme a lo establecido en el Titulo XVII, Articulo 63 del RIOSH." : "La presente amonestacion se emite conforme a lo establecido en el Titulo XVII, Articulo 63 del Reglamento Interno de Orden, Higiene y Seguridad de la empresa.")], AlignmentType.JUSTIFIED, 80),
      P([T("Se le hace presente que la reiteracion de este tipo de conductas podra ser considerada incumplimiento grave de las obligaciones que impone el contrato de trabajo, causal de termino contemplada en el Articulo 160 N7 del Codigo del Trabajo.")], AlignmentType.JUSTIFIED, 80),
      P([T("El trabajador tiene derecho a formular sus descargos por escrito dentro del plazo de 3 dias habiles desde la recepcion de esta carta.")], AlignmentType.JUSTIFIED, 200),
      P([T("Acuse de recibo:")], AlignmentType.JUSTIFIED, 300),
      new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[4680,4680], rows:[
        new TableRow({children:[
          new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:0,bottom:0,left:120,right:120},children:[P([T("_________________________________")],AlignmentType.CENTER,0),P([T(emp.rep,{bold:true})],AlignmentType.CENTER,0),P([T("Empleador")],AlignmentType.CENTER,0)]}),
          new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:0,bottom:0,left:120,right:120},children:[P([T("_________________________________")],AlignmentType.CENTER,0),P([T(data.tNombre,{bold:true})],AlignmentType.CENTER,0),P([T("Trabajador / Fecha recepcion")],AlignmentType.CENTER,0)]}),
        ]}),
      ]}),
    ]
  }]});

  const blob = await Packer.toBlob(doc);
  saveAs(blob, "Amonestacion_Inasistencia_"+data.tNombre.replace(" ","_")+".docx");
}

// ── Generar Amonestacion Atrasos ──────────────────────────────────────────────
async function generarAmonAtraso(data) {
  const emp = EMP[data.empresa];
  const doc = new Document({ sections:[{
    properties:{ page:{ size:{width:11906,height:16838}, margin:{top:1134,right:1134,bottom:1134,left:1134} } },
    children:[
      titulo("CARTA DE AMONESTACION"),
      P([T("Por Atrasos Reiterados",{bold:true})], AlignmentType.CENTER, 200),
      P([T("Ñuñoa, "+fmtF(data.aaFecha))], AlignmentType.JUSTIFIED, 120),
      P([T("Senor(a):")], AlignmentType.JUSTIFIED, 40),
      P([T(data.tNombre,{bold:true})], AlignmentType.JUSTIFIED, 40),
      P([T("Cargo: "+data.tCargo)], AlignmentType.JUSTIFIED, 40),
      P([T("RUT: "+data.tRut)], AlignmentType.JUSTIFIED, 120),
      P([T("Por medio de la presente, "+emp.nombre+", RUT "+emp.rut+", representada por "+emp.rep+", viene en amonestarlo(a) formalmente por los siguientes atrasos registrados:")], AlignmentType.JUSTIFIED, 80),
      P([T(data.aaDetalle,{bold:true})], AlignmentType.JUSTIFIED, 80),
      P([T("Los atrasos descritos constituyen una infraccion a las obligaciones establecidas en el Titulo VII, Articulo 27 del RIOSH de la empresa y a lo pactado en el contrato de trabajo. Conforme a la tabla de multas del Titulo XVII, Articulo 63 del RIOSH: "+RIOSH_TABLA_ATRASOS)], AlignmentType.JUSTIFIED, 80),
      P([T(data.aaReinc ? "Dado que usted presenta reincidencia en atrasos, la presente amonestacion se emite con copia a la Inspeccion del Trabajo correspondiente, conforme al Titulo XVII Art. 63 del RIOSH." : "La presente amonestacion se emite conforme al Titulo XVII, Articulo 63 del Reglamento Interno de Orden, Higiene y Seguridad.")], AlignmentType.JUSTIFIED, 80),
      P([T("Se le hace presente que conforme al Articulo 160 N7 del Codigo del Trabajo y lo pactado en su contrato, llegar atrasado 7 o mas dias en un mismo mes calendario sera considerado incumplimiento grave de las obligaciones contractuales, causal suficiente para poner termino al contrato sin derecho a indemnizacion.")], AlignmentType.JUSTIFIED, 80),
      P([T("El trabajador tiene derecho a formular sus descargos por escrito dentro del plazo de 3 dias habiles desde la recepcion de esta carta.")], AlignmentType.JUSTIFIED, 200),
      P([T("Acuse de recibo:")], AlignmentType.JUSTIFIED, 300),
      new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[4680,4680], rows:[
        new TableRow({children:[
          new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:0,bottom:0,left:120,right:120},children:[P([T("_________________________________")],AlignmentType.CENTER,0),P([T(emp.rep,{bold:true})],AlignmentType.CENTER,0),P([T("Empleador")],AlignmentType.CENTER,0)]}),
          new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:0,bottom:0,left:120,right:120},children:[P([T("_________________________________")],AlignmentType.CENTER,0),P([T(data.tNombre,{bold:true})],AlignmentType.CENTER,0),P([T("Trabajador / Fecha recepcion")],AlignmentType.CENTER,0)]}),
        ]}),
      ]}),
    ]
  }]});

  const blob = await Packer.toBlob(doc);
  saveAs(blob, "Amonestacion_Atraso_"+data.tNombre.replace(" ","_")+".docx");
}

// ── Generar No Renovacion ─────────────────────────────────────────────────────
async function generarNoRenovacion(data) {
  const emp = EMP[data.empresa];

  // Nombre empresa en formato titulo
  const empNombreT = emp.nombre === "MALVARROSA SPA" ? "Malvarrosa Spa" : "Venta de Alimentos Eric Felipe Hansen Figueroa Eirl";

  // Bordes invisibles para encabezado
  const nb = { style:BorderStyle.NONE, size:0, color:"FFFFFF" };
  const nbAll = { top:nb, bottom:nb, left:nb, right:nb };
  function encRow(label, valor) {
    return new TableRow({ children:[
      new TableCell({ borders:nbAll, width:{size:2000,type:WidthType.DXA}, margins:{top:30,bottom:30,left:0,right:0},
        children:[new Paragraph({children:[new TextRun({text:label, font:FONT, size:18})]})]
      }),
      new TableCell({ borders:nbAll, width:{size:7360,type:WidthType.DXA}, margins:{top:30,bottom:30,left:0,right:0},
        children:[new Paragraph({children:[new TextRun({text:valor, font:FONT, size:18})]})]
      }),
    ]});
  }

  const tablaEnc = new Table({
    width:{size:9360,type:WidthType.DXA}, columnWidths:[2000,7360],
    borders:{top:nb,bottom:nb,left:nb,right:nb,insideH:nb,insideV:nb},
    rows:[
      encRow("Razón Social", empNombreT),
      encRow("R.U.T.", emp.rut),
      encRow("Dirección", emp.dir),
      encRow("Comuna", emp.comuna),
      encRow("Giro", emp.giro),
    ]
  });

  const lineaSep = new Paragraph({
    spacing:{after:200,before:80},
    border:{bottom:{style:BorderStyle.SINGLE,size:12,color:"185FA5"}},
    children:[]
  });

  const tituloNR = new Paragraph({
    alignment:AlignmentType.CENTER,
    spacing:{after:280},
    children:[new TextRun({
      text:"AVISO NO RENOVACIÓN CONTRATO PLAZO FIJO",
      font:FONT, size:22, bold:true, color:"185FA5",
      underline:{type:UnderlineType.SINGLE}
    })]
  });

  const tablaFirmas = new Table({
    width:{size:9360,type:WidthType.DXA}, columnWidths:[4680,4680],
    rows:[
      new TableRow({children:[
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:240,bottom:40,left:120,right:120},
          children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"_________________________________",font:FONT,size:SZ})]})]
        }),
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:240,bottom:40,left:120,right:120},
          children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"_________________________________",font:FONT,size:SZ})]})]
        }),
      ]}),
      new TableRow({children:[
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:0,bottom:0,left:120,right:120},children:[
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Representante Legal",font:FONT,size:SZ})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:empNombreT.toUpperCase(),font:FONT,size:SZ,bold:true})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:emp.rut,font:FONT,size:SZ})]}),
        ]}),
        new TableCell({borders:noBorders,width:{size:4680,type:WidthType.DXA},margins:{top:0,bottom:0,left:120,right:120},children:[
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Recibí Notificación",font:FONT,size:SZ})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:data.tNombre,font:FONT,size:SZ,bold:true})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:data.tRut,font:FONT,size:SZ})]}),
        ]}),
      ]}),
    ]
  });

  const doc = new Document({ sections:[{
    properties:{ page:{ size:{width:11906,height:16838}, margin:{top:1134,right:1134,bottom:1134,left:1134} } },
    children:[
      tablaEnc,
      lineaSep,
      tituloNR,
      new Paragraph({spacing:{after:160},children:[new TextRun({text:"Ñuñoa, "+fmtF(data.nrFechaDoc),font:FONT,size:SZ})]}),
      new Paragraph({spacing:{after:40},children:[new TextRun({text:"Señor (a):"+data.tNombre,font:FONT,size:SZ})]}),
      new Paragraph({spacing:{after:40},children:[new TextRun({text:"Rut : "+data.tRut,font:FONT,size:SZ})]}),
      new Paragraph({spacing:{after:200},children:[new TextRun({text:"P R E S E N T E",font:FONT,size:SZ,bold:true})]}),
      new Paragraph({alignment:AlignmentType.JUSTIFIED,spacing:{after:160},children:[
        new TextRun({text:"Nos permitimos comunicar que, con esta fecha, ",font:FONT,size:SZ}),
        new TextRun({text:fmtF(data.nrFechaTermino),font:FONT,size:SZ}),
        new TextRun({text:", se ha resuelto no renovar su contrato de trabajo que lo vincula con la empresa, por la causal del artículo ",font:FONT,size:SZ}),
        new TextRun({text:"Art. 159 Número 4",font:FONT,size:SZ,bold:true}),
        new TextRun({text:", del Código del Trabajo, esto es Vencimiento del plazo convenido en el contrato. La duración del contrato de plazo fijo no podrá exceder de un año.",font:FONT,size:SZ}),
      ]}),
      new Paragraph({alignment:AlignmentType.JUSTIFIED,spacing:{after:160},children:[new TextRun({text:"Informo que sus cotizaciones previsionales se encuentran al día. Además, le adjuntamos certificado de cotizaciones de las entidades de previsión a las que se encuentra afiliado, que dan cuenta que las cotizaciones previsionales, del período trabajado, se encuentran pagadas.",font:FONT,size:SZ})]}),
      new Paragraph({alignment:AlignmentType.JUSTIFIED,spacing:{after:160},children:[new TextRun({text:"Asimismo se ha escogido el metodo Electrónico para pago de los diferentes conceptos del finiquito, si usted no desea que esto sea realice de esta manera y sea presencial, debe informar su decisión con antelación a la empresa, para que se coordine el pago según su preferencia ante la Dirección del Trabajo o un Notario Público. De no hacer ninguna indicación se entenderá que el pago será electrónico.",font:FONT,size:SZ})]}),
      new Paragraph({alignment:AlignmentType.JUSTIFIED,spacing:{after:400},children:[new TextRun({text:"Usted podrá al momento de aceptar la subscripcion del finiquito realizar una reserva de derechos para accionar judicialmente contra el empleador como consigna el codigo del trabajo art. 177.",font:FONT,size:SZ})]}),
      tablaFirmas,
      new Paragraph({spacing:{after:160},children:[]}),
      new Paragraph({spacing:{after:40},children:[new TextRun({text:"C.C. Jefatura Directa",font:FONT,size:SZ})]}),
      new Paragraph({spacing:{after:40},children:[new TextRun({text:"C.C. "+data.tNombre,font:FONT,size:SZ})]}),
      new Paragraph({spacing:{after:40},children:[new TextRun({text:"C.C. Dirección del Trabajo",font:FONT,size:SZ})]}),
    ]
  }]});

  const blob = await Packer.toBlob(doc);
  saveAs(blob, "No_Renovacion_"+data.tNombre.replace(/ /g,"_")+".docx");
}

// ── Estilos ───────────────────────────────────────────────────────────────────
const S = {
  inp:  { width:"100%", padding:"10px", border:"1px solid #e0e0db", borderRadius:8, fontSize:16, color:"#1a1a1a", outline:"none", boxSizing:"border-box", WebkitAppearance:"none" },
  sel:  { width:"100%", padding:"10px", border:"1px solid #e0e0db", borderRadius:8, fontSize:16, color:"#1a1a1a", background:"white", boxSizing:"border-box", WebkitAppearance:"none" },
  ta:   { width:"100%", padding:"10px", border:"1px solid #e0e0db", borderRadius:8, fontSize:14, color:"#1a1a1a", resize:"vertical", minHeight:80, fontFamily:"inherit", boxSizing:"border-box" },
  lbl:  { fontSize:13, color:"#555", display:"block", marginBottom:5, marginTop:14 },
  card: { background:"white", border:"1px solid #e8e8e5", borderRadius:12, padding:"1.1rem", marginBottom:"1rem" },
  btnP: { background:"#185FA5", color:"white", border:"none", width:"100%", padding:14, borderRadius:8, fontSize:16, cursor:"pointer", fontWeight:700, marginTop:10 },
  r2:   { display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 },
  aE:   { padding:"10px 14px", borderRadius:8, fontSize:13, marginTop:8, background:"#FCEBEB", color:"#A32D2D" },
  aS:   { padding:"10px 14px", borderRadius:8, fontSize:13, marginTop:8, background:"#E1F5EE", color:"#0F6E56" },
};
const pill = (a) => ({ padding:"8px 16px", borderRadius:20, fontSize:13, cursor:"pointer", border:a?"1px solid #185FA5":"1px solid #e0e0db", background:a?"#185FA5":"white", color:a?"white":"#1a1a1a", fontWeight:a?600:400 });
const dtcSt = (a) => ({ padding:"10px 6px", border:a?"2px solid #185FA5":"1px solid #e0e0db", borderRadius:10, cursor:"pointer", textAlign:"center", background:a?"#EBF3FF":"white" });

function LBL({children, first}) { return <label style={{...S.lbl, marginTop:first?0:14}}>{children}</label>; }
function FormTrabSimple({nombre, onNombre, rut, onRut, cargo, onCargo, trabSheet, loadingTrabs, onSelectTrab}) {
  const [modo, setModo] = useState("lista"); // "lista" | "manual"
  return (
    <div style={S.card}>
      <div style={{fontSize:15,fontWeight:700,marginBottom:10}}>Trabajador</div>
      <div style={{display:"flex",gap:8,marginBottom:14}}>
        <button style={pill(modo==="lista")} onClick={()=>setModo("lista")}>📋 Seleccionar</button>
        <button style={pill(modo==="manual")} onClick={()=>setModo("manual")}>✏️ Ingresar manual</button>
      </div>
      {modo==="lista" && (
        loadingTrabs
          ? <div style={{color:"#888",fontSize:13,padding:"10px 0"}}>Cargando trabajadores...</div>
          : trabSheet.length===0
            ? <div style={{color:"#888",fontSize:13,padding:"10px 0"}}>No hay trabajadores activos en esta empresa. Genera primero un contrato.</div>
            : <div>
                {trabSheet.map(t=>(
                  <div key={t.id} onClick={()=>{ onSelectTrab(t); setModo("datos"); }}
                    style={{display:"flex",alignItems:"center",padding:"10px 0",borderBottom:"1px solid #e8e8e5",gap:10,cursor:"pointer"}}>
                    <div style={{width:34,height:34,borderRadius:"50%",background:"#EBF3FF",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,color:"#185FA5",flexShrink:0}}>{(t.nombre||"?").charAt(0)}</div>
                    <div style={{flex:1}}>
                      <div style={{fontSize:13,fontWeight:600}}>{t.nombre} {t.apellido}</div>
                      <div style={{fontSize:11,color:"#888"}}>{t.cargo||"-"} · {t.rut}</div>
                    </div>
                    <span style={{color:"#185FA5",fontSize:18}}>›</span>
                  </div>
                ))}
              </div>
      )}
      {(modo==="manual"||modo==="datos") && (
        <div>
          {modo==="datos" && <div style={{background:"#E1F5EE",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#0F6E56",marginBottom:10}}>✓ Datos cargados — puedes editar si es necesario</div>}
          <div style={S.r2}>
            <div><LBL first>Nombre y apellido</LBL><input style={S.inp} value={nombre} onChange={e=>onNombre(e.target.value)} placeholder="Juan Perez"/></div>
            <div><LBL first>RUT</LBL><input style={S.inp} value={rut} onChange={e=>onRut(fmtRut(e.target.value))} placeholder="12.345.678-9"/></div>
          </div>
          <LBL>Cargo</LBL><input style={S.inp} value={cargo} onChange={e=>onCargo(e.target.value)} placeholder="Garzon - Barista"/>
        </div>
      )}
    </div>
  );
}
function FormPF({value, onChange}) {
  const u = useCallback((k,val) => onChange(p=>({...p,[k]:val})), [onChange]);
  return (
    <div style={S.card}>
      <div style={{fontSize:15,fontWeight:700,marginBottom:14}}>Datos del trabajador</div>
      <div style={S.r2}>
        <div><LBL first>Nombre</LBL><input style={S.inp} value={value.nombre} onChange={e=>u("nombre",e.target.value)} placeholder="Nombre"/></div>
        <div><LBL first>Apellido(s)</LBL><input style={S.inp} value={value.apellido} onChange={e=>u("apellido",e.target.value)} placeholder="Apellidos"/></div>
      </div>
      <div style={S.r2}>
        <div><LBL>RUT</LBL><input style={S.inp} value={value.rut} onChange={e=>u("rut",fmtRut(e.target.value))} placeholder="12.345.678-9"/></div>
        <div><LBL>Nacionalidad</LBL><input style={S.inp} value={value.nac} onChange={e=>u("nac",e.target.value)}/></div>
      </div>
      <div style={S.r2}>
        <div><LBL>Fecha nac.</LBL><input style={S.inp} type="date" value={value.nacF} onChange={e=>u("nacF",e.target.value)}/></div>
        <div><LBL>Estado civil</LBL>
          <select style={S.sel} value={value.civil} onChange={e=>u("civil",e.target.value)}>
            {["Soltero(a)","Casado(a)","Divorciado(a)","Viudo(a)"].map(o=><option key={o}>{o}</option>)}
          </select>
        </div>
      </div>
      <LBL>Calle</LBL><input style={S.inp} value={value.calle} onChange={e=>u("calle",e.target.value)} placeholder="Av. Ejemplo"/>
      <div style={S.r2}>
        <div><LBL>Numero</LBL><input style={S.inp} value={value.num} onChange={e=>u("num",e.target.value)} placeholder="123"/></div>
        <div><LBL>Comuna</LBL><input style={S.inp} value={value.comuna} onChange={e=>u("comuna",e.target.value)} placeholder="Ñuñoa"/></div>
      </div>
      <LBL>Ciudad</LBL><input style={S.inp} value={value.ciudad} onChange={e=>u("ciudad",e.target.value)} placeholder="Santiago"/>
      <div style={S.r2}>
        <div><LBL>Email</LBL><input style={S.inp} value={value.email} onChange={e=>u("email",e.target.value)} placeholder="correo@gmail.com"/></div>
        <div><LBL>Telefono</LBL><input style={S.inp} value={value.tel} onChange={e=>u("tel",e.target.value)} placeholder="9XXXXXXXX"/></div>
      </div>
      <div style={S.r2}>
        <div><LBL>Banco</LBL><input style={S.inp} value={value.banco} onChange={e=>u("banco",e.target.value)} placeholder="Banco Estado"/></div>
        <div><LBL>N cuenta</LBL><input style={S.inp} value={value.cuenta} onChange={e=>u("cuenta",e.target.value)} placeholder="0000000"/></div>
      </div>
      <LBL>Tipo cuenta</LBL>
      <select style={S.sel} value={value.tipoCuenta} onChange={e=>u("tipoCuenta",e.target.value)}>
        {["Cuenta Rut","Cuenta Vista","Cuenta Corriente","Cuenta de Ahorro"].map(o=><option key={o}>{o}</option>)}
      </select>
    </div>
  );
}

export default function App() {
  const today = new Date().toISOString().split("T")[0];
  const [docTab, setDocTab] = useState("plazo_fijo");
  const [empresa, setEmpresa] = useState("Malvarrosa SpA");
  const [alert, setAlert] = useState({t:"",m:""});
  const [loading, setLoading] = useState(false);
  const [trabSheet, setTrabSheet] = useState([]);
  const [loadingTrabs, setLoadingTrabs] = useState(false);

  // Cargar trabajadores del Sheet cuando cambia empresa o tab
  const cargarTrabajadores = useCallback(async (emp) => {
    setLoadingTrabs(true);
    const lista = await apiListar(emp);
    setTrabSheet(lista.filter(t => t.estado === "Activo"));
    setLoadingTrabs(false);
  }, []);

  useEffect(() => {
    if (docTab !== "plazo_fijo") cargarTrabajadores(empresa);
  }, [empresa, docTab, cargarTrabajadores]);

  const [pf, setPf] = useState({nombre:"",apellido:"",rut:"",nac:"Chilena",nacF:"",civil:"Soltero(a)",calle:"",num:"",comuna:"",ciudad:"",email:"",tel:"",banco:"",cuenta:"",tipoCuenta:"Cuenta Rut"});
  const [inicio, setInicio] = useState(today);
  const [termino, setTermino] = useState(today);
  const [cargoKey, setCargoKey] = useState("");
  const [cargoOtro, setCargoOtro] = useState("");
  const [funciones, setFunciones] = useState("");
  const [horas, setHoras] = useState(30);
  const [colacion, setColacion] = useState(30);
  const [sueldo, setSueldo] = useState(385000);
  const [turno, setTurno] = useState(DIAS_C.map(()=>({on:false,e:"",s:""})));
  const updTurno = useCallback((i,k,val)=>setTurno(p=>p.map((r,j)=>j===i?{...r,[k]:val}:r)),[]);

  const [tNombre, setTNombre] = useState("");
  const [tRut, setTRut] = useState("");
  const [tCargo, setTCargo] = useState("");

  const [anFecha, setAnFecha] = useState(today);
  const [anTipo, setAnTipo] = useState("plazo");
  const [anDesc, setAnDesc] = useState("");
  const [anCond, setAnCond] = useState("");

  const [aiFecha, setAiFecha] = useState(today);
  const [aiDias, setAiDias] = useState("");
  const [aiMotivo, setAiMotivo] = useState("");
  const [aiReinc, setAiReinc] = useState(false);

  const [aaFecha, setAaFecha] = useState(today);
  const [aaDetalle, setAaDetalle] = useState("");
  const [aaReinc, setAaReinc] = useState(false);

  const [nrFechaDoc, setNrFechaDoc] = useState(today);
  const [nrFechaTermino, setNrFechaTermino] = useState(today);
  const [nrMotivo, setNrMotivo] = useState("");

  const getCargoN = () => {
    if (cargoKey==="garzon") return "Garzon - Barista";
    if (cargoKey==="cocina") return "Personal de Cocina y Produccion";
    if (cargoKey==="operario") return "Operario Multifuncional de Almacen y Alimentos al Paso";
    return cargoOtro;
  };

  const generar = async () => {
    setLoading(true); setAlert({t:"",m:""});
    try {
      if (docTab==="plazo_fijo") {
        if (!pf.nombre||!pf.apellido) throw new Error("Ingresa nombre y apellido");
        if (!getCargoN()) throw new Error("Selecciona un cargo");
        await generarContratoPF({...pf, empresa, inicio, termino, horas, colacion, sueldo, turno, funciones, cargo:getCargoN()});
        // Guardar en Google Sheets
        await apiGuardar({
          id: Date.now().toString(),
          nombre: pf.nombre, apellido: pf.apellido, rut: pf.rut,
          empresa, cargo: getCargoN(),
          domicilio: [pf.calle,pf.num,pf.comuna,pf.ciudad].filter(Boolean).join(", "),
          email: pf.email, telefono: pf.tel,
          banco: pf.banco, tipoCuenta: pf.tipoCuenta, cuenta: pf.cuenta,
          fechaIngreso: inicio, estado: "Activo"
        });
      } else if (docTab==="anexo") {
        if (!tNombre) throw new Error("Ingresa el nombre del trabajador");
        if (!anDesc) throw new Error("Describe el cambio");
        await generarAnexo({empresa, tNombre, tRut, tCargo, anFecha, anTipo, anDesc, anCond});
      } else if (docTab==="am_inasistencia") {
        if (!tNombre) throw new Error("Ingresa el nombre del trabajador");
        if (!aiDias) throw new Error("Ingresa las fechas de inasistencia");
        await generarAmonInasistencia({empresa, tNombre, tRut, tCargo, aiFecha, aiDias, aiMotivo, aiReinc});
      } else if (docTab==="am_atraso") {
        if (!tNombre) throw new Error("Ingresa el nombre del trabajador");
        if (!aaDetalle) throw new Error("Ingresa el detalle de los atrasos");
        await generarAmonAtraso({empresa, tNombre, tRut, tCargo, aaFecha, aaDetalle, aaReinc});
      } else if (docTab==="no_renovacion") {
        if (!tNombre) throw new Error("Ingresa el nombre del trabajador");
        await generarNoRenovacion({empresa, tNombre, tRut, tCargo, nrFechaDoc, nrFechaTermino, nrMotivo});
        // Buscar si el trabajador ya existe en el Sheet
        const listaActual = await apiListar(empresa);
        const existe = listaActual.find(t =>
          t.rut === tRut || t.nombre+" "+t.apellido === tNombre
        );
        if (existe) {
          // Desactivar el existente
          await apiDesactivar(existe.id);
        } else {
          // Crear como inactivo si no existia
          const partes = tNombre.trim().split(" ");
          const apellido = partes.slice(-1)[0];
          const nombre = partes.slice(0,-1).join(" ");
          await apiGuardar({
            id: Date.now().toString(),
            nombre, apellido, rut: tRut,
            empresa, cargo: tCargo,
            domicilio: "", email: "", telefono: "",
            banco: "", tipoCuenta: "", cuenta: "",
            fechaIngreso: "", estado: "Inactivo"
          });
        }
        // Refrescar lista
        await cargarTrabajadores(empresa);
      }
      setAlert({t:"s", m:"Documento generado y descargado correctamente."});
    } catch(e) {
      setAlert({t:"e", m:"Error: "+e.message});
    }
    setLoading(false);
  };

  const DOC_TYPES = [
    ["plazo_fijo","📅","Contrato PF"],
    ["anexo","📄","Anexo"],
    ["am_inasistencia","🗓️","Amonest. inasist."],
    ["am_atraso","⏰","Amonest. atrasos"],
    ["no_renovacion","🚫","No renovacion"],
  ];

  return (
    <div style={{maxWidth:480,margin:"0 auto",background:"#f5f5f0",minHeight:"100vh",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif"}}>
      <div style={{background:"#185FA5",color:"white",padding:"1rem 1.2rem",display:"flex",alignItems:"center",gap:10,position:"sticky",top:0,zIndex:10}}>
        <div style={{fontSize:22}}>📄</div>
        <div>
          <div style={{fontSize:16,fontWeight:700}}>Documentos Laborales</div>
          <div style={{fontSize:11,opacity:.8}}>Malvarrosa SpA · Eric Hansen EIRL</div>
        </div>
      </div>

      <div style={{padding:"1rem 1rem 4rem"}}>
        <div style={S.card}>
          <div style={{fontSize:14,fontWeight:700,marginBottom:10,color:"#444"}}>Empresa</div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {["Malvarrosa SpA","Eric Hansen EIRL"].map(e=>(
              <button key={e} style={pill(empresa===e)} onClick={()=>setEmpresa(e)}>{e}</button>
            ))}
          </div>
        </div>

        <div style={S.card}>
          <div style={{fontSize:14,fontWeight:700,marginBottom:10,color:"#444"}}>Tipo de documento</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
            {DOC_TYPES.map(([k,ico,lbl])=>(
              <div key={k} style={dtcSt(docTab===k)} onClick={()=>{setDocTab(k);setAlert({t:"",m:""});}}>
                <div style={{fontSize:18,marginBottom:3}}>{ico}</div>
                <div style={{fontSize:10,color:docTab===k?"#185FA5":"#666",fontWeight:docTab===k?600:400,lineHeight:1.3}}>{lbl}</div>
              </div>
            ))}
          </div>
        </div>

        {docTab==="plazo_fijo" && <>
          <FormPF value={pf} onChange={setPf}/>
          <div style={S.card}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:14}}>Condiciones del contrato</div>
            <div style={S.r2}>
              <div><LBL first>Fecha inicio</LBL><input style={S.inp} type="date" value={inicio} onChange={e=>setInicio(e.target.value)}/></div>
              <div><LBL first>Fecha termino</LBL><input style={S.inp} type="date" value={termino} onChange={e=>setTermino(e.target.value)}/></div>
            </div>
            <LBL>Cargo</LBL>
            <select style={S.sel} value={cargoKey} onChange={e=>{const k=e.target.value;setCargoKey(k);setFunciones(FN[k]||"");}}>
              <option value="">-- Selecciona un cargo --</option>
              <option value="garzon">Garzon - Barista</option>
              <option value="cocina">Personal de Cocina y Produccion</option>
              <option value="operario">Operario Multifuncional de Almacen y Alimentos al Paso</option>
              <option value="otro">Otro</option>
            </select>
            {cargoKey==="otro"&&<><LBL>Nombre del cargo</LBL><input style={S.inp} value={cargoOtro} onChange={e=>setCargoOtro(e.target.value)} placeholder="Ej: Cajero"/></>}
            <LBL>Funciones <span style={{color:"#999",fontSize:11}}>(editable)</span></LBL>
            <textarea style={S.ta} value={funciones} onChange={e=>setFunciones(e.target.value)} placeholder="Selecciona un cargo..."/>
            <LBL>Lugar de trabajo</LBL>
            <input style={{...S.inp,background:"#f5f5f0",color:"#666"}} value={getLugar(empresa)} readOnly/>
            <LBL>Jornada semanal</LBL>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              {[[42,"42 hrs"],[30,"30 hrs"],[20,"20 hrs"]].map(([h,lbl])=>(
                <button key={h} style={pill(horas===h)} onClick={()=>{setHoras(h);setSueldo(SUELDOS[h]);}}>{lbl}</button>
              ))}
            </div>
            <LBL>Colacion</LBL>
            <div style={{display:"flex",gap:8}}>
              {[[30,"30 min"],[60,"1 hora"]].map(([m,lbl])=>(
                <button key={m} style={pill(colacion===m)} onClick={()=>setColacion(m)}>{lbl}</button>
              ))}
            </div>
            <LBL>Sueldo base mensual</LBL>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <input style={{...S.inp,flex:1}} type="number" value={sueldo} onChange={e=>setSueldo(parseInt(e.target.value)||0)}/>
              <span style={{background:"#EBF3FF",color:"#185FA5",padding:"5px 12px",borderRadius:20,fontWeight:700,fontSize:13,whiteSpace:"nowrap"}}>${sueldo.toLocaleString("es-CL")}</span>
            </div>
            <div style={{fontSize:11,color:"#999",marginTop:4}}>42h=$539.000 | 30h=$385.000 | 20h=$256.667</div>
            <LBL>Horario semanal</LBL>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr>{["Dia","","Entrada","Salida"].map((h,i)=><th key={i} style={{background:"#f5f5f0",padding:"6px 4px",textAlign:"center",border:"1px solid #e8e8e5",color:"#888",fontWeight:500}}>{h}</th>)}</tr></thead>
              <tbody>{DIAS_C.map((d,i)=>(
                <tr key={i} style={{background:turno[i].on?"white":"#fafaf8"}}>
                  <td style={{padding:"4px 6px",border:"1px solid #e8e8e5",fontWeight:600,color:turno[i].on?"#185FA5":"#bbb",textAlign:"center",fontSize:12}}>{d}</td>
                  <td style={{padding:4,border:"1px solid #e8e8e5",textAlign:"center"}}><input type="checkbox" checked={turno[i].on} onChange={ev=>updTurno(i,"on",ev.target.checked)} style={{width:18,height:18,cursor:"pointer",accentColor:"#185FA5"}}/></td>
                  <td style={{padding:3,border:"1px solid #e8e8e5"}}><input value={turno[i].e} onChange={ev=>updTurno(i,"e",ev.target.value)} placeholder="09:00" disabled={!turno[i].on} style={{border:"none",background:"transparent",color:turno[i].on?"#1a1a1a":"#ccc",padding:"4px",width:"100%",fontSize:13,textAlign:"center",outline:"none"}}/></td>
                  <td style={{padding:3,border:"1px solid #e8e8e5"}}><input value={turno[i].s} onChange={ev=>updTurno(i,"s",ev.target.value)} placeholder="18:00" disabled={!turno[i].on} style={{border:"none",background:"transparent",color:turno[i].on?"#1a1a1a":"#ccc",padding:"4px",width:"100%",fontSize:13,textAlign:"center",outline:"none"}}/></td>
                </tr>
              ))}</tbody>
            </table>
          </div>
        </>}

        {docTab==="anexo" && <>
          <FormTrabSimple
            nombre={tNombre} onNombre={setTNombre}
            rut={tRut} onRut={setTRut}
            cargo={tCargo} onCargo={setTCargo}
            trabSheet={trabSheet} loadingTrabs={loadingTrabs}
            onSelectTrab={t=>{ setTNombre(t.nombre+" "+t.apellido); setTRut(t.rut); setTCargo(t.cargo||""); }}
          />
          <div style={S.card}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:14}}>Datos del anexo</div>
            <LBL first>Fecha del anexo</LBL><input style={S.inp} type="date" value={anFecha} onChange={e=>setAnFecha(e.target.value)}/>
            <LBL>Tipo de modificacion</LBL>
            <select style={S.sel} value={anTipo} onChange={e=>setAnTipo(e.target.value)}>
              <option value="plazo">Prorroga de plazo fijo</option>
              <option value="indefinido">Conversion a indefinido (Art.159 N4 CT)</option>
              <option value="sueldo">Cambio de sueldo</option>
              <option value="cargo">Cambio de cargo</option>
              <option value="horario">Cambio de horario</option>
              <option value="otro">Otro</option>
            </select>
            <LBL>Descripcion del cambio</LBL>
            <textarea style={S.ta} value={anDesc} onChange={e=>setAnDesc(e.target.value)} placeholder="Describe que cambia..."/>
            <LBL>Nueva condicion contractual</LBL>
            <textarea style={S.ta} value={anCond} onChange={e=>setAnCond(e.target.value)} placeholder="Ej: A partir del [fecha], la jornada sera de..."/>
          </div>
        </>}

        {docTab==="am_inasistencia" && <>
          <FormTrabSimple
            nombre={tNombre} onNombre={setTNombre}
            rut={tRut} onRut={setTRut}
            cargo={tCargo} onCargo={setTCargo}
            trabSheet={trabSheet} loadingTrabs={loadingTrabs}
            onSelectTrab={t=>{ setTNombre(t.nombre+" "+t.apellido); setTRut(t.rut); setTCargo(t.cargo||""); }}
          />
          <div style={S.card}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:14}}>Amonestacion por inasistencias</div>
            <LBL first>Fecha de la carta</LBL><input style={S.inp} type="date" value={aiFecha} onChange={e=>setAiFecha(e.target.value)}/>
            <LBL>Dias de inasistencia (fechas exactas)</LBL>
            <input style={S.inp} value={aiDias} onChange={e=>setAiDias(e.target.value)} placeholder="Ej: 12, 13 y 15 de mayo de 2026"/>
            <LBL>Motivo o justificacion entregada</LBL>
            <textarea style={S.ta} value={aiMotivo} onChange={e=>setAiMotivo(e.target.value)} placeholder="Sin justificacion / El trabajador indico que..."/>
            <LBL>Reincidencia</LBL>
            <div style={{display:"flex",gap:8,marginTop:6}}>
              <button style={pill(!aiReinc)} onClick={()=>setAiReinc(false)}>No — primera vez</button>
              <button style={pill(aiReinc)} onClick={()=>setAiReinc(true)}>Si — reincidente</button>
            </div>
          </div>
        </>}

        {docTab==="am_atraso" && <>
          <FormTrabSimple
            nombre={tNombre} onNombre={setTNombre}
            rut={tRut} onRut={setTRut}
            cargo={tCargo} onCargo={setTCargo}
            trabSheet={trabSheet} loadingTrabs={loadingTrabs}
            onSelectTrab={t=>{ setTNombre(t.nombre+" "+t.apellido); setTRut(t.rut); setTCargo(t.cargo||""); }}
          />
          <div style={S.card}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:14}}>Amonestacion por atrasos</div>
            <LBL first>Fecha de la carta</LBL><input style={S.inp} type="date" value={aaFecha} onChange={e=>setAaFecha(e.target.value)}/>
            <LBL>Detalle de los atrasos</LBL>
            <textarea style={S.ta} value={aaDetalle} onChange={e=>setAaDetalle(e.target.value)} placeholder="Ej: 12/05 llegada 09:20 (20 min), 14/05 llegada 09:35 (35 min)"/>
            <LBL>Reincidencia</LBL>
            <div style={{display:"flex",gap:8,marginTop:6}}>
              <button style={pill(!aaReinc)} onClick={()=>setAaReinc(false)}>No — primera vez</button>
              <button style={pill(aaReinc)} onClick={()=>setAaReinc(true)}>Si — reincidente</button>
            </div>
          </div>
        </>}

        {docTab==="no_renovacion" && <>
          <FormTrabSimple
            nombre={tNombre} onNombre={setTNombre}
            rut={tRut} onRut={setTRut}
            cargo={tCargo} onCargo={setTCargo}
            trabSheet={trabSheet} loadingTrabs={loadingTrabs}
            onSelectTrab={t=>{ setTNombre(t.nombre+" "+t.apellido); setTRut(t.rut); setTCargo(t.cargo||""); }}
          />
          <div style={S.card}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:14}}>Carta de no renovacion</div>
            <div style={{background:"#FFF3E0",borderRadius:8,padding:"10px 14px",fontSize:12,color:"#E65100",marginBottom:14,border:"1px solid #FFB74D"}}>Art. 159 N4 CT — Notificacion formal de que el contrato a plazo fijo NO sera renovado.</div>
            <div style={S.r2}>
              <div><LBL first>Fecha del documento</LBL><input style={S.inp} type="date" value={nrFechaDoc} onChange={e=>setNrFechaDoc(e.target.value)}/></div>
              <div><LBL first>Fecha termino contrato</LBL><input style={S.inp} type="date" value={nrFechaTermino} onChange={e=>setNrFechaTermino(e.target.value)}/></div>
            </div>
            <LBL>Motivo <span style={{color:"#999",fontSize:11}}>(opcional)</span></LBL>
            <textarea style={S.ta} value={nrMotivo} onChange={e=>setNrMotivo(e.target.value)} placeholder="Ej: Necesidades operacionales, termino de proyecto..."/>
          </div>
        </>}

        <button style={{...S.btnP,opacity:loading?0.7:1}} onClick={generar} disabled={loading}>
          {loading?"⏳ Generando...":"⬇️ Generar y descargar documento"}
        </button>
        {alert.m && <div style={alert.t==="e"?S.aE:S.aS}>{alert.m}</div>}
      </div>
    </div>
  );
}
