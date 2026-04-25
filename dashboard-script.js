function crearDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Referencia directa por nombre — sin depender del orden
  const datos = ss.getSheetByName("LipiWork Turnos");
  if (!datos) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la hoja 'LipiWork Turnos'");
    return;
  }

  let dash = ss.getSheetByName("Dashboard");
  if (!dash) dash = ss.insertSheet("Dashboard");
  else dash.clear();

  // ── TÍTULO ──
  dash.getRange("A1").setValue("DASHBOARD LIPIWORK")
    .setFontSize(18).setFontWeight("bold").setFontColor("#F47920");
  dash.getRange("A2").setValue("Actualizado: " + new Date().toLocaleString("es-CL"))
    .setFontColor("#6B7280").setFontSize(10);

  // ── MÉTRICAS ──
  // Columnas: A=Timestamp  B=Nombre  C=Día  D=Tipo  E=Horario
  dash.getRange("A4").setValue("Total Registros");
  dash.getRange("B4").setFormula("=COUNTA('LipiWork Turnos'!A2:A)");

  dash.getRange("A5").setValue("Turnos 4 Horas");
  dash.getRange("B5").setFormula("=COUNTIF('LipiWork Turnos'!D2:D;\"4 Horas\")");

  dash.getRange("A6").setValue("Turnos 8 Horas");
  dash.getRange("B6").setFormula("=COUNTIF('LipiWork Turnos'!D2:D;\"8 Horas\")");

  dash.getRange("A7").setValue("Personas distintas");
  dash.getRange("B7").setFormula("=SUMPRODUCTO((1/CONTAR.SI('LipiWork Turnos'!B2:B;'LipiWork Turnos'!B2:B))*('LipiWork Turnos'!B2:B<>\"\"))");

  // ── TABLA POR DÍA ──
  const dias = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"];
  dash.getRange("D4").setValue("Día").setFontWeight("bold").setBackground("#1C2F5E").setFontColor("white");
  dash.getRange("E4").setValue("Turnos").setFontWeight("bold").setBackground("#1C2F5E").setFontColor("white");
  dias.forEach((d, i) => {
    dash.getRange(5 + i, 4).setValue(d);
    dash.getRange(5 + i, 5).setFormula(`=COUNTIF('LipiWork Turnos'!C2:C;"${d}")`);
  });

  // ── TABLA TIPO ──
  dash.getRange("D13").setValue("Tipo").setFontWeight("bold").setBackground("#1C2F5E").setFontColor("white");
  dash.getRange("E13").setValue("Cantidad").setFontWeight("bold").setBackground("#1C2F5E").setFontColor("white");
  dash.getRange("D14").setValue("4 Horas");
  dash.getRange("E14").setFormula("=COUNTIF('LipiWork Turnos'!D2:D;\"4 Horas\")");
  dash.getRange("D15").setValue("8 Horas");
  dash.getRange("E15").setFormula("=COUNTIF('LipiWork Turnos'!D2:D;\"8 Horas\")");

  // ── GRÁFICOS ──
  dash.getCharts().forEach(c => dash.removeChart(c));

  const chartDia = dash.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dash.getRange("D4:E11"))
    .setPosition(14, 1, 0, 0)
    .setOption("title", "Turnos por día de la semana")
    .setOption("colors", ["#F47920"])
    .setOption("width", 420).setOption("height", 280)
    .build();
  dash.insertChart(chartDia);

  const chartTipo = dash.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dash.getRange("D13:E15"))
    .setPosition(14, 5, 0, 0)
    .setOption("title", "Distribución 4h vs 8h")
    .setOption("colors", ["#F47920","#1C2F5E"])
    .setOption("width", 320).setOption("height", 280)
    .build();
  dash.insertChart(chartTipo);

  SpreadsheetApp.getUi().alert("✅ Dashboard listo con " + (datos.getLastRow()-1) + " registros");
}

function onFormSubmit() {
  crearDashboard();
}
