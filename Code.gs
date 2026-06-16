// ─────────────────────────────────────────────────────────────
//  CONSULTORIO ESCOLAR — CBTis No. 179
//  Pega este código en Google Apps Script y despliega como
//  Web App (Ejecutar como: Yo | Acceso: Cualquier persona)
// ─────────────────────────────────────────────────────────────

const HEADERS = [
  'Fecha de registro',
  'Apellido paterno', 'Apellido materno', 'Nombre(s)',
  'Fecha nacimiento', 'CURP', 'Semestre', 'Especialidad', 'Turno', 'Correo',
  'Tutor nombre', 'Tel. principal', 'Tel. alternativo', 'Parentesco', 'Aseguradora',
  'Estilo de aprendizaje', 'Socialización',
  'Disc. física', 'Disc. visual', 'Disc. auditiva', 'Disc. intelectual', 'Disc. psicosocial',
  'Necesidades específicas', 'Tiene certificado',
  'Enf. crónicas', 'Otra enfermedad', 'Medicamentos',
  'Con quién vive', 'Personas en hogar', 'Estudiantes en hogar',
  'Tipo vivienda', 'Servicios', 'Ingreso familiar', 'Proveedor económico',
  'Trabaja', 'Escolaridad padre', 'Escolaridad madre',
  'Dispositivos', 'Becas', 'Transporte', 'Tiempo traslado',
  'Firma tutor', 'Firma alumno', 'Fecha firma'
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.openById('1RIeChU3mA5fTy_-3gT2IrfVah-X_-toBqjjjdLOv-34').getActiveSheet();

    // Crear encabezados con formato si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      var hRange = sheet.getRange(1, 1, 1, HEADERS.length);
      hRange.setBackground('#1a6b3c');
      hRange.setFontColor('#ffffff');
      hRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160);
    }

    var arr = function(val) {
      return Array.isArray(val) ? val.join(', ') : (val || '');
    };

    var row = [
      new Date(),
      data.ap_paterno      || '', data.ap_materno    || '', data.nombres       || '',
      data.fecha_nac       || '', data.curp           || '', data.semestre      || '',
      data.especialidad    || '', data.turno          || '', data.email         || '',
      data.tutor_nombre    || '', data.tel_principal  || '', data.tel_alt       || '',
      data.parentesco      || '', data.aseguradora    || '',
      data.estilo_aprendizaje || '', data.socializacion || '',
      arr(data.disc_fisica), arr(data.disc_visual), arr(data.disc_auditiva),
      arr(data.disc_intelectual), arr(data.disc_psico),
      data.desc_necesidades   || '', data.tiene_certificado || '',
      arr(data.enf_cronica), data.enf_cronica_otro || '', data.medicamentos || '',
      data.con_quien_vive     || '', data.num_personas_hogar    || '',
      data.num_estudiantes_hogar || '', data.tipo_vivienda      || '',
      arr(data.servicios), data.ingreso_familia || '', data.proveedor     || '',
      data.trabaja            || '', data.escolaridad_padre     || '',
      data.escolaridad_madre  || '',
      arr(data.dispositivos), arr(data.becas),
      data.transporte         || '', data.tiempo_traslado       || '',
      data.firma_tutor        || '', data.firma_alumno          || '',
      data.fecha_firma        || ''
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, mensaje: 'API Consultorio Escolar activa' }))
    .setMimeType(ContentService.MimeType.JSON);
}
