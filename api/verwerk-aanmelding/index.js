const { getGraphClient } = require('../shared/graph');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');

// ── PDF generator (inline, same logic as generate-pdf function) ──────────────
async function buildPdf(d) {
  const ACCENT    = rgb(0/255, 138/255, 209/255);
  const ACCENT_DK = rgb(0/255, 110/255, 168/255);
  const WHITE     = rgb(1, 1, 1);
  const TEXT      = rgb(26/255, 29/255, 35/255);
  const TEXT2     = rgb(90/255, 96/255, 112/255);
  const BORDER    = rgb(228/255, 231/255, 237/255);
  const OK_BG     = rgb(237/255, 250/255, 243/255);
  const OK        = rgb(26/255, 127/255, 75/255);

  const pdfDoc = await PDFDocument.create();
  const page   = pdfDoc.addPage([595.28, 841.89]);
  const { width, height } = page.getSize();
  const fontReg  = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  page.drawRectangle({ x: 0, y: height - 80, width, height: 80, color: ACCENT });
  page.drawText('Aanmelding Vrijwilliger', { x: 24, y: height - 36, size: 20, font: fontBold, color: WHITE });
  page.drawText('de Fietsboot — Stichting Veerdiensten op de Vecht', { x: 24, y: height - 56, size: 10, font: fontReg, color: rgb(0.85, 0.93, 0.98) });

  const nu = new Date().toLocaleDateString('nl-NL', { day:'2-digit', month:'long', year:'numeric' });
  page.drawText(nu, { x: width - 120, y: height - 46, size: 9, font: fontReg, color: WHITE });

  let cy = height - 100;
  const ML = 24, MR = 24, CW = width - ML - MR;

  function sectionHeader(title) {
    cy -= 14;
    page.drawRectangle({ x: ML, y: cy - 2, width: CW, height: 18, color: ACCENT });
    page.drawText(title.toUpperCase(), { x: ML + 8, y: cy + 2, size: 8.5, font: fontBold, color: WHITE });
    cy -= 10;
  }

  function row(label, value, highlight = false) {
    if (!value && value !== 0) return;
    const rowH = 18;
    if (highlight) page.drawRectangle({ x: ML, y: cy - rowH + 4, width: CW, height: rowH, color: OK_BG });
    page.drawText(label, { x: ML + 6, y: cy - 10, size: 8.5, font: fontBold, color: TEXT2 });
    page.drawText(String(value), { x: ML + 170, y: cy - 10, size: 9, font: fontReg, color: highlight ? OK : TEXT });
    page.drawLine({ start: { x: ML, y: cy - rowH + 4 }, end: { x: ML + CW, y: cy - rowH + 4 }, thickness: 0.4, color: BORDER });
    cy -= rowH;
  }

  function gap(px = 6) { cy -= px; }

  sectionHeader('Persoonsgegevens');
  gap(4);
  const tussenvoeg = d.tussenvoegsel ? ` ${d.tussenvoegsel}` : '';
  const volledigeNaam = `${d.voornaam || ''}${tussenvoeg} ${d.achternaam || ''}`.trim();
  row('Naam', volledigeNaam);
  row('Adres', d.adres ? `${d.adres}, ${d.postcode || ''} ${d.woonplaats || ''}`.trim() : '');
  row('Telefoon (vast)', d.telefoon);
  row('Mobiel', d.mobiel);
  row('Noodcontact', d.telefoon_nood);
  row('E-mailadres', d.email);
  row('Geboortedatum', d.geboortedatum);
  row('Beroep / werkervaring', d.beroep);
  row('Ander vrijwilligerswerk', d.ander_vrijwilligerswerk);
  gap(8);

  sectionHeader('Aanmelding');
  gap(4);
  row('Voorkeur vaargebied', d.vaargebied);
  row('Aangemeld voor functie', d.functie, true);
  gap(8);

  if (d.functie === 'Schipper (m/v)') {
    sectionHeader('Vaarervaring Schipper');
    gap(4);
    row('Jaren ervaring motorschepen', d.jaren_ervaring);
    row('Lengte schip (m)', d.lengte_schip);
    row('Type schepen', d.type_schepen);
    row('Bekendheid Utrechtse Vecht', d.vecht_schipper);
    row('Bekendheid Loosdrechtse Plassen', d.loosdrecht_schipper);
    gap(8);
  } else if (d.functie === 'Gastheer (m/v)') {
    sectionHeader('Vaarervaring Gastheer');
    gap(4);
    row('Bekendheid Utrechtse Vecht', d.vecht_gastheer);
    row('Bekendheid Loosdrechtse Plassen', d.loosdrecht_gastheer);
    gap(8);
  }

  const certs = Array.isArray(d.certificaten) ? d.certificaten : [];
  if (certs.length > 0 || d.bijlagen_opmerking) {
    sectionHeader('Certificaten & Bijlagen');
    gap(4);
    if (certs.length) row('Certificaten', certs.join(', '));
    if (d.bijlagen_opmerking) row('Bijlagen opmerking', d.bijlagen_opmerking);
    gap(8);
  }

  if (d.bijzonderheden) {
    sectionHeader('Bijzonderheden / Opmerkingen');
    gap(4);
    const words = d.bijzonderheden.split(' ');
    let line = '';
    const lineH = 13;
    cy -= 10;
    for (const w of words) {
      const test = line ? `${line} ${w}` : w;
      if (fontReg.widthOfTextAtSize(test, 9) > CW - 12) {
        page.drawText(line, { x: ML + 6, y: cy, size: 9, font: fontReg, color: TEXT });
        cy -= lineH; line = w;
      } else { line = test; }
    }
    if (line) { page.drawText(line, { x: ML + 6, y: cy, size: 9, font: fontReg, color: TEXT }); cy -= lineH; }
    gap(8);
  }

  if (d.avg_akkoord) {
    gap(4);
    page.drawRectangle({ x: ML, y: cy - 14, width: CW, height: 22, color: OK_BG });
    page.drawText('✓  AVG-toestemming gegeven', { x: ML + 8, y: cy - 6, size: 9, font: fontBold, color: OK });
    cy -= 26;
  }

  page.drawRectangle({ x: 0, y: 0, width, height: 28, color: ACCENT });
  page.drawText('www.defietsboot.nl  |  secretaris@defietsboot.nl', { x: ML, y: 10, size: 8, font: fontReg, color: WHITE });
  page.drawText(`Aangemeld op ${nu}`, { x: width - 160, y: 10, size: 8, font: fontReg, color: WHITE });

  return Buffer.from(await pdfDoc.save());
}

// ── Excel: append row to aanmeldingen.xlsx ────────────────────────────────────
async function schrijfNaarExcel(client, d) {
  const userEmail = process.env.GRAPH_ONEDRIVE_USER;
  const filename  = 'aanmeldingen-vrijwilligers.xlsx';
  const sheetName = 'Aanmeldingen';

  let fileId, driveId;
  try {
    const search = await client.api(`/users/${userEmail}/drive/root/search(q='${filename}')`).get();
    const file = search.value.find(f => f.name === filename);
    if (file) {
      fileId  = file.id;
      driveId = file.parentReference.driveId;
    }
  } catch (e) { /* file doesn't exist yet — will be created below */ }

  const XLSX = require('xlsx');

  if (!fileId) {
    // Create new workbook with header row
    const wb = XLSX.utils.book_new();
    const headers = [['Datum','Voornaam','Tussenvoegsel','Achternaam','Adres','Postcode','Woonplaats',
      'Telefoon','Mobiel','Noodcontact','Email','Geboortedatum','Beroep','Ander vrijwilligerswerk',
      'Vaargebied','Functie','Jaren ervaring','Lengte schip','Type schepen',
      'Vecht bekendheid','Loosdrecht bekendheid','Certificaten','Bijlagen opmerking','Bijzonderheden']];
    const ws = XLSX.utils.aoa_to_sheet(headers);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const uploadPath = process.env.GRAPH_AANMELDINGEN_PATH || `MS365/Aanmeldingen/${filename}`;
    const uploaded = await client.api(`/users/${userEmail}/drive/root:/${uploadPath}:/content`)
      .header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      .put(buf);
    fileId  = uploaded.id;
    driveId = uploaded.parentReference.driveId;
  }

  const wbBase = `/drives/${driveId}/items/${fileId}/workbook`;
  const session = await client.api(`${wbBase}/createSession`).post({ persistChanges: true });
  const sid = session.id;

  try {
    // Get used range to find next empty row
    const usedRange = await client.api(`${wbBase}/worksheets/${sheetName}/usedRange`)
      .header('workbook-session-id', sid).get();
    const nextRow = usedRange.rowCount + 1; // 1-based

    const nu = new Date().toLocaleDateString('nl-NL');
    const certs = Array.isArray(d.certificaten) ? d.certificaten.join(', ') : '';
    const funkyVecht = d.functie === 'Schipper (m/v)' ? d.vecht_schipper : d.vecht_gastheer;
    const funkyLoos  = d.functie === 'Schipper (m/v)' ? d.loosdrecht_schipper : d.loosdrecht_gastheer;

    const values = [[
      nu,
      d.voornaam || '', d.tussenvoegsel || '', d.achternaam || '',
      d.adres || '', d.postcode || '', d.woonplaats || '',
      d.telefoon || '', d.mobiel || '', d.telefoon_nood || '',
      d.email || '', d.geboortedatum || '', d.beroep || '', d.ander_vrijwilligerswerk || '',
      d.vaargebied || '', d.functie || '',
      d.jaren_ervaring || '', d.lengte_schip || '', d.type_schepen || '',
      funkyVecht || '', funkyLoos || '',
      certs, d.bijlagen_opmerking || '', d.bijzonderheden || ''
    ]];

    // Append row at next position
    const rangeAddr = `A${nextRow}:X${nextRow}`;
    await client.api(`${wbBase}/worksheets/${sheetName}/range(address='${rangeAddr}')`)
      .header('workbook-session-id', sid)
      .patch({ values });

  } finally {
    await client.api(`${wbBase}/closeSession`)
      .header('workbook-session-id', sid).post({}).catch(() => {});
  }
}

// ── Email helpers ─────────────────────────────────────────────────────────────
async function stuurEmailSecretaris(client, d, pdfBytes) {
  const tussenvoeg = d.tussenvoegsel ? ` ${d.tussenvoegsel}` : '';
  const naam = `${d.voornaam || ''}${tussenvoeg} ${d.achternaam || ''}`.trim();
  const certs = Array.isArray(d.certificaten) ? d.certificaten.join(', ') : '—';

  const html = `
<div style="font-family:Arial,sans-serif;max-width:600px;color:#1a1d23">
  <div style="background:#008AD1;padding:20px 24px;border-radius:8px 8px 0 0">
    <h2 style="color:#fff;margin:0">Nieuwe aanmelding vrijwilliger</h2>
    <p style="color:#d0eaf9;margin:6px 0 0">de Fietsboot</p>
  </div>
  <div style="background:#f7f8fa;padding:20px 24px;border:1px solid #e4e7ed;border-top:none;border-radius:0 0 8px 8px">
    <p>Er is een nieuwe aanmelding binnengekomen. De gegevens zijn bijgevoegd als PDF.</p>
    <table style="width:100%;border-collapse:collapse;font-size:14px">
      <tr><td style="padding:6px 0;color:#5a6070;width:160px"><b>Naam</b></td><td>${naam}</td></tr>
      <tr><td style="padding:6px 0;color:#5a6070"><b>Functie</b></td><td style="color:#008AD1;font-weight:600">${d.functie || '—'}</td></tr>
      <tr><td style="padding:6px 0;color:#5a6070"><b>Vaargebied</b></td><td>${d.vaargebied || '—'}</td></tr>
      <tr><td style="padding:6px 0;color:#5a6070"><b>Email</b></td><td><a href="mailto:${d.email}">${d.email}</a></td></tr>
      <tr><td style="padding:6px 0;color:#5a6070"><b>Mobiel</b></td><td>${d.mobiel || '—'}</td></tr>
      <tr><td style="padding:6px 0;color:#5a6070"><b>Certificaten</b></td><td>${certs}</td></tr>
    </table>
    <p style="font-size:12px;color:#9aa0ad;margin-top:20px">Zie bijgevoegde PDF voor alle details.</p>
  </div>
</div>`;

  await client.api(`/users/${process.env.GRAPH_MAIL_FROM}/sendMail`).post({
    message: {
      subject: `Nieuwe aanmelding vrijwilliger: ${naam} (${d.functie || ''})`,
      body: { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: process.env.GRAPH_MAIL_TO || 'secretaris@defietsboot.nl' } }],
      attachments: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: `aanmelding-${naam.replace(/\s+/g, '-').toLowerCase()}.pdf`,
          contentType: 'application/pdf',
          contentBytes: pdfBytes.toString('base64'),
          isInline: false
        },
        ...(Array.isArray(d.bijlagen) ? d.bijlagen.map(b => ({
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: b.name,
          contentType: b.type || 'application/octet-stream',
          contentBytes: b.b64,
          isInline: false
        })) : [])
      ]
    },
    saveToSentItems: false
  });
}

async function stuurBevestigingAanmelder(client, d, pdfBytes) {
  const tussenvoeg = d.tussenvoegsel ? ` ${d.tussenvoegsel}` : '';
  const naam = `${d.voornaam || ''}${tussenvoeg} ${d.achternaam || ''}`.trim();

  const html = `
<div style="font-family:Arial,sans-serif;max-width:600px;color:#1a1d23">
  <div style="background:#008AD1;padding:20px 24px;border-radius:8px 8px 0 0">
    <h2 style="color:#fff;margin:0">Bedankt voor uw aanmelding!</h2>
    <p style="color:#d0eaf9;margin:6px 0 0">de Fietsboot — Stichting Veerdiensten op de Vecht</p>
  </div>
  <div style="background:#f7f8fa;padding:20px 24px;border:1px solid #e4e7ed;border-top:none;border-radius:0 0 8px 8px">
    <p>Beste ${d.voornaam || 'vrijwilliger'},</p>
    <p>Hartelijk dank voor uw aanmelding als <strong>${d.functie || 'vrijwilliger'}</strong> bij de Fietsboot. 
    Wij hebben uw aanmelding in goede orde ontvangen.</p>
    <p>Bijgevoegd vindt u een overzicht van de door u ingevulde gegevens.</p>
    <div style="background:#edfaf3;border-left:4px solid #1a7f4b;padding:12px 16px;border-radius:4px;margin:16px 0">
      <p style="margin:0;color:#1a7f4b;font-weight:600">Wat gebeurt er nu?</p>
      <p style="margin:6px 0 0;color:#1a7f4b;font-size:14px">
        Wij nemen zo spoedig mogelijk contact met u op om uw aanmelding verder te bespreken.
      </p>
    </div>
    ${d.functie === 'Schipper (m/v)' ? `
    <div style="background:#fffbeb;border-left:4px solid #92620a;padding:12px 16px;border-radius:4px;margin:16px 0">
      <p style="margin:0;color:#92620a;font-size:14px">
        <strong>Vergeet niet:</strong> stuur uw Vaarbewijs en Marifooncertificaat naar 
        <a href="mailto:secretaris@defietsboot.nl">secretaris@defietsboot.nl</a>
      </p>
    </div>` : ''}
    <p style="font-size:12px;color:#9aa0ad;margin-top:24px;border-top:1px solid #e4e7ed;padding-top:12px">
      de Fietsboot · Stichting Veerdiensten op de Vecht · 
      <a href="https://www.defietsboot.nl" style="color:#008AD1">www.defietsboot.nl</a>
    </p>
  </div>
</div>`;

  await client.api(`/users/${process.env.GRAPH_MAIL_FROM}/sendMail`).post({
    message: {
      subject: 'Bevestiging aanmelding vrijwilliger de Fietsboot',
      body: { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: d.email } }],
      attachments: [{
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: `aanmelding-vrijwilliger-defietsboot.pdf`,
        contentType: 'application/pdf',
        contentBytes: pdfBytes.toString('base64'),
        isInline: false
      }]
    },
    saveToSentItems: false
  });
}

// ── Main handler ──────────────────────────────────────────────────────────────
module.exports = async function (context, req) {
  try {
    const d = req.body;
    if (!d || !d.email || !d.voornaam || !d.achternaam) {
      context.res = { status: 400, body: 'Verplichte velden ontbreken' };
      return;
    }

    const client   = getGraphClient();
    const pdfBytes = await buildPdf(d);

    await Promise.all([
      schrijfNaarExcel(client, d),
      stuurEmailSecretaris(client, d, pdfBytes),
      stuurBevestigingAanmelder(client, d, pdfBytes)
    ]);

    context.res = { status: 200, body: { ok: true } };

  } catch (err) {
    context.log.error('verwerk-aanmelding error:', err);
    context.res = { status: 500, body: `Verwerking mislukt: ${err.message}` };
  }
};
