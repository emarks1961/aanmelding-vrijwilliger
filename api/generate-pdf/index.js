const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');

// Fietsboot brand colours
const ACCENT    = rgb(0/255, 138/255, 209/255);   // #008AD1
const ACCENT_DK = rgb(0/255, 110/255, 168/255);   // #006ea8
const WHITE     = rgb(1, 1, 1);
const TEXT      = rgb(26/255, 29/255, 35/255);     // #1a1d23
const TEXT2     = rgb(90/255, 96/255, 112/255);    // #5a6070
const BORDER    = rgb(228/255, 231/255, 237/255);  // #e4e7ed
const OK_BG     = rgb(237/255, 250/255, 243/255);  // #edfaf3
const OK        = rgb(26/255, 127/255, 75/255);    // #1a7f4b

module.exports = async function (context, req) {
  try {
    const d = req.body;
    if (!d) {
      context.res = { status: 400, body: 'Geen data ontvangen' };
      return;
    }

    const pdfDoc = await PDFDocument.create();
    const page   = pdfDoc.addPage([595.28, 841.89]); // A4
    const { width, height } = page.getSize();

    const fontReg  = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    // ── Header band ──────────────────────────────────────────────────────────
    page.drawRectangle({ x: 0, y: height - 80, width, height: 80, color: ACCENT });

    // Title
    page.drawText('Aanmelding Vrijwilliger', {
      x: 24, y: height - 36,
      size: 20, font: fontBold, color: WHITE
    });
    page.drawText('de Fietsboot — Stichting Veerdiensten op de Vecht', {
      x: 24, y: height - 56,
      size: 10, font: fontReg, color: rgb(0.85, 0.93, 0.98)
    });

    // Date stamp top-right
    const nu = new Date().toLocaleDateString('nl-NL', { day:'2-digit', month:'long', year:'numeric' });
    page.drawText(nu, {
      x: width - 120, y: height - 46,
      size: 9, font: fontReg, color: WHITE
    });

    // ── Helpers ───────────────────────────────────────────────────────────────
    let cy = height - 100; // current Y (top of available area)
    const ML = 24;         // margin left
    const MR = 24;         // margin right
    const CW = width - ML - MR;

    function sectionHeader(title) {
      cy -= 14;
      page.drawRectangle({ x: ML, y: cy - 2, width: CW, height: 18, color: ACCENT });
      page.drawText(title.toUpperCase(), {
        x: ML + 8, y: cy + 2,
        size: 8.5, font: fontBold, color: WHITE
      });
      cy -= 10;
    }

    function row(label, value, highlight = false) {
      if (!value && value !== 0) return; // skip empty
      const rowH = 18;
      if (highlight) {
        page.drawRectangle({ x: ML, y: cy - rowH + 4, width: CW, height: rowH, color: OK_BG });
      }
      page.drawText(label, {
        x: ML + 6, y: cy - 10,
        size: 8.5, font: fontBold, color: TEXT2
      });
      page.drawText(String(value), {
        x: ML + 170, y: cy - 10,
        size: 9, font: fontReg, color: highlight ? OK : TEXT
      });
      // subtle bottom border
      page.drawLine({
        start: { x: ML, y: cy - rowH + 4 },
        end:   { x: ML + CW, y: cy - rowH + 4 },
        thickness: 0.4, color: BORDER
      });
      cy -= rowH;
    }

    function gap(px = 6) { cy -= px; }

    // ── Sectie 1: Persoonsgegevens ────────────────────────────────────────────
    sectionHeader('Persoonsgegevens');
    gap(4);

    const voornaam    = d.voornaam || '';
    const tussenvoeg  = d.tussenvoegsel ? ` ${d.tussenvoegsel}` : '';
    const achternaam  = d.achternaam || '';
    const volledigeNaam = `${voornaam}${tussenvoeg} ${achternaam}`.trim();

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

    // ── Sectie 2: Aanmelding ──────────────────────────────────────────────────
    sectionHeader('Aanmelding');
    gap(4);
    row('Voorkeur vaargebied', d.vaargebied);
    row('Aangemeld voor functie', d.functie, true);
    gap(8);

    // ── Sectie 3/4: Vaarervaring ──────────────────────────────────────────────
    if (d.functie === 'Schipper (m/v)') {
      sectionHeader('Vaarervaring Schipper');
      gap(4);
      row('Jaren ervaring motorschepen', d.jaren_ervaring);
      row('Lengte schip (m)', d.lengte_schip);
      row('Type schepen', d.type_schepen);
      row('Bekendheid Utrechtse Vecht', d.vecht_schipper);
      row('Bekendheid Loosdrechtse Plassen', d.loosdrecht_schipper);
    } else if (d.functie === 'Gastheer (m/v)') {
      sectionHeader('Vaarervaring Gastheer');
      gap(4);
      row('Bekendheid Utrechtse Vecht', d.vecht_gastheer);
      row('Bekendheid Loosdrechtse Plassen', d.loosdrecht_gastheer);
    }
    gap(8);

    // ── Sectie 5: Certificaten ────────────────────────────────────────────────
    const certs = Array.isArray(d.certificaten) ? d.certificaten : [];
    if (certs.length > 0 || d.bijlagen_opmerking) {
      sectionHeader('Certificaten & Bijlagen');
      gap(4);
      if (certs.length) row('Certificaten', certs.join(', '));
      if (d.bijlagen_opmerking) row('Bijlagen opmerking', d.bijlagen_opmerking);
      gap(8);
    }

    // ── Opmerkingen ───────────────────────────────────────────────────────────
    if (d.bijzonderheden) {
      sectionHeader('Bijzonderheden / Opmerkingen');
      gap(4);
      // Wrap long text manually
      const words = d.bijzonderheden.split(' ');
      let line = '';
      const lineH = 13;
      cy -= 10;
      for (const w of words) {
        const test = line ? `${line} ${w}` : w;
        if (fontReg.widthOfTextAtSize(test, 9) > CW - 12) {
          page.drawText(line, { x: ML + 6, y: cy, size: 9, font: fontReg, color: TEXT });
          cy -= lineH;
          line = w;
        } else { line = test; }
      }
      if (line) {
        page.drawText(line, { x: ML + 6, y: cy, size: 9, font: fontReg, color: TEXT });
        cy -= lineH;
      }
      gap(8);
    }

    // ── AVG ───────────────────────────────────────────────────────────────────
    if (d.avg_akkoord) {
      gap(4);
      page.drawRectangle({ x: ML, y: cy - 14, width: CW, height: 22, color: OK_BG });
      page.drawText('✓  AVG-toestemming gegeven', {
        x: ML + 8, y: cy - 6,
        size: 9, font: fontBold, color: OK
      });
      cy -= 26;
    }

    // ── Footer ────────────────────────────────────────────────────────────────
    page.drawRectangle({ x: 0, y: 0, width, height: 28, color: ACCENT });
    page.drawText('www.defietsboot.nl  |  secretaris@defietsboot.nl', {
      x: ML, y: 10,
      size: 8, font: fontReg, color: WHITE
    });
    page.drawText(`Aangemeld op ${nu}`, {
      x: width - 160, y: 10,
      size: 8, font: fontReg, color: WHITE
    });

    // ── Serialize ─────────────────────────────────────────────────────────────
    const pdfBytes = await pdfDoc.save();

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/pdf' },
      body: Buffer.from(pdfBytes),
      isRaw: true
    };

  } catch (err) {
    context.log.error('generate-pdf error:', err);
    context.res = { status: 500, body: `PDF generatie mislukt: ${err.message}` };
  }
};
