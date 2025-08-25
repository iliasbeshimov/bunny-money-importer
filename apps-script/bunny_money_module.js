/*
  Bunny Money Importer Module (for runtime loading via Apps Script loader)

  Usage: Host this file (raw) on GitHub. In your bound Apps Script project, add a small
  loader that fetches this code and evals it, then calls BunnyMoney.onOpen() and
  BunnyMoney.addFromNumisBids(). See loader snippet in the setup instructions.
*/

var BunnyMoney = (function() {
  // Public API
  function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('Bunny Money')
      .addItem('Add from NumisBids URL', 'BM_addFromNumisBids')
      .addItem('Force reload module (5m cache)', 'BM_reloadModule')
      .addToUi();
  }

  function addFromNumisBids() {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.prompt('Add NumisBids lot', 'Paste the lot URL:', ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    const url = (resp.getResponseText() || '').trim();
    if (!/^https?:\/\/(www\.)?numisbids\.com\//i.test(url)) {
      ui.alert('Please paste a valid NumisBids URL.');
      return;
    }

    const html = UrlFetchApp.fetch(url, { followRedirects: true, muteHttpExceptions: true }).getContentText();
    const parsed = parseNumisBids(html, url);
    const enriched = enrichWithHeuristics(parsed);
    const finalData = maybeFillWithAI(enriched);

    const sheet = SpreadsheetApp.getActiveSheet();
    const headers = getHeaders(sheet);

    // Build row strictly for existing headers (no new columns)
    const imgs = Array.isArray(finalData.images) ? finalData.images : [];
    const map = {
      'Description': finalData.description || '',
      'Country / Region': finalData.region || '',
      'Years': finalData.years || '',
      'Value': finalData.value || '', // face value only for modern coins (numeric), else blank
      'Denomination': finalData.denomination || '',
      'Material': finalData.material || '',
      'Notes': finalData.notes || '',
      'Link': finalData.url || url
    };

    // Images only if the columns exist
    if (headers.indexOf('Image 1') !== -1) {
      map['Image 1'] = imgs[0] ? '=IMAGE("' + imgs[0] + '")' : '';
    }
    if (headers.indexOf('Image 2') !== -1) {
      map['Image 2'] = imgs[1] ? '=IMAGE("' + imgs[1] + '")' : '';
    }

    const row = headers.map(function(h) { return (h in map ? map[h] : ''); });
    sheet.appendRow(row);

    const newRow = sheet.getLastRow();
    const descIdx = headers.indexOf('Description');
    if (descIdx >= 0) sheet.getRange(newRow, descIdx + 1).setWrap(true);

    // Extra images -> add URLs as note on Image 2 cell (no new columns)
    if (imgs.length > 2 && headers.indexOf('Image 2') !== -1) {
      const col = headers.indexOf('Image 2') + 1;
      const r = sheet.getRange(newRow, col);
      const existingNote = r.getNote();
      const note = 'More images:\n' + imgs.slice(2).join('\n');
      r.setNote(existingNote ? (existingNote + '\n' + note) : note);
      r.setWrap(true);
    }

    ui.alert('Added', 'Row appended for: ' + (finalData.title || url), ui.ButtonSet.OK);
  }

  /* =================== Parsing (NumisBids/Noonans) =================== */

  function parseNumisBids(html, url) {
    const title = getMatch(html, /<title[^>]*>([\s\S]*?)<\/title>/i);
    const description = extractLotDescription(html) || decodeHtml(normalizeMultiline(extractMainBlock(html) || title || url)).trim();

    // Images (prefer og:image and media coin images; exclude logos/headers)
    const images = collectCoinImages(html);

    // Use combined text for extraction
    const full = (title + ' ' + description).replace(/\s+/g, ' ');

    // Denomination (expandable)
    const denomination = pickOne(full, /\b(Tetradrachm|Drachm|Denarius|Aureus|Sestertius|Hexas|Fals|Dirham|Obol|As|Antoninianus|Solidus|Stater|Penny|Pence|Euro|Cent|Dinar|Dollar|Franc|Mark|Pfennig|Yen|Yuan|Won|Rupee|Rial|Ringgit|Lira|Dinar|Dirham|Krone|Krona|Forint|Kuna|Zloty|Grosz|Peso|Rand|Real|Leu)\b/i);
    let material = detectMaterial(full);

    // Years like “130-3”, “480-461 BC”, “420 - 413 v. Chr.”
    let years = '';
    const yRange = description.match(/\b(ca\.?\s*)?(\d{2,4})\s*[-–]\s*(\d{1,4})\s*(BC|BCE|AD|CE|v\.?\s*Chr\.?)?/i);
    const ySingle = description.match(/\b(\d{3,4})\s*(BC|BCE|AD|CE|v\.?\s*Chr\.?)\b/i);
    if (yRange) {
      const prefix = yRange[1] ? 'ca. ' : '';
      years = prefix + yRange[2] + '-' + yRange[3] + (yRange[4] ? ' ' + yRange[4] : '');
    } else if (ySingle) {
      years = ySingle[1] + ' ' + ySingle[2];
    }

    // References (RIC, RSC, SNG, HGC) -> Notes
    const refs = Array.from(new Set((full.match(/\b(RIC\s*[^);,.]+|RSC\s*[^);,.]+|SNG\s*[^);,.]+|HGC\s*[^);,.]+)\b/ig) || [])));
    const notes = refs.length ? 'Refs: ' + refs.join('; ') : '';

    // Material fallback by denom (e.g., Denarius -> Silver)
    if (!material && /\bDenarius\b/i.test(denomination || '')) material = 'Silver';

    // Region from CAPS token if present; else blank (filled by heuristics/AI)
    let region = '';
    // Accept comma or period after REGION (e.g., "SICILY, Messana." or "SICILY. Messana.")
    const capToken = description.match(/\b([A-ZÄÖÜß]{3,}(?: [A-ZÄÖÜß]{3,})*)(?:,|\.)/);
    if (capToken) region = normalizeRegion(capToken[1]);

    return { title, url, description, region, years, denomination, material, notes, images };
  }

  /* =================== Heuristics & AI Backfill =================== */

  function enrichWithHeuristics(o) {
    const text = (o.title || '') + ' ' + (o.description || '');
    const ancient = /\b(BC|BCE|AD|CE)\b/i.test(o.years || '') ||
      /\b(AR|AE|AV|Denarius|Sestertius|Tetradrachm|Obol|Drachm|Aureus)\b/i.test(((o.denomination||'') + ' ' + (o.material||'') + ' ' + text));

    if (ancient && !o.region) o.region = 'Rome';

    // Normalize images to https
    if (Array.isArray(o.images)) {
      o.images = o.images.map(function(u) { return (u && u.startsWith('//') ? 'https:' + u : u); });
    } else {
      o.images = [];
    }

    // Modern face value / denomination normalization; ancient => value blank
    if (!ancient) {
      const fv = extractFaceValueAndUnit(text);
      if (fv) {
        o.value = fv.value;             // numeric only
        o.denomination = fv.unitCanon;  // normalized currency name
      }
    } else {
      o.value = '';
    }

    return o;
  }

  function maybeFillWithAI(parsed) {
    var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!key) return parsed; // optional
    try {
      var prompt = [
        'Task: Extract only accurate fields for a coin listing. If a field is uncertain or not clearly supported by the text, return an empty string for it.',
        'If ancient and modern country is not applicable, set region to the empire/polity (e.g., Rome).',
        'Return compact JSON with keys: region, years, denomination, material.',
        'Text follows:',
        parsed.title || '', parsed.description || ''
      ].join('\n');

      var body = {
        model: 'gpt-4o-mini',
        temperature: 0.0,
        messages: [
          { role: 'system', content: 'Extract accurate, minimal fields. If unsure, return empty strings. Reply with JSON only.' },
          { role: 'user', content: prompt }
        ]
      };
      var resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + key },
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      });
      var data = JSON.parse(resp.getContentText() || '{}');
      var content = (((data.choices || [])[0] || {}).message || {}).content || '{}';
      var ai = JSON.parse(safeJson(content));
      var out = Object.assign({}, parsed);
      ['region','years','denomination','material'].forEach(function(k){
        if (!out[k] && ai && typeof ai[k] === 'string') out[k] = ai[k].trim();
      });
      return out;
    } catch (e) {
      return parsed;
    }
  }

  /* =================== Helpers =================== */

  function extractFaceValueAndUnit(text) {
    if (!text) return null;
    var units = [
      'euro','cent','cents','penny','pence','pound','pounds','dollar','dollars','centavo','centavos',
      'centime','centimes','franc','francs','mark','marks','pfennig','yen','yuan','won','rupee','rupees',
      'rial','ringgit','lira','dinar','dirham','krona','krone','kroner','kronor','forint',
      'kuna','zloty','grosz','groszy','peso','pesos','rand','real','reais','leu','lei','øre','ore'
    ];
    var re = new RegExp('\\b(\\d+(?:\\.\\d+)?)\\s*(' + units.join('|') + ')\\b', 'i');
    var m = text.match(re);
    if (!m) return null;
    var value = m[1];
    var unitRaw = m[2].toLowerCase();
    return { value: value, unitCanon: normalizeCurrencyUnit(unitRaw, value) };
  }

  function normalizeCurrencyUnit(unit, valueStr) {
    var map = {
      euro: 'Euro',
      cent: 'Cent', cents: 'Cent',
      penny: 'Pence', pence: 'Pence',
      pound: 'Pound', pounds: 'Pound',
      dollar: 'Dollar', dollars: 'Dollar',
      centavo: 'Centavo', centavos: 'Centavo',
      centime: 'Centime', centimes: 'Centime',
      franc: 'Franc', francs: 'Franc',
      mark: 'Mark', marks: 'Mark', pfennig: 'Pfennig',
      yen: 'Yen', yuan: 'Yuan', won: 'Won',
      rupee: 'Rupee', rupees: 'Rupee',
      rial: 'Rial', ringgit: 'Ringgit', lira: 'Lira',
      dinar: 'Dinar', dirham: 'Dirham',
      krona: 'Krone', krone: 'Krone', kroner: 'Krone', kronor: 'Krone',
      forint: 'Forint', kuna: 'Kuna', zloty: 'Zloty',
      grosz: 'Grosz', groszy: 'Grosz',
      peso: 'Peso', pesos: 'Peso',
      rand: 'Rand', real: 'Real', reais: 'Real',
      leu: 'Leu', lei: 'Leu',
      'øre': 'Ore', ore: 'Ore'
    };
    return map[unit] || (unit.charAt(0).toUpperCase() + unit.slice(1));
  }

  function getHeaders(sheet) {
    var lastCol = sheet.getLastColumn();
    return lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  }

  function extractMainBlock(html) {
    var h = html;
    h = h.replace(/\s*<br\s*\/??\s*>\s*/ig, '\n')
         .replace(/\s*<\/p\s*>\s*/ig, '\n')
         .replace(/<li\s*>/ig, '\n• ')
         .replace(/<\/li\s*>/ig, '\n');
    var candidates = [
      /<div[^>]+class=["'][^"']*viewlottext[^"']*["'][^>]*>([\s\S]*?)<\/div>/i,
      /<div[^>]+class=["'][^"']*(lot-desc|description|lot_text|lotDescription)[^"']*["'][^>]*>([\s\S]*?)<\/div>/i,
      /<article[^>]*>([\s\S]*?)<\/article>/i
    ];
    for (var i=0;i<candidates.length;i++) {
      var m = h.match(candidates[i]);
      if (m) return stripTags(m[m.length - 1]);
    }
    var body = getMatch(h, /<body[^>]*>([\s\S]*?)<\/body>/i);
    return body ? stripTags(body) : '';
  }

  // Extract just the visible lot description text, skipping estimate and helper blocks.
  function extractLotDescription(html) {
    // Narrow to the .viewlottext block first
    var vm = html.match(/<div[^>]+class=["'][^"']*viewlottext[^"']*["'][^>]*>([\s\S]*?)<\/div>/i);
    var block = vm ? vm[1] : html;
    // Collect all <div class="description"> blocks within, excluding id="postbid" and id="watchnote"
    var out = [];
    var re = /<div([^>]+)class=["'][^"']*description[^"']*["'][^>]*>([\s\S]*?)<\/div>/ig;
    var m;
    while ((m = re.exec(block)) !== null) {
      var attrs = m[1] || '';
      if (/id=["'](postbid|watchnote)["']/i.test(attrs)) continue; // skip hidden/note blocks
      var inner = m[2];
      // Keep only substantial text (skip tiny helper divs)
      var text = stripTags(inner).trim();
      if (text) out.push(text);
    }
    if (out.length) {
      return decodeHtml(normalizeMultiline(out.join('\n\n'))).trim();
    }
    return '';
  }

  function collectCoinImages(html) {
    var out = [];
    // Prefer OpenGraph image first
    var og = getMeta(html, 'property', 'og:image');
    if (og) out.push(og);
    // Then only coin images hosted under media.numisbids.com/sales (exclude logos/headers)
    var imgRe = /<img[^>]+src=["']([^"']+)["'][^>]*>/ig;
    var m;
    while ((m = imgRe.exec(html)) !== null) {
      var u = m[1];
      if (!u) continue;
      if (/static\.numisbids\.com\/images\//i.test(u)) continue; // skip headers/logos/icons
      if (/logo|header|favicon|mstile|email\.png/i.test(u)) continue;
      if (/media\.numisbids\.com\/sales\//i.test(u)) out.push(u);
    }
    // de-dup & normalize protocol
    var uniq = Array.from(new Set(out));
    return uniq.map(function(u){ return u && u.startsWith('//') ? 'https:' + u : u; });
  }

  function detectMaterial(text) {
    if (!text) return '';
    if (/\b(AR|Silver)\b/i.test(text)) return 'Silver';
    if (/\b(AV|Gold)\b/i.test(text)) return 'Gold';
    if (/\b(AE|Bronze|Copper)\b/i.test(text)) return 'Copper';
    if (/Electrum/i.test(text)) return 'Electrum';
    if (/Billon/i.test(text)) return 'Billon';
    if (/Nickel/i.test(text)) return 'Nickel';
    return '';
  }

  function normalizeMultiline(s) {
    return s.replace(/\u00A0/g, ' ')
            .replace(/ +\n/g, '\n')
            .replace(/\n +/g, '\n')
            .replace(/\n{3,}/g, '\n\n')
            .trim();
  }

  function stripTags(s) {
    return s.replace(/<script[\s\S]*?<\/script>/ig, '')
            .replace(/<style[\s\S]*?<\/style>/ig, '')
            .replace(/<[^>]+>/g, '')
            .replace(/[\t\r]+/g, '');
  }

  function decodeHtml(s) {
    var map = { '&amp;':'&','&quot;':'"','&#39;':'\'','&lt;':'<','&gt;':'>','&nbsp;':' ' };
    return s.replace(/&(amp|quot|#39|lt|gt|nbsp);/g, function(m){ return map[m] || m; });
  }

  function getMeta(html, attr, name) {
    var re = new RegExp('<meta[^>]+' + attr + '=["\\\']' + escapeRegex(name) + '["\\\'][^>]+content=["\\\']([^"\\\']+)["\\\']', 'i');
    return getMatch(html, re);
  }

  function getMatch(s, re) {
    var m = s && s.match(re);
    return (m && m[1]) ? m[1].trim() : '';
  }

  function escapeRegex(s) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function pickOne(text, re, idx) {
    if (!text) return '';
    var m = text.match(re);
    if (!m) return '';
    var v = (idx ? m[idx] : (m[1] || m[0])) + '';
    return v.trim().replace(/\s+/g, ' ');
  }

  function normalizeRegion(t) {
    var x = t.trim().toUpperCase();
    var map = { 'SICILY': 'Sicily', 'SIZILIEN': 'Sicily', 'ROME': 'Rome', 'ROMA': 'Rome', 'BRUTTIUM': 'BRUTTIUM', 'MONGOL': 'Mongol' };
    return map[x] || (x.charAt(0) + x.slice(1).toLowerCase());
  }

  function safeJson(s) {
    var m = s && s.match(/\{[\s\S]*\}/);
    return m ? m[0] : '{}';
  }

  // Expose public API
  return { onOpen: onOpen, addFromNumisBids: addFromNumisBids };
})();
