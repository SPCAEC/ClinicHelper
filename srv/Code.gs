/** Clinic Label Helper — Backend (v0.1)
 *
 * - Dashboard + Generate Labels API
 * - For each selected sheet row:
 *    • build placeholder map from CSV data
 *    • clone Owner + Pet Slides templates
 *    • replace {{placeholders}}
 *    • export as PDF (1 page each)
 *    • send 4 PDFs (owner x2, pet x2) to Render merge service (base64)
 *    • save merged PDF in output folder
 *    • write EventDate + EventLocation back to sheet
 */


/** Web app entry */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  return t.evaluate()
    .setTitle(CFG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/** HTML Service dev helper (optional) */
function _testGenerateSingle() {
  const payload = {
    eventDate: '2025-11-22',
    eventLocation: 'SPCA Lipsey Clinic',
    startRow: 2,
    endRow: 5,
  };
  return apiGenerateLabels(payload);
}


/** Public API: called from js.app.html */
function apiGenerateLabels(payload) {
  try {
    const eventDate = String(payload.eventDate || '').trim();
    const eventLocation = String(payload.eventLocation || '').trim();
    const startRow = Number(payload.startRow);
    const endRow = Number(payload.endRow);

    if (!eventDate || !eventLocation) {
      throw new Error('Event Date and Location are required.');
    }
    if (!startRow || !endRow || startRow < 2 || endRow < startRow) {
      throw new Error('Please provide a valid row range (row 2 or higher).');
    }

    const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
    const sheet = ss.getSheetByName(CFG.SHEET_NAME);
    if (!sheet) throw new Error('Sheet not found: ' + CFG.SHEET_NAME);

    const lastCol = sheet.getLastColumn();
    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const colIndex = buildColumnIndex_(headerRow);

    // Ensure EventDate & EventLocation columns exist
    const ensured = ensureEventColumns_(sheet, headerRow, colIndex);
    const idxEventDate = ensured.idxEventDate;
    const idxEventLocation = ensured.idxEventLocation;

    const numRows = endRow - startRow + 1;
    const rows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();

    let successCount = 0;
    const errors = [];

    rows.forEach((rowValues, i) => {
      const rowNum = startRow + i;
      try {
        const placeholders = buildPlaceholderMap_(rowValues, colIndex);
        const outputName = buildOutputFileName_(placeholders);

        // Create merged 4-page label PDF
        const mergedFile = createMergedLabelPdfForRow_(
          placeholders,
          outputName
        );

        // Write date + location back to sheet
        const rowRange = sheet.getRange(rowNum, 1, 1, lastCol);
        const rowArr = rowRange.getValues()[0];

        rowArr[idxEventDate] = eventDate;
        rowArr[idxEventLocation] = eventLocation;
        rowRange.setValues([rowArr]);

        successCount++;
      } catch (err) {
        Logger.log('Row ' + rowNum + ' failed: ' + err);
        errors.push({ row: rowNum, message: String(err) });
      }
    });

    return {
      ok: true,
      successCount,
      errorCount: errors.length,
      errors,
    };

  } catch (err) {
    Logger.log('apiGenerateLabels error: ' + err);
    return {
      ok: false,
      error: String(err),
    };
  }
}


/** Build a map of header → index (0-based). */
function buildColumnIndex_(headerRow) {
  const map = {};
  headerRow.forEach((name, i) => {
    if (!name) return;
    map[String(name).trim()] = i;
  });
  return map;
}


/**
 * Ensure EventDate and EventLocation columns exist.
 * Returns 0-based indexes (within the row array).
 */
function ensureEventColumns_(sheet, headerRow, colIndex) {
  const lastCol = headerRow.length;
  let idxEventDate = colIndex[CFG.COLS.EVENT_DATE];
  let idxEventLocation = colIndex[CFG.COLS.EVENT_LOCATION];

  let currentLastCol = lastCol;
  const updates = headerRow.slice();

  if (typeof idxEventDate !== 'number') {
    updates[currentLastCol] = CFG.COLS.EVENT_DATE;
    idxEventDate = currentLastCol;
    currentLastCol++;
  }

  if (typeof idxEventLocation !== 'number') {
    updates[currentLastCol] = CFG.COLS.EVENT_LOCATION;
    idxEventLocation = currentLastCol;
    currentLastCol++;
  }

  if (currentLastCol > lastCol) {
    sheet.getRange(1, 1, 1, currentLastCol).setValues([updates]);
  }

  return { idxEventDate, idxEventLocation };
}


/** Safely get cell by column name using map. */
function getVal_(rowValues, colIndex, name) {
  const idx = colIndex[name];
  if (typeof idx !== 'number') return '';
  const v = rowValues[idx];
  return v == null ? '' : String(v);
}


/** Build placeholder map from a CSV row. */
function buildPlaceholderMap_(rowValues, colIndex) {
  const firstName = getVal_(rowValues, colIndex, CFG.COLS.FIRST_NAME);
  const lastName = getVal_(rowValues, colIndex, CFG.COLS.LAST_NAME);
  const addr1 = getVal_(rowValues, colIndex, CFG.COLS.ADDRESS1);
  const addr2 = getVal_(rowValues, colIndex, CFG.COLS.ADDRESS2);
  const city = getVal_(rowValues, colIndex, CFG.COLS.CITY);
  const state = getVal_(rowValues, colIndex, CFG.COLS.STATE);
  const zip = getVal_(rowValues, colIndex, CFG.COLS.ZIP);
  const phone = getVal_(rowValues, colIndex, CFG.COLS.PHONE);
  const email = getVal_(rowValues, colIndex, CFG.COLS.EMAIL);
  const petName = getVal_(rowValues, colIndex, CFG.COLS.PET_NAME);
  const species = getVal_(rowValues, colIndex, CFG.COLS.SPECIES);
  const age = getVal_(rowValues, colIndex, CFG.COLS.AGE);
  const breed = getVal_(rowValues, colIndex, CFG.COLS.BREED);
  const color = getVal_(rowValues, colIndex, CFG.COLS.COLOR);
  const sex = getVal_(rowValues, colIndex, CFG.COLS.SEX);
  const snRaw = getVal_(rowValues, colIndex, CFG.COLS.SN);
  const reason = getVal_(rowValues, colIndex, CFG.COLS.REASON);

  const ownerName = joinNonEmpty_([firstName, lastName], ' ');
  const streetAddress = joinNonEmpty_([addr1, addr2], ' ');
  const cityStateZip = buildCityStateZip_(city, state, zip);
  const snCode = mapSNtoIntactCode_(snRaw); // Y/N/U for "Intact?"

  return {
    '{{OwnerName}}': ownerName,
    '{{StreetAddress}}': streetAddress,
    '{{CityStateZip}}': cityStateZip,
    '{{PhoneNumber}}': phone,
    '{{Email}}': email,
    '{{PetName}}': petName,
    '{{Species}}': species,
    '{{Age}}': age,
    '{{Breed}}': breed,
    '{{Color}}': color,
    '{{Sex}}': sex,
    '{{SN}}': snCode,
    '{{Reason}}': reason,
    // Keep raw components handy for naming
    _firstName: firstName,
    _lastName: lastName,
    _petName: petName,
  };
}


/** Join non-empty parts with a separator. */
function joinNonEmpty_(parts, sep) {
  const cleaned = parts
    .map(p => (p == null ? '' : String(p).trim()))
    .filter(Boolean);
  return cleaned.join(sep);
}


function buildCityStateZip_(city, state, zip) {
  const c = city ? String(city).trim() : '';
  const s = state ? String(state).trim() : '';
  const z = zip ? String(zip).trim() : '';

  let firstPart = '';
  if (c && s) firstPart = c + ', ' + s;
  else if (c) firstPart = c;
  else if (s) firstPart = s;

  if (firstPart && z) return firstPart + ' ' + z;
  if (!firstPart && z) return z;
  return firstPart;
}


/**
 * Map sheet's "Spayed or Neutered?" answer → label value for "Intact"
 * Sheet values: Yes / No / Unknown
 * Label expects "Intact" (Y/N/U)
 *  - Yes  (spayed/neutered) ⇒ Intact? No ⇒ 'N'
 *  - No   (not altered)     ⇒ Intact? Yes ⇒ 'Y'
 *  - Unknown                ⇒ 'U'
 */
function mapSNtoIntactCode_(snRaw) {
  const val = String(snRaw || '').trim().toLowerCase();
  if (!val) return 'U';

  if (val === 'yes') return 'N';
  if (val === 'no') return 'Y';
  if (val === 'unknown') return 'U';

  // Support leading letters
  if (val[0] === 'y') return 'N';
  if (val[0] === 'n') return 'Y';
  if (val[0] === 'u') return 'U';

  return 'U';
}


/** Build output file name: LastName_FirstName_PetName_ClinicLabel.pdf */
function buildOutputFileName_(placeholders) {
  const lastName = (placeholders._lastName || '').trim();
  const firstName = (placeholders._firstName || '').trim();
  const petName = (placeholders._petName || '').trim();

  const segments = [lastName, firstName, petName]
    .map(s => s.replace(/[^\w\-]+/g, ''))
    .filter(Boolean);

  const base = segments.join('_') || 'ClinicLabel';
  return base + '_ClinicLabel.pdf';
}


/**
 * For a single row:
 *  - Clone Owner and Pet templates
 *  - Replace placeholders
 *  - Export PDFs
 *  - Send 4 PDFs (Owner x2, Pet x2) to Render merge service (base64)
 *  - Save merged PDF in output folder
 */
function createMergedLabelPdfForRow_(placeholders, outputName) {
  const ownerTemplateFile = DriveApp.getFileById(CFG.OWNER_TEMPLATE_ID);
  const petTemplateFile = DriveApp.getFileById(CFG.PET_TEMPLATE_ID);

  const ownerCopy = ownerTemplateFile.makeCopy('Owner Label TMP');
  const petCopy = petTemplateFile.makeCopy('Pet Label TMP');

  try {
    const ownerPres = SlidesApp.openById(ownerCopy.getId());
    const petPres = SlidesApp.openById(petCopy.getId());

    applyPlaceholdersToPresentation_(ownerPres, placeholders);
    applyPlaceholdersToPresentation_(petPres, placeholders);

    const ownerBlob = ownerCopy.getAs(MimeType.PDF).setName('owner.pdf');
    const petBlob = petCopy.getAs(MimeType.PDF).setName('pet.pdf');

    // We want 4 pages: Owner x2, Pet x2
    const blobs = [ownerBlob, ownerBlob, petBlob, petBlob];

    const mergedInfo = mergePdfsViaRender_(blobs, outputName);
    return mergedInfo;

  } finally {
    // Clean up temporary Slides
    try { ownerCopy.setTrashed(true); } catch (e) {}
    try { petCopy.setTrashed(true); } catch (e) {}
  }
}


/** Replace {{placeholders}} in all text boxes for a Slides presentation. */
function applyPlaceholdersToPresentation_(presentation, placeholders) {
  Object.keys(placeholders).forEach(key => {
    if (!key.startsWith('{{')) return; // ignore internal _fields
    const value = placeholders[key] || '';
    presentation.replaceAllText(key, value);
  });
}


/**
 * Merge PDFs via the same Render service used by Transportation Helper.
 *
 * pdfBlobs: array of Blob objects (owner x2, pet x2)
 * outputName: final merged filename
 *
 * Request:
 *  {
 *    outputName: "file.pdf",
 *    files: [
 *      { name: "part1.pdf", contentBase64: "..." },
 *      ...
 *    ]
 *  }
 *
 * Response (JSON):
 *  { contentBase64, fileName, fileUrl }   // same shape as mergePDFs_
 */
function mergePdfsViaRender_(pdfBlobs, outputName) {
  const url = CFG.MERGE_SERVICE_URL || 'https://pdf-merge-service.onrender.com/merge';
  Logger.log('Merging %s PDFs via %s', pdfBlobs.length, url);

  const files = pdfBlobs.map((blob, i) => {
    try {
      const bytes = blob.getBytes();
      if (!bytes || !bytes.length) {
        Logger.log('⚠️ Skipping invalid or empty blob for part %s', i + 1);
        return null;
      }
      const base64 = Utilities.base64Encode(bytes);
      if (!base64) {
        Logger.log('⚠️ Skipping blob with missing base64 data for part %s', i + 1);
        return null;
      }
      return {
        name: blob.getName() || ('part' + (i + 1) + '.pdf'),
        contentBase64: base64
      };
    } catch (err) {
      Logger.log('⚠️ Error reading blob for part %s: %s', i + 1, err);
      return null;
    }
  }).filter(Boolean);

  if (!files.length) {
    throw new Error('No valid PDFs to merge.');
  }

  const payload = JSON.stringify({
    outputName: outputName || 'merged.pdf',
    files
  });
  Logger.log('Payload prepared for merge: %s files', files.length);

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  Logger.log('Merge response %s: %s', code, text.substring(0, Math.min(500, text.length)));

  if (code < 200 || code >= 300) {
    throw new Error('Merge API error: ' + code + ' — ' + text);
  }

  let merged;
  try {
    merged = JSON.parse(text);
  } catch (err) {
    throw new Error('Invalid JSON from merge API: ' + text.slice(0, 300));
  }

  // Same response handling pattern as Transportation Helper
  const base64Out =
    merged.contentBase64 ||
    (merged.merged && merged.merged.contentBase64) || // just in case
    null;

  const fileName = merged.fileName || outputName || 'merged.pdf';
  let fileUrl = merged.fileUrl || null;

  const folder = DriveApp.getFolderById(CFG.OUTPUT_FOLDER_ID);
  let file = null;

  if (base64Out) {
    const bytesOut = Utilities.base64Decode(base64Out);
    const blobOut = Utilities.newBlob(bytesOut, MimeType.PDF, fileName);
    file = folder.createFile(blobOut).setName(fileName);
    fileUrl = file.getUrl();
  } else if (!fileUrl) {
    throw new Error('Merge service response did not contain contentBase64 or fileUrl.');
  }

  return {
    fileId: file ? file.getId() : null,
    url: fileUrl,
    name: fileName
  };
}