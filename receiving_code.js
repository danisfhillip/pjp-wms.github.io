
function createRolls(
  sku, category, fabricType, design, product,
  colors, width, gsm, location, token
) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const session = getLoginSession(token);

    if (!session.success) {
      return {
        success: false,
        message: 'Session login tidak valid. Silakan login kembali.'
      };
    }

    const user = session.userId;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSS = getLogSpreadsheet();
    const sheet = ss.getSheetByName('ROLL');
    const logSheet = logSS.getSheetByName('RECEIVING_LOG');

    if (!sheet) {
      return { success: false, message: 'Sheet ROLL tidak ditemukan.' };
    }

    if (!logSheet) {
      return {
        success: false,
        message: 'Sheet RECEIVING_LOG tidak ditemukan.'
      };
    }

    const categoryCodes = {
      'Dyeing': '110508',
      'Printing': '110509',
      'Sublim (DGP)': '110510',
      'Putihan': '110507'
    };

    if (!categoryCodes[category]) {
      return { success: false, message: 'Kategori tidak valid.' };
    }

    if (!fabricType || !product || !width || Number(width) <= 0) {
      return {
        success: false,
        message: 'Data jenis kain, nama produk, dan lebar wajib diisi.'
      };
    }

    if (!Array.isArray(colors) || colors.length === 0) {
      return {
        success: false,
        message: 'Tambahkan minimal satu warna.'
      };
    }

    let totalQuantity = 0;

    for (let i = 0; i < colors.length; i++) {
      const color = String(colors[i].color || '').trim();
      const quantity = Number(colors[i].quantity);

      if (
        !color ||
        !Number.isInteger(quantity) ||
        quantity < 1
      ) {
        return {
          success: false,
          message: 'Data warna atau jumlah roll tidak valid pada baris ' + (i + 1) + '.'
        };
      }

      totalQuantity += quantity;
    }

    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(function(header) {
        return String(header).trim();
      });

    const getIndex = function(name) {
      return headers.indexOf(name);
    };

    const rollIdIndex = getIndex('Roll ID');
    const skuIndex = getIndex('SKU');
    const categoryIndex = getIndex('Kategori');
    const fabricTypeIndex = getIndex('Jenis Kain');
    const designIndex = getIndex('Kode Desain');
    const productIndex = getIndex('Nama Produk');
    const colorIndex = getIndex('Warna');
    const widthIndex = getIndex('Lebar');
    const gsmIndex = getIndex('GSM');
    const locationIndex = getIndex('Location ID');

    const requiredColumns = [
      ['Roll ID', rollIdIndex],
      ['SKU', skuIndex],
      ['Kategori', categoryIndex],
      ['Jenis Kain', fabricTypeIndex],
      ['Kode Desain', designIndex],
      ['Nama Produk', productIndex],
      ['Warna', colorIndex],
      ['Lebar', widthIndex],
      ['GSM', gsmIndex],
      ['Location ID', locationIndex]
    ];

    for (let i = 0; i < requiredColumns.length; i++) {
      if (requiredColumns[i][1] === -1) {
        return {
          success: false,
          message: 'Kolom "' + requiredColumns[i][0] + '" tidak ditemukan di sheet ROLL.'
        };
      }
    }

    const categoryCode = categoryCodes[category];
    const dateCode = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyMMdd'
    );

    const prefix = categoryCode + '-' + dateCode + '-';
    const existingData = sheet.getDataRange().getValues();
    let maxNumber = 0;

    for (let i = 1; i < existingData.length; i++) {
      const rollId = String(
        existingData[i][rollIdIndex] || ''
      ).trim();

      if (rollId.indexOf(prefix) === 0) {
        const sequence = rollId.substring(prefix.length);
        const number = Number(sequence);

        if (
          /^\d{4}$/.test(sequence) &&
          Number.isInteger(number) &&
          number > maxNumber
        ) {
          maxNumber = number;
        }
      }
    }

    if (maxNumber + totalQuantity > 9999) {
      return {
        success: false,
        message: 'Nomor Roll sudah melebihi batas 9999 untuk kategori dan tanggal ini.'
      };
    }

    const rows = [];
    const rollIds = [];
    const logs = [];
    const now = new Date();

    colors.forEach(function(item) {
      const color = String(item.color).trim();
      const quantity = Number(item.quantity);

      for (let i = 1; i <= quantity; i++) {
        const rollNumber = String(maxNumber + 1).padStart(4, '0');
        const rollId = prefix + rollNumber;
        maxNumber++;

        const row = new Array(headers.length).fill('');

        row[rollIdIndex] = rollId;
        row[skuIndex] = sku;
        row[categoryIndex] = category;
        row[fabricTypeIndex] = fabricType;
        row[designIndex] = design;
        row[productIndex] = product;
        row[colorIndex] = color;
        row[widthIndex] = width;
        row[gsmIndex] = gsm;
        row[locationIndex] = location || '';

        rows.push(row);

        rollIds.push(rollId);

        logs.push([
          now, rollId, sku, category, fabricType,
          design, product, color, width, gsm, user
        ]);
      }
    });

    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      headers.length
    ).setValues(rows);

    logSheet.getRange(
      logSheet.getLastRow() + 1,
      1,
      logs.length,
      11
    ).setValues(logs);

    logSheet
      .getRange(logSheet.getLastRow() - logs.length + 1, 1, logs.length, 1)
      .setNumberFormat('MM/dd/yyyy HH:mm');

    SpreadsheetApp.flush();

    return {
      success: true,
      rollIds: rollIds
    };

  } finally {
    lock.releaseLock();
  }
}
