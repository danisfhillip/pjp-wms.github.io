const LOG_SPREADSHEET_ID='10vs1ToqXUx08HL74QdQVSpm189C_qMEGym8tn1jCAqo';

function getLogSpreadsheet(){
  return SpreadsheetApp.openById(LOG_SPREADSHEET_ID);
}

function doGet(e){
  const page=e&&e.parameter&&e.parameter.page?e.parameter.page:'home';
  let file='index';
  let title='Warehouse System';

  if(page==='login'){file='login';title='Login';
  }else if(page==='manage_user'){file='manage_user';title='Manage User';
  }else if(page==='scanner'){file='scanner';title='Stock Opname Roll';
  }else if(page==='barcode'){file='barcode';title='Print Barcode';
  }else if(page==='receiving'){file='receiving';title='Barang Masuk';
  }else if(page==='location'){file='location';title='QR Location';
  }else if(page==='movement'){file='movement';title='Pergerakan Barang';
  }else if(page==='sample_cut'){file='sample_cut';title='Potong Sampel';
  }else if(page==='stock'){file='stock';title='Cek Stok';
  }else if(page==='dashboard'){file='dashboard';title='Dashboard Gudang';
  }else if(page==='dashboard_detail'){file='dashboard_detail';title='Detail Roll';
  }else if(page==='outbound'){file='outbound';title='Barang Keluar';
  }else if(page==='location_qr_print'){
    return HtmlService
      .createTemplateFromFile('location_qr_print')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('QR Location')
      .addMetaTag('viewport','width=device-width,initial-scale=1.0');
  }

  const template=HtmlService.createTemplateFromFile(file);
  template.pageTitle=title;
  template.detailType=e&&e.parameter&&e.parameter.type?e.parameter.type:'available';
  template.webAppUrl=ScriptApp.getService().getUrl();

  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle(title)
    .addMetaTag('viewport','width=device-width, initial-scale=1.0');
}

function include(filename){
  const content=HtmlService.createHtmlOutputFromFile(filename).getContent();
  const webAppUrl=JSON.stringify(ScriptApp.getService().getUrl());
  return content.replace(/__WEB_APP_URL__/g,webAppUrl);
}

function getWebAppUrl(){
  return ScriptApp.getService().getUrl();
}

function getCurrentUserEmail(){
  return Session.getActiveUser().getEmail()||'UNKNOWN';
}

function testCurrentUser(){
  return {
    activeUser:Session.getActiveUser().getEmail()||'UNKNOWN',
    effectiveUser:Session.getEffectiveUser().getEmail()||'UNKNOWN'
  };
}

/*** ================================* MENCARI ROLL* ================================ ***/
function getRoll(rollId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ROLL');

    if (!sheet) {
      return {
        success: false,
        message: 'Sheet ROLL tidak ditemukan.'
      };
    }

    rollId = String(rollId).trim().toLowerCase();

    if (!rollId) {
      return {
        success: false,
        message: 'Roll ID kosong.'
      };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return {
        success: false,
        message: 'Sheet ROLL belum memiliki data.'
      };
    }

    const headers = data[0].map(function(header) {
      return String(header).trim();
    });

    const rollIndex = headers.findIndex(function(header) {
      return header.toLowerCase() === 'roll id';
    });

    const skuIndex = headers.findIndex(function(header) {
      return header.toLowerCase() === 'sku';
    });

    const colorIndex = headers.findIndex(function(header) {
      return header.toLowerCase() === 'color';
    });

    const qtyIndex = headers.findIndex(function(header) {
      return header.toLowerCase() === 'qty';
    });

    const locationIndex = headers.findIndex(function(header) {
      return header.toLowerCase() === 'location id';
    });

    if (rollIndex === -1) {
      return {
        success: false,
        message: 'Kolom "Roll ID" tidak ditemukan.'
      };
    }

    for (let i = 1; i < data.length; i++) {
      const currentRoll = String(data[i][rollIndex])
        .trim()
        .toLowerCase();

      if (currentRoll === rollId) {
        return {
          success: true,
          data: {
            'Roll ID': String(data[i][rollIndex]),
            'SKU': skuIndex !== -1
              ? formatCellValue(data[i][skuIndex])
              : '',
            'Color': colorIndex !== -1
              ? String(data[i][colorIndex])
              : '',
            'Qty': qtyIndex !== -1
              ? String(data[i][qtyIndex])
              : '',
            'Location ID': locationIndex !== -1
              ? String(data[i][locationIndex])
              : ''
          }
        };
      }
    }

    return {
      success: false,
      message: 'Roll "' + rollId + '" tidak ditemukan.'
    };

  } catch (error) {
    return {
      success: false,
      message: 'Error getRoll: ' + error.message
    };
  }
}

/*** Format nilai cell ***/
function formatCellValue(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  return String(value ?? '');
}

/*** ================================* CEK RAK* ================================*/
function getRack(locationId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('RACK');

    if (!sheet) {
      return {
        success: false,
        message: 'Sheet RACK tidak ditemukan.'
      };
    }

    locationId = String(locationId)
      .trim()
      .toLowerCase();

    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return {
        success: false,
        message: 'Sheet RACK belum memiliki data.'
      };
    }

    const headers = data[0].map(function(header) {
      return String(header).trim();
    });

    const locationIndex = headers.findIndex(function(header) {
      return header.toLowerCase() === 'location id';
    });

    if (locationIndex === -1) {
      return {
        success: false,
        message: 'Kolom "Location ID" tidak ditemukan.'
      };
    }

    for (let i = 1; i < data.length; i++) {
      const currentLocation = String(data[i][locationIndex])
        .trim()
        .toLowerCase();

      if (currentLocation === locationId) {
        const result = {};

        headers.forEach(function(header, index) {
          result[header] = data[i][index];
        });

        return {
          success: true,
          data: result
        };
      }
    }

    return {
      success: false,
      message: 'Rak "' + locationId + '" tidak ditemukan.'
    };

  } catch (error) {
    return {
      success: false,
      message: 'Error getRack: ' + error.message
    };
  }
}

/*** ================================* PINDAHKAN ROLL* ================================*/
function moveRoll(rollId, newLocation, operator) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const rollSheet = ss.getSheetByName('ROLL');
    const rackSheet = ss.getSheetByName('RACK');
    const logSheet = ss.getSheetByName('MOVEMENT_LOG');

    if (!rollSheet) {
      return {
        success: false,
        message: 'Sheet ROLL tidak ditemukan.'
      };
    }

    if (!rackSheet) {
      return {
        success: false,
        message: 'Sheet RACK tidak ditemukan.'
      };
    }

    if (!logSheet) {
      return {
        success: false,
        message: 'Sheet MOVEMENT_LOG tidak ditemukan.'
      };
    }

    rollId = String(rollId).trim();
    newLocation = String(newLocation).trim();
    operator = String(operator || 'Unknown').trim();

    if (!rollId) {
      return {
        success: false,
        message: 'Roll ID kosong.'
      };
    }

    if (!newLocation) {
      return {
        success: false,
        message: 'Lokasi tujuan kosong.'
      };
    }

    /**
     * CARI ROLL
     */
    const rollData = rollSheet.getDataRange().getValues();

    const rollHeaders = rollData[0].map(function(header) {
      return String(header).trim();
    });

    const rollIdIndex = rollHeaders.findIndex(function(header) {
      return header.toLowerCase() === 'roll id';
    });

    const locationIndex = rollHeaders.findIndex(function(header) {
      return header.toLowerCase() === 'location id';
    });

    if (rollIdIndex === -1 || locationIndex === -1) {
      return {
        success: false,
        message: 'Kolom Roll ID atau Location ID tidak ditemukan.'
      };
    }

    let rollRow = -1;
    let oldLocation = '';

    for (let i = 1; i < rollData.length; i++) {
      const currentRoll = String(
        rollData[i][rollIdIndex]
      ).trim();

      if (
        currentRoll.toLowerCase() ===
        rollId.toLowerCase()
      ) {
        rollRow = i + 1;

        oldLocation = String(
          rollData[i][locationIndex]
        ).trim();

        break;
      }
    }

    if (rollRow === -1) {
      return {
        success: false,
        message: 'Roll "' + rollId + '" tidak ditemukan.'
      };
    }

    /*** CEK RAK TUJUAN ***/
    const rackData = rackSheet.getDataRange().getValues();

    const rackHeaders = rackData[0].map(function(header) {
      return String(header).trim();
    });

    const rackLocationIndex = rackHeaders.findIndex(function(header) {
      return header.toLowerCase() === 'location id';
    });

    if (rackLocationIndex === -1) {
      return {
        success: false,
        message: 'Kolom Location ID pada RACK tidak ditemukan.'
      };
    }

    let rackRow = -1;

    for (let i = 1; i < rackData.length; i++) {
      const currentLocation = String(
        rackData[i][rackLocationIndex]
      ).trim();

      if (
        currentLocation.toLowerCase() ===
        newLocation.toLowerCase()
      ) {
        rackRow = i + 1;
        break;
      }
    }

    if (rackRow === -1) {
      return {
        success: false,
        message: 'Rak "' + newLocation + '" tidak ditemukan.'
      };
    }

    /*** UPDATE LOCATION ROLL ***/
    rollSheet
      .getRange(
        rollRow,
        locationIndex + 1
      )
      .setValue(newLocation);

    /*** UPDATE CURRENT ROLL RAK ***/
    updateRackCount(
      rackSheet,
      rackData,
      rackHeaders,
      oldLocation,
      newLocation
    );

    /*** CATAT MOVEMENT LOG ***/
    logSheet.appendRow([
      new Date(),
      rollId,
      oldLocation,
      newLocation,
      operator
    ]);

    return {
      success: true,
      message: 'Roll berhasil dipindahkan.',
      rollId: rollId,
      from: oldLocation,
      to: newLocation
    };

  } catch (error) {
    return {
      success: false,
      message: 'Error moveRoll: ' + error.message
    };
  }
}

/*** ================================ * UPDATE JUMLAH ROLL DI RAK * ================================ ***/
function updateRackCount(
  sheet,
  data,
  headers,
  oldLocation,
  newLocation
) {
  const locationIndex = headers.findIndex(function(header) {
    return header.toLowerCase() === 'location id';
  });

  const currentRollIndex = headers.findIndex(function(header) {
    return header.toLowerCase() === 'current roll';
  });

  if (
    locationIndex === -1 ||
    currentRollIndex === -1
  ) {
    return;
  }

  for (let i = 1; i < data.length; i++) {
    const location = String(
      data[i][locationIndex]
    ).trim();

    /*** KURANGI RAK LAMA ***/
    if (
      location.toLowerCase() ===
      String(oldLocation)
        .trim()
        .toLowerCase()
    ) {
      let current = Number(
        data[i][currentRollIndex]
      ) || 0;

      current = Math.max(0, current - 1);

      sheet
        .getRange(
          i + 1,
          currentRollIndex + 1
        )
        .setValue(current);
    }

    /*** TAMBAH RAK BARU ***/
    if (
      location.toLowerCase() ===
      String(newLocation)
        .trim()
        .toLowerCase()
    ) {
      let current = Number(
        data[i][currentRollIndex]
      ) || 0;

      current = current + 1;

      sheet
        .getRange(
          i + 1,
          currentRollIndex + 1
        )
        .setValue(current);
    }
  }
}

/*** ================================ * TEST * ================================ ***/
function testConnection() {
  Logger.log("KONEKSI BERHASIL");
}

function testSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log(
    "Spreadsheet: " + ss.getName()
  );

  const sheet = ss.getSheetByName("ROLL");

  if (!sheet) {
    Logger.log("SHEET ROLL TIDAK DITEMUKAN");
    return;
  }

  Logger.log("SHEET ROLL DITEMUKAN");

  const data = sheet.getDataRange().getValues();

  Logger.log(
    "Jumlah baris: " + data.length
  );
}

function testRoll() {
  const result = getRoll("ROLL-001");

  Logger.log(
    JSON.stringify(result)
  );
}

function ping() {
  return {
    success: true,
    message: "Web App berhasil terhubung ke Apps Script."
  };
}

function generateLocationMaster(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('LOCATION_MASTER');

  if(!sheet) sheet=ss.insertSheet('LOCATION_MASTER');

  sheet.clearContents();

  const rows=[['Location Type','Rack','Posisi','Level','Location ID']];
  const racks=[
    {name:'Rack A',positions:8},
    {name:'Rack B',positions:8},
    {name:'Rack C',positions:8},
    {name:'Rack D',positions:10},
    {name:'Rack E',positions:10}
  ];
  const levels=['A','B','C'];

  for(let r=0;r<racks.length;r++){
    const rack=racks[r];

    for(let pos=1;pos<=rack.positions;pos++){
      const position=String(pos).padStart(2,'0');

      for(let l=0;l<levels.length;l++){
        const level=levels[l];

        rows.push([
          'Rak',
          rack.name,
          position,
          level,
          rack.name+' '+position+'-'+level
        ]);
      }
    }
  }

  sheet.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  sheet.getRange(1,1,1,5).setFontWeight('bold');
  sheet.autoResizeColumns(1,5);
}

function getLocationMaster(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('LOCATION_MASTER');

  if(!sheet) return [];

  const data=sheet.getDataRange().getValues();
  if(data.length<2) return [];

  return data.slice(1).map(function(row){
    return {
      type:String(row[0]).trim(),
      rack:String(row[1]).trim(),
      position:String(row[2]).trim(),
      level:String(row[3]).trim(),
      id:String(row[4]).trim()
    };
  });
}


/*** QR CODE LOCATION / RAK / PALLET ***/
function getLocationMaster(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('LOCATION_MASTER');
  if(!sheet)return [];
  const data=sheet.getDataRange().getValues();
  if(data.length<2)return [];
  return data.slice(1).map(function(row){
    return {
      type:String(row[0]).trim(),
      rack:String(row[1]).trim(),
      position:String(row[2]).trim(),
      level:String(row[3]).trim(),
      id:String(row[4]).trim()
    };
  });
}
