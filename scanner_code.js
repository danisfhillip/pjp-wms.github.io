// function doGet() {
//   return HtmlService.createHtmlOutputFromFile('scanner')
//     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
//     .setTitle('Stock Opname Roll')
//     .addMetaTag('viewport','width=device-width, initial-scale=1.0');
// }

function getRollForOpname(rollId){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName('ROLL');

    if(!sheet){
      return {
        success:false,
        message:'Sheet ROLL tidak ditemukan.'
      };
    }

    rollId=String(rollId||'').trim();

    if(!rollId){
      return {
        success:false,
        message:'Barcode kosong.'
      };
    }

    const lastRow=sheet.getLastRow();
    const lastCol=sheet.getLastColumn();

    if(lastRow<2){
      return {
        success:false,
        message:'Sheet ROLL belum memiliki data.'
      };
    }

    // Ambil header sekali
    const headers=sheet
      .getRange(1,1,1,lastCol)
      .getValues()[0]
      .map(function(header){
        return String(header).trim();
      });

    function getIndex(name){
      return headers.findIndex(function(header){
        return header.toLowerCase()===name.toLowerCase();
      });
    }

    const rollIndex=getIndex('Roll ID');

    if(rollIndex===-1){
      return {
        success:false,
        message:'Kolom "Roll ID" tidak ditemukan.'
      };
    }

    // Cari Roll ID langsung menggunakan TextFinder
    const found=sheet
      .getRange(2,rollIndex+1,lastRow-1,1)
      .createTextFinder(rollId)
      .matchEntireCell(true)
      .matchCase(false)
      .findNext();

    if(!found){
      return {
        success:false,
        message:'Barcode "'+rollId+'" tidak ditemukan.'
      };
    }

    // Ambil hanya row yang ditemukan
    const foundRow=found.getRow();
    const row=sheet
      .getRange(foundRow,1,1,lastCol)
      .getValues()[0];

    const meterIndex=getIndex('Meter');
    const yardIndex=getIndex('Yard');
    const skuIndex=getIndex('SKU');
    const designIndex=getIndex('Kode Desain');
    const productIndex=getIndex('Nama Produk');
    const colorIndex=getIndex('Warna');
    const widthIndex=getIndex('Lebar');
    const gsmIndex=getIndex('GSM');
    const gradeIndex=getIndex('Grade');
    const statusIndex=getIndex('Status');
    const lastOpnameIndex=getIndex('Last Opname');
    const noteIndex=getIndex('Keterangan');
    const palletIndex=getIndex('Pallet ID');
    const locationIndex=getIndex('Location ID');

    const meter=
      meterIndex!==-1&&row[meterIndex]!==''
        ? Number(row[meterIndex])
        : '';

    const yard=
      yardIndex!==-1&&row[yardIndex]!==''
        ? Number(row[yardIndex])
        : '';

    return {
      success:true,
      data:{
        'Roll ID':String(row[rollIndex]),
        'SKU':skuIndex!==-1?formatCellValue(row[skuIndex]):'',
        'Kode Desain':designIndex!==-1?formatCellValue(row[designIndex]):'',
        'Nama Produk':productIndex!==-1?formatCellValue(row[productIndex]):'',
        'Warna':colorIndex!==-1?formatCellValue(row[colorIndex]):'',
        'Lebar':widthIndex!==-1?formatCellValue(row[widthIndex]):'',
        'GSM':gsmIndex!==-1?formatCellValue(row[gsmIndex]):'',
        'Meter':!isNaN(meter)?meter:'',
        'Yard':!isNaN(yard)?yard:'',
        'Unit':
          meter!==''&&yard===''?'Meter':
          yard!==''&&meter===''?'Yard':
          meter===''&&yard===''?'':'Conflict',
        'Grade':gradeIndex!==-1?formatCellValue(row[gradeIndex]):'',
        'Status':statusIndex!==-1?formatCellValue(row[statusIndex]):'',
        'Last Opname':lastOpnameIndex!==-1?formatCellValue(row[lastOpnameIndex]):'',
        'Keterangan':noteIndex!==-1?formatCellValue(row[noteIndex]):'',
        'Pallet ID':palletIndex!==-1?formatCellValue(row[palletIndex]):'',
        'Location ID':locationIndex!==-1?formatCellValue(row[locationIndex]):''
      }
    };

  }catch(error){
    return {
      success:false,
      message:'Error getRollForOpname: '+error.message
    };
  }
}

function getRollOpnameDatabase(){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName('ROLL');

    if(!sheet){
      return {
        success:false,
        message:'Sheet ROLL tidak ditemukan.'
      };
    }

    const lastRow=sheet.getLastRow();
    const lastCol=sheet.getLastColumn();

    if(lastRow<2){
      return {
        success:true,
        data:[]
      };
    }

    const values=sheet
      .getRange(1,1,lastRow,lastCol)
      .getValues();

    const headers=values[0].map(function(header){
      return String(header).trim();
    });

    function getIndex(name){
      return headers.findIndex(function(header){
        return header.toLowerCase()===name.toLowerCase();
      });
    }

    const rollIndex=getIndex('Roll ID');

    if(rollIndex===-1){
      return {
        success:false,
        message:'Kolom "Roll ID" tidak ditemukan.'
      };
    }

    const meterIndex=getIndex('Meter');
    const yardIndex=getIndex('Yard');
    const skuIndex=getIndex('SKU');
    const designIndex=getIndex('Kode Desain');
    const productIndex=getIndex('Nama Produk');
    const colorIndex=getIndex('Warna');
    const widthIndex=getIndex('Lebar');
    const gsmIndex=getIndex('GSM');
    const gradeIndex=getIndex('Grade');
    const statusIndex=getIndex('Status');
    const lastOpnameIndex=getIndex('Last Opname');
    const noteIndex=getIndex('Keterangan');
    const palletIndex=getIndex('Pallet ID');
    const locationIndex=getIndex('Location ID');

    const data=[];

    for(let i=1;i<values.length;i++){
      const row=values[i];

      const rollId=String(row[rollIndex]||'').trim();

      if(!rollId)continue;

      const meter=
        meterIndex!==-1&&row[meterIndex]!==''
          ? Number(row[meterIndex])
          : '';

      const yard=
        yardIndex!==-1&&row[yardIndex]!==''
          ? Number(row[yardIndex])
          : '';

      data.push({
        'Roll ID':rollId,
        'SKU':skuIndex!==-1?formatCellValue(row[skuIndex]):'',
        'Kode Desain':designIndex!==-1?formatCellValue(row[designIndex]):'',
        'Nama Produk':productIndex!==-1?formatCellValue(row[productIndex]):'',
        'Warna':colorIndex!==-1?formatCellValue(row[colorIndex]):'',
        'Lebar':widthIndex!==-1?formatCellValue(row[widthIndex]):'',
        'GSM':gsmIndex!==-1?formatCellValue(row[gsmIndex]):'',
        'Meter':!isNaN(meter)?meter:'',
        'Yard':!isNaN(yard)?yard:'',
        'Unit':
          meter!==''&&yard===''?'Meter':
          yard!==''&&meter===''?'Yard':
          meter===''&&yard===''?'':'Conflict',
        'Grade':gradeIndex!==-1?formatCellValue(row[gradeIndex]):'',
        'Status':statusIndex!==-1?formatCellValue(row[statusIndex]):'',
        'Last Opname':lastOpnameIndex!==-1?formatCellValue(row[lastOpnameIndex]):'',
        'Keterangan':noteIndex!==-1?formatCellValue(row[noteIndex]):'',
        'Pallet ID':palletIndex!==-1?formatCellValue(row[palletIndex]):'',
        'Location ID':locationIndex!==-1?formatCellValue(row[locationIndex]):''
      });
    }

    return {
      success:true,
      data:data
    };

  }catch(error){
    return {
      success:false,
      message:'Error getRollOpnameDatabase: '+error.message
    };
  }
}

function saveStockOpname(rollId,actualMeter,actualYard,actualGrade,actualStatus,actualLocation,keteranganRoll,token){
  try{
    const session=getLoginSession(token);
    if(!session.success)return {success:false,message:'Session login tidak valid. Silakan login kembali.'};

    const user=session.userId;
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const logSS=getLogSpreadsheet();
    const rollSheet=ss.getSheetByName('ROLL');
    const logSheet=logSS.getSheetByName('OPNAME_LOG');

    if(!rollSheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};
    if(!logSheet)return {success:false,message:'Sheet OPNAME_LOG tidak ditemukan.'};

    const lastRow=rollSheet.getLastRow();
    const lastCol=rollSheet.getLastColumn();

    if(lastRow<2)return {success:false,message:'Sheet ROLL belum memiliki data.'};

    const headers=rollSheet.getRange(1,1,1,lastCol).getValues()[0].map(function(header){
      return String(header).trim();
    });

    function getIndex(name){
      return headers.findIndex(function(header){
        return header.toLowerCase()===name.toLowerCase();
      });
    }

    const rollIndex=getIndex('Roll ID');
    const meterIndex=getIndex('Meter');
    const yardIndex=getIndex('Yard');
    const gradeIndex=getIndex('Grade');
    const palletIndex=getIndex('Pallet ID');
    const locationIndex=getIndex('Location ID');
    const statusIndex=getIndex('Status');
    const lastOpnameIndex=getIndex('Last Opname');
    const noteIndex=getIndex('Keterangan');

    if(rollIndex===-1)return {success:false,message:'Kolom "Roll ID" tidak ditemukan.'};
    if(meterIndex===-1)return {success:false,message:'Kolom "Meter" tidak ditemukan.'};
    if(yardIndex===-1)return {success:false,message:'Kolom "Yard" tidak ditemukan.'};
    if(gradeIndex===-1)return {success:false,message:'Kolom "Grade" tidak ditemukan.'};
    if(statusIndex===-1)return {success:false,message:'Kolom "Status" tidak ditemukan.'};
    if(lastOpnameIndex===-1)return {success:false,message:'Kolom "Last Opname" tidak ditemukan.'};
    if(locationIndex===-1)return {success:false,message:'Kolom "Location ID" tidak ditemukan.'};
    if(noteIndex===-1){return {success:false,message:'Kolom "Keterangan" tidak ditemukan di sheet ROLL.'};
    }

    const searchRange=rollSheet.getRange(2,rollIndex+1,lastRow-1,1);

    const found=searchRange
      .createTextFinder(String(rollId).trim())
      .matchEntireCell(true)
      .matchCase(false)
      .findNext();

    if(!found){
      return {
        success:false,
        message:'Barcode "'+rollId+'" tidak ditemukan.'
      };
    }

    const rowNumber=found.getRow();
    const row=rollSheet.getRange(rowNumber,1,1,lastCol).getValues()[0];

    const databaseMeter=row[meterIndex]!==''?Number(row[meterIndex]):'';
    const databaseYard=row[yardIndex]!==''?Number(row[yardIndex]):'';
    const databaseGrade=String(row[gradeIndex]||'').trim().toUpperCase();
    const databaseStatus=String(row[statusIndex]||'').trim();
    const databaseLocation=String(row[locationIndex]||'').trim();
    const databasePallet=palletIndex!==-1?row[palletIndex]:'';

    const meter=actualMeter!==''&&actualMeter!==null?Number(actualMeter):'';
    const yard=actualYard!==''&&actualYard!==null?Number(actualYard):'';
    const grade=String(actualGrade||'').trim().toUpperCase();
    const status=String(actualStatus||'').trim();
    const location=String(actualLocation||'').trim();

    if(!location)return {success:false,message:'Location aktual harus dipilih.'};

    if(meter!==''&&(isNaN(meter)||meter<0)){
      return {success:false,message:'Meter aktual tidak valid.'};
    }

    if(yard!==''&&(isNaN(yard)||yard<0)){
      return {success:false,message:'Yard aktual tidak valid.'};
    }

    if(meter!==''&&yard!==''){
      return {success:false,message:'Meter dan Yard tidak boleh diisi bersamaan.'};
    }

    if(meter===''&&yard===''){
      return {success:false,message:'Panjang aktual harus diisi.'};
    }

    const validGrades=['A','A1','B','BS'];
    const validStatuses=['Fastmoving','Slowmoving','Returan','Deadstock'];

    if(!validGrades.includes(grade)){
      return {success:false,message:'Grade aktual harus A, A1, B, atau BS.'};
    }

    if(!validStatuses.includes(status)){
      return {success:false,message:'Status aktual tidak valid.'};
    }

    const databaseHasMeter=databaseMeter!=='';
    const databaseHasYard=databaseYard!=='';

    let lengthMatch=false;

    if(databaseHasMeter&&meter!==''){
      lengthMatch=databaseMeter===meter;
    }

    if(databaseHasYard&&yard!==''){
      lengthMatch=databaseYard===yard;
    }

    const unitMatch=
      (databaseHasMeter&&meter!=='')||
      (databaseHasYard&&yard!=='');

    const gradeMatch=databaseGrade===grade;
    const statusMatch=databaseStatus.toLowerCase()===status.toLowerCase();
    const locationMatch=databaseLocation.toLowerCase()===location.toLowerCase();

    let perubahan=[];

    if(!lengthMatch||!unitMatch)perubahan.push('Panjang berbeda');
    if(!gradeMatch)perubahan.push('Grade berbeda');
    if(!statusMatch)perubahan.push('Status berubah');
    if(!locationMatch)perubahan.push('Lokasi berubah');

    let keteranganOpname='Sesuai';

    if(perubahan.length===1){
      keteranganOpname=perubahan[0];
    }else if(perubahan.length>1){
      keteranganOpname=perubahan.slice(0,-1).join(', ')+' dan '+perubahan[perubahan.length-1];
    }

    const opnameDate=new Date();

    // Simpan log OPNAME_LOG
    const logRow=logSheet.getLastRow()+1;

    logSheet.getRange(logRow,1,1,9).setValues([[
      opnameDate,
      rollId,
      meter,
      yard,
      grade,
      status,
      keteranganOpname,
      user,
      String(keteranganRoll||'').trim()
    ]]);

    // Update ROLL sekaligus.
    // J:P = Meter, Yard, Grade, Pallet ID, Location ID, Status, Last Opname
    if(
      meterIndex===9&&
      yardIndex===10&&
      gradeIndex===11&&
      palletIndex===12&&
      locationIndex===13&&
      statusIndex===14&&
      lastOpnameIndex===15
    ){
      rollSheet.getRange(rowNumber,10,1,7).setValues([[
        meter,
        yard,
        grade,
        databasePallet,
        location,
        status,
        opnameDate
      ]]);
    }else{
      // Fallback jika posisi kolom berubah
      rollSheet.getRange(rowNumber,meterIndex+1).setValue(meter);
      rollSheet.getRange(rowNumber,yardIndex+1).setValue(yard);
      rollSheet.getRange(rowNumber,gradeIndex+1).setValue(grade);
      rollSheet.getRange(rowNumber,statusIndex+1).setValue(status);
      rollSheet.getRange(rowNumber,lastOpnameIndex+1).setValue(opnameDate);
      rollSheet.getRange(rowNumber,locationIndex+1).setValue(location);
    }

    // Simpan keterangan permanen pada data roll
    rollSheet.getRange(rowNumber,noteIndex+1)
      .setValue(String(keteranganRoll||'').trim());

    return {
      success:true,
      databaseMeter:databaseMeter,
      actualMeter:meter,
      databaseYard:databaseYard,
      actualYard:yard,
      databaseGrade:databaseGrade,
      actualGrade:grade,
      databaseStatus:databaseStatus,
      actualStatus:status,
      statusChanged:!statusMatch,
      keterangan:keteranganOpname,
      actualKeterangan:String(keteranganRoll||'').trim(),
      lengthMatch:lengthMatch,
      gradeMatch:gradeMatch,
      databaseLocation:databaseLocation,
      actualLocation:location,
      locationMatch:locationMatch
    };

  }catch(error){
    return {
      success:false,
      message:'Error saveStockOpname: '+error.message
    };
  }
}

function formatCellValue(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'dd MMM yyyy, HH:mm'
    );
  }

  return String(value ?? '');
}
