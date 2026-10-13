function getRollForSampleCut(rollId){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('ROLL');
  if(!sheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};

  const data=sheet.getDataRange().getValues();
  if(data.length<2)return {success:false,message:'Data ROLL kosong.'};

  const headers=data[0].map(function(h){return String(h).trim();});
  const index=function(name){return headers.findIndex(function(h){return h.toLowerCase()===name.toLowerCase();});};

  const rollIndex=index('Roll ID');
  const productIndex=index('Nama Produk');
  const colorIndex=index('Warna');
  const meterIndex=index('Meter');
  const yardIndex=index('Yard');
  const gradeIndex=index('Grade');
  const locationIndex=index('Location ID');

  if(rollIndex===-1)return {success:false,message:'Kolom Roll ID tidak ditemukan.'};

  rollId=String(rollId||'').trim().toLowerCase();

  for(let i=1;i<data.length;i++){
    if(String(data[i][rollIndex]||'').trim().toLowerCase()===rollId){
      const meter=meterIndex!==-1&&data[i][meterIndex]!==''?Number(data[i][meterIndex]):'';
      const yard=yardIndex!==-1&&data[i][yardIndex]!==''?Number(data[i][yardIndex]):'';

      return {
        success:true,
        data:{
          'Roll ID':String(data[i][rollIndex]),
          'Nama Produk':productIndex!==-1?String(data[i][productIndex]||''):'',
          'Warna':colorIndex!==-1?String(data[i][colorIndex]||''):'',
          'Meter':!isNaN(meter)?meter:'',
          'Yard':!isNaN(yard)?yard:'',
          'Grade':gradeIndex!==-1?String(data[i][gradeIndex]||''):'',
          'Location ID':locationIndex!==-1?String(data[i][locationIndex]||''):''
        }
      };
    }
  }

  return {success:false,message:'Roll ID tidak ditemukan.'};
}

function saveSampleCut(rollId,jumlahPotong,keterangan,token){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const session=getLoginSession(token);
    if(!session.success)return {success:false,message:'Session login tidak valid. Silakan login kembali.'};
    const user=session.userId;

    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const logSS=getLogSpreadsheet();
    const rollSheet=ss.getSheetByName('ROLL');
    const logSheet=logSS.getSheetByName('SAMPLE_CUT_LOG');

    if(!rollSheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};
    if(!logSheet)return {success:false,message:'Sheet SAMPLE_CUT_LOG tidak ditemukan.'};

    const data=rollSheet.getDataRange().getValues();
    if(data.length<2)return {success:false,message:'Data ROLL kosong.'};

    const headers=data[0].map(function(h){return String(h).trim();});
    const index=function(name){return headers.findIndex(function(h){return h.toLowerCase()===name.toLowerCase();});};

    const rollIndex=index('Roll ID');
    const meterIndex=index('Meter');
    const yardIndex=index('Yard');

    if(rollIndex===-1)return {success:false,message:'Kolom Roll ID tidak ditemukan.'};
    if(meterIndex===-1)return {success:false,message:'Kolom Meter tidak ditemukan.'};
    if(yardIndex===-1)return {success:false,message:'Kolom Yard tidak ditemukan.'};

    const requestedId=String(rollId||'').trim().toLowerCase();
    const potong=Number(jumlahPotong);

    if(!potong||potong<=0)return {success:false,message:'Jumlah potong tidak valid.'};

    for(let i=1;i<data.length;i++){
      const currentId=String(data[i][rollIndex]||'').trim();

      if(currentId.toLowerCase()!==requestedId)continue;

      const meter=data[i][meterIndex];
      const yard=data[i][yardIndex];

      let unit='';
      let sebelum='';
      let sesudah='';

      if(meter!==''&&meter!==null){
        sebelum=Number(meter);
        unit='Meter';
        sesudah=sebelum-potong;
      }else if(yard!==''&&yard!==null){
        sebelum=Number(yard);
        unit='Yard';
        sesudah=sebelum-potong;
      }else{
        return {success:false,message:'Roll belum memiliki panjang Meter atau Yard.'};
      }

      if(isNaN(sebelum)||isNaN(sesudah))return {success:false,message:'Panjang roll tidak valid.'};
      if(potong>sebelum)return {success:false,message:'Jumlah potong melebihi panjang roll.'};
      if(sesudah<0)return {success:false,message:'Sisa panjang roll tidak boleh kurang dari 0.'};

      if(unit==='Meter'){
        data[i][meterIndex]=sesudah;
      }else{
        data[i][yardIndex]=sesudah;
      }

      const now=new Date();

      logSheet.appendRow([
        now,
        currentId,
        sebelum,
        potong,
        sesudah,
        unit,
        keterangan||'',
        user
      ]);

      rollSheet.getRange(i+1,1,1,headers.length).setValues([data[i]]);
      logSheet.getRange(logSheet.getLastRow(),1).setNumberFormat('MM/dd/yyyy HH:mm');
      SpreadsheetApp.flush();

      return {
        success:true,
        rollId:currentId,
        unit:unit,
        sebelum:sebelum,
        potong:potong,
        sesudah:sesudah
      };
    }

    return {success:false,message:'Roll ID tidak ditemukan.'};
  }finally{
    lock.releaseLock();
  }
}
