function getRollForMovement(rollId){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('ROLL');
  if(!sheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};

  const data=sheet.getDataRange().getValues();
  if(data.length<2)return {success:false,message:'Data ROLL kosong.'};

  const headers=data[0].map(function(h){return String(h).trim();});
  const index=function(name){return headers.indexOf(name);};
  const rollIndex=index('Roll ID');

  for(let i=1;i<data.length;i++){
    if(String(data[i][rollIndex]).trim()===String(rollId).trim()){
      return {
        success:true,
        row:i+1,
        data:{
          'Roll ID':data[i][index('Roll ID')],
          'Nama Produk':data[i][index('Nama Produk')],
          'Warna':data[i][index('Warna')],
          'Meter':data[i][index('Meter')],
          'Yard':data[i][index('Yard')],
          'Grade':data[i][index('Grade')],
          'Location ID':data[i][index('Location ID')],
          'Pallet ID':data[i][index('Pallet ID')]
        }
      };
    }
  }

  return {success:false,message:'Roll ID tidak ditemukan.'};
}

function getMovementRollData(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('ROLL');
  if(!sheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};

  const data=sheet.getDataRange().getValues();
  if(data.length<2)return {success:false,message:'Data ROLL kosong.'};

  const headers=data[0].map(function(h){return String(h).trim();});
  const index=function(name){return headers.indexOf(name);};
  const rollIndex=index('Roll ID');

  if(rollIndex===-1)return {success:false,message:'Kolom Roll ID tidak ditemukan.'};

  const rolls={};

  for(let i=1;i<data.length;i++){
    const rollId=String(data[i][rollIndex]||'').trim();
    if(!rollId)continue;

    rolls[rollId.toLowerCase()]={
      'Roll ID':data[i][index('Roll ID')],
      'Nama Produk':data[i][index('Nama Produk')],
      'Warna':data[i][index('Warna')],
      'Meter':data[i][index('Meter')],
      'Yard':data[i][index('Yard')],
      'Grade':data[i][index('Grade')],
      'Location ID':data[i][index('Location ID')],
      'Pallet ID':data[i][index('Pallet ID')]
    };
  }

  return {success:true,rolls:rolls};
}

function saveMovement(rollIds,targetLocation,keterangan,token){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const session=getLoginSession(token);
    if(!session.success)return {success:false,message:'Session login tidak valid. Silakan login kembali.'};

    const user=session.userId;
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const logSS=getLogSpreadsheet();
    const rollSheet=ss.getSheetByName('ROLL');
    const logSheet=logSS.getSheetByName('MOVEMENT_LOG');

    if(!rollSheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};
    if(!logSheet)return {success:false,message:'Sheet MOVEMENT_LOG tidak ditemukan.'};
    if(!Array.isArray(rollIds)||!rollIds.length)return {success:false,message:'Tidak ada Roll yang dipilih.'};
    if(!targetLocation)return {success:false,message:'Location tujuan belum dipilih.'};

    const data=rollSheet.getDataRange().getValues();
    if(data.length<2)return {success:false,message:'Data ROLL kosong.'};

    const headers=data[0].map(function(h){return String(h).trim();});
    const rollIndex=headers.findIndex(function(h){return h.toLowerCase()==='roll id';});
    const locationIndex=headers.findIndex(function(h){return h.toLowerCase()==='location id';});

    if(rollIndex===-1||locationIndex===-1)return {success:false,message:'Kolom Roll ID atau Location ID tidak ditemukan.'};

    const requestedIds=rollIds.map(function(id){return String(id).trim().toLowerCase();});
    const foundIds=[];
    const logs=[];
    const now=new Date();

    for(let r=1;r<data.length;r++){
      const currentId=String(data[r][rollIndex]||'').trim();
      const currentIdLower=currentId.toLowerCase();
      const requestedIndex=requestedIds.indexOf(currentIdLower);

      if(requestedIndex===-1)continue;

      const oldLocation=String(data[r][locationIndex]||'').trim();

      if(oldLocation===String(targetLocation).trim()){
        foundIds.push(currentIdLower);
        continue;
      }

      data[r][locationIndex]=targetLocation;
      foundIds.push(currentIdLower);
      logs.push([now,currentId,oldLocation,targetLocation,keterangan||'',user]);
    }

    const notFound=requestedIds.filter(function(id){return foundIds.indexOf(id)===-1;});

    if(notFound.length){
      return {success:false,message:'Roll tidak ditemukan: '+notFound.join(', ')};
    }

    if(logs.length){
      rollSheet.getRange(1,1,data.length,headers.length).setValues(data);
      logSheet.getRange(logSheet.getLastRow()+1,1,logs.length,6).setValues(logs);
      logSheet.getRange(logSheet.getLastRow()-logs.length+1,1,logs.length,1).setNumberFormat('MM/dd/yyyy HH:mm');
      SpreadsheetApp.flush();
    }

    return {success:true,count:logs.length};
  }finally{
    lock.releaseLock();
  }
}
