/********* DASHBOARD DATA **********/
function getDashboardData(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('ROLL');

  if(!sheet){
    return{
      success:false,
      message:'Sheet ROLL tidak ditemukan.'
    };
  }

  const lastRow=sheet.getLastRow();
  const lastCol=sheet.getLastColumn();

  let totalRoll=0;
  let stockAvailable=0;
  let notOpname=0;

  if(lastRow>=2){

    // Baca seluruh kolom ROLL secara dinamis
    const data=sheet.getRange(
      1,
      1,
      lastRow,
      lastCol
    ).getValues();

    const headers=data[0].map(function(h){
      return String(h).trim().toLowerCase();
    });

    const index=function(name){
      return headers.indexOf(name.toLowerCase());
    };

    const rollIndex=index('roll id');
    const meterIndex=index('meter');
    const yardIndex=index('yard');
    const lastOpnameIndex=index('last opname');

    if(rollIndex===-1){
      return{
        success:false,
        message:'Kolom Roll ID tidak ditemukan.'
      };
    }

    if(lastOpnameIndex===-1){
      return{
        success:false,
        message:'Kolom Last Opname tidak ditemukan.'
      };
    }

    for(let i=1;i<data.length;i++){

      const rollId=String(
        data[i][rollIndex]||''
      ).trim();

      if(!rollId)continue;

      totalRoll++;

      const lastOpname=data[i][lastOpnameIndex];

      // Belum pernah opname
      if(lastOpname===''||lastOpname===null){
        notOpname++;
        continue;
      }

      const meter=
        meterIndex!==-1
          ?Number(data[i][meterIndex]||0)
          :0;

      const yard=
        yardIndex!==-1
          ?Number(data[i][yardIndex]||0)
          :0;

      // Sudah opname dan masih memiliki panjang
      if(meter>0||yard>0){
        stockAvailable++;
      }
    }
  }

  const outbound=getTodayLogCount_('OUTBOUND_LOG');

  return{
    success:true,
    totalRoll:totalRoll,
    stockAvailable:stockAvailable,
    notOpname:notOpname,
    outbound:outbound
  };
}



/********* DASHBOARD ACTIVITY **********/
function getDashboardActivity(){

  try{

    const result={
      receiving:0,
      opname:getTodayLogCount_('OPNAME_LOG'),
      movement:getTodayLogCount_('MOVEMENT_LOG'),
      sampleCut:getTodayLogCount_('SAMPLE_CUT_LOG'),
      outbound:getTodayLogCount_('OUTBOUND_LOG')
    };

    return{
      success:true,
      data:result
    };

  }catch(error){

    return{
      success:false,
      message:'Error getDashboardActivity: '+error.message
    };

  }
}



/********* HITUNG AKTIVITAS HARI INI **********/
function getTodayLogCount_(sheetName){
  const logSS=getLogSpreadsheet();
  const sheet=logSS.getSheetByName(sheetName);

  if(!sheet||sheet.getLastRow()<2)return 0;

  const data=sheet.getDataRange().getValues();
  if(data.length<2)return 0;

  const headers=data[0].map(function(h){
    return String(h).trim().toLowerCase();
  });

  let dateIndex=headers.indexOf('tanggal');

  // Kalau header Tanggal tidak ditemukan,
  // gunakan kolom pertama sebagai fallback
  if(dateIndex===-1){
    dateIndex=0;
  }

  const timezone=Session.getScriptTimeZone();
  const todayText=Utilities.formatDate(new Date(),timezone,'yyyy-MM-dd');

  let count=0;

  for(let i=1;i<data.length;i++){
    const value=data[i][dateIndex];
    if(!value)continue;

    let dateText='';

    if(value instanceof Date){
      dateText=Utilities.formatDate(
        value,
        timezone,
        'yyyy-MM-dd'
      );
    }else{
      const parsed=new Date(value);

      if(!isNaN(parsed.getTime())){
        dateText=Utilities.formatDate(
          parsed,
          timezone,
          'yyyy-MM-dd'
        );
      }
    }

    if(dateText===todayText){
      count++;
    }
  }

  return count;
}



/********* DASHBOARD DETAIL **********/
function getDashboardRollDetail(type){
  if(type==='outbound'){
    const logSS=getLogSpreadsheet();
    const sheet=logSS.getSheetByName('OUTBOUND_LOG');
    if(!sheet||sheet.getLastRow()<2)return {success:true,data:[]};

    const data=sheet.getDataRange().getValues();
    const headers=data[0].map(function(h){return String(h).trim();});
    const index=function(name){return headers.findIndex(function(h){return h.toLowerCase()===name.toLowerCase();});};

    const dateIndex=index('Tanggal');
    const rollIndex=index('Roll ID');
    const productIndex=index('Nama Produk');
    const colorIndex=index('Warna');
    const meterIndex=index('Meter');
    const yardIndex=index('Yard');
    const gradeIndex=index('Grade');
    const destinationIndex=index('Tujuan');
    const customerIndex=index('Customer');
    const documentIndex=index('No Dokumen');
    const outboundIdIndex=index('Outbound ID');

    if(dateIndex===-1)return {success:false,message:'Kolom Tanggal tidak ditemukan di OUTBOUND_LOG.'};

    const timezone=Session.getScriptTimeZone();
    const todayText=Utilities.formatDate(new Date(),timezone,'yyyy-MM-dd');
    const result=[];

    for(let i=1;i<data.length;i++){
      const value=data[i][dateIndex];
      if(!value)continue;

      let dateText='';
      if(value instanceof Date){
        dateText=Utilities.formatDate(value,timezone,'yyyy-MM-dd');
      }else{
        const parsed=new Date(value);
        if(!isNaN(parsed.getTime()))dateText=Utilities.formatDate(parsed,timezone,'yyyy-MM-dd');
      }

      if(dateText!==todayText)continue;

      result.push({
        rollId:rollIndex!==-1?String(data[i][rollIndex]||'-'):'-',
        product:productIndex!==-1?String(data[i][productIndex]||'-'):'-',
        color:colorIndex!==-1?String(data[i][colorIndex]||'-'):'-',
        meter:meterIndex!==-1?data[i][meterIndex]:'',
        yard:yardIndex!==-1?data[i][yardIndex]:'',
        grade:gradeIndex!==-1?String(data[i][gradeIndex]||'-'):'-',
        destination:destinationIndex!==-1?String(data[i][destinationIndex]||'-'):'-',
        customer:customerIndex!==-1?String(data[i][customerIndex]||'-'):'-',
        documentNo:documentIndex!==-1?String(data[i][documentIndex]||'-'):'-',
        outboundId:outboundIdIndex!==-1?String(data[i][outboundIdIndex]||'-'):'-'
      });
    }

    return {success:true,data:result};
  }

  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('ROLL');
  if(!sheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};

  const data=sheet.getDataRange().getValues();
  if(data.length<2)return {success:true,data:[]};

  const headers=data[0].map(function(h){return String(h).trim();});
  const index=function(name){return headers.findIndex(function(h){return h.toLowerCase()===name.toLowerCase();});};

  const rollIndex=index('Roll ID');
  const productIndex=index('Nama Produk');
  const colorIndex=index('Warna');
  const meterIndex=index('Meter');
  const yardIndex=index('Yard');
  const lastOpnameIndex=index('Last Opname');

  if(rollIndex===-1)return {success:false,message:'Kolom Roll ID tidak ditemukan.'};

  const result=[];

  for(let i=1;i<data.length;i++){
    const rollId=String(data[i][rollIndex]||'').trim();
    if(!rollId)continue;

    const lastOpname=lastOpnameIndex!==-1?data[i][lastOpnameIndex]:'';
    const meter=meterIndex!==-1?Number(data[i][meterIndex]||0):0;
    const yard=yardIndex!==-1?Number(data[i][yardIndex]||0):0;

    let match=false;

    if(type==='notOpname'){
      match=lastOpname==='';
    }else if(type==='available'){
      match=lastOpname!==''&&(meter>0||yard>0);
    }

    if(!match)continue;

    result.push({
      rollId:rollId,
      product:productIndex!==-1?String(data[i][productIndex]||'-'):'-',
      color:colorIndex!==-1?String(data[i][colorIndex]||'-'):'-',
      location:String(data[i][index('Location ID')]||data[i][index('Pallet ID')]||'-'),
      meter:meterIndex!==-1?data[i][meterIndex]:'',
      yard:yardIndex!==-1?data[i][yardIndex]:'',
      grade:String(data[i][index('Grade')]||''),
      status:String(data[i][index('Status')]||'')
    });
  }

  return {success:true,data:result};
}
