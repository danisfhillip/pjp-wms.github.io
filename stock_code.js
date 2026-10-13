function searchStock(search,grade,status,rack){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('ROLL');
  if(!sheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};

  const data=sheet.getDataRange().getValues();
  if(data.length<2)return {success:true,totalRoll:0,totalLength:'0',results:[]};

  const headers=data[0].map(function(h){return String(h).trim();});
  const index=function(name){return headers.findIndex(function(h){return h.toLowerCase()===name.toLowerCase();});};

  const rollIndex=index('Roll ID');
  const skuIndex=index('SKU');
  const designIndex=index('Kode Desain');
  const productIndex=index('Nama Produk');
  const colorIndex=index('Warna');
  const meterIndex=index('Meter');
  const yardIndex=index('Yard');
  const gradeIndex=index('Grade');
  const locationIndex=index('Location ID');
  const statusIndex=index('Status');
  const categoryIndex=index('Kategori');
  const fabricTypeIndex=index('Jenis Kain');
  const widthIndex=index('Lebar');
  const gsmIndex=index('GSM');

  if(rollIndex===-1)return {success:false,message:'Kolom Roll ID tidak ditemukan.'};

  search=String(search||'').trim().toLowerCase();
  grade=Array.isArray(grade)?grade.map(function(x){return String(x).trim().toLowerCase();}):[];
  status=Array.isArray(status)?status.map(function(x){return String(x).trim().toLowerCase();}):[];
  rack=Array.isArray(rack)?rack.map(function(x){return String(x).trim().toLowerCase();}):[];

  const results=[];
  let totalMeter=0;
  let totalYard=0;
  let meterCount=0;
  let yardCount=0;
  let allStockRoll=0;
  let allStockMeter=0;
  const allStockProducts=new Set();
  const allStockCategories=new Set();

  for(let i=1;i<data.length;i++){
    const row=data[i];

    const roll=String(row[rollIndex]||'').trim();
    const sku=skuIndex!==-1?String(row[skuIndex]||'').trim():'';
    const design=designIndex!==-1?String(row[designIndex]||'').trim():'';
    const product=productIndex!==-1?String(row[productIndex]||'').trim():'';
    const color=colorIndex!==-1?String(row[colorIndex]||'').trim():'';
    const rowGrade=gradeIndex!==-1?String(row[gradeIndex]||'').trim().toLowerCase():'';
    const rowStatus=statusIndex!==-1?String(row[statusIndex]||'').trim().toLowerCase():'';
    const location=locationIndex!==-1?String(row[locationIndex]||'').trim():'';
    const meter=Number(row[meterIndex]||0);
    const yard=Number(row[yardIndex]||0);

    if(meter<=0&&yard<=0)continue;

    allStockRoll++;
    if(meter>0)allStockMeter+=meter;

    const rowCategory=categoryIndex!==-1?String(row[categoryIndex]||'').trim():'';
    const rowFabricType=fabricTypeIndex!==-1?String(row[fabricTypeIndex]||'').trim():'';
    const rowWidth=widthIndex!==-1?row[widthIndex]:'';
    const rowGsm=gsmIndex!==-1?row[gsmIndex]:'';

    allStockProducts.add([
      rowCategory,
      rowFabricType,
      product,
      color,
      rowGsm,
      rowWidth
    ].join('|'));

    if(rowCategory)allStockCategories.add(rowCategory);

    const searchText=(sku+' '+design+' '+product+' '+color).toLowerCase();

    if(search&&!searchText.includes(search))continue;
    if(grade.length&&!grade.includes(rowGrade))continue;
    if(status.length&&!status.includes(rowStatus))continue;
    if(rack.length&&!rack.some(function(x){return location.toLowerCase().startsWith(x)}))continue;

    let length='';
    if(meterIndex!==-1&&row[meterIndex]!==''){
      const value=Number(row[meterIndex]);
      if(!isNaN(value)){
        length=value+' Meter';
        totalMeter+=value;
        meterCount++;
      }
    }else if(yardIndex!==-1&&row[yardIndex]!==''){
      const value=Number(row[yardIndex]);
      if(!isNaN(value)){
        length=value+' Yard';
        totalYard+=value;
        yardCount++;
      }
    }

    results.push({
      rollId:roll,
      category:categoryIndex!==-1?String(row[categoryIndex]||'').trim():'',
      fabricType:fabricTypeIndex!==-1?String(row[fabricTypeIndex]||'').trim():'',
      sku:sku,
      design:design,
      product:product,
      color:color,
      width:widthIndex!==-1?row[widthIndex]:'',
      gsm:gsmIndex!==-1?row[gsmIndex]:'',
      meter:meterIndex!==-1&&row[meterIndex]!==''?Number(row[meterIndex]):null,
      yard:yardIndex!==-1&&row[yardIndex]!==''?Number(row[yardIndex]):null,
      length:length,
      grade:rowGrade?rowGrade.toUpperCase():'',
      status:rowStatus,
      location:location
    });
  }

  let totalLength='0';
  if(meterCount&&yardCount){
    totalLength=totalMeter+' Meter + '+totalYard+' Yard';
  }else if(meterCount){
    totalLength=totalMeter+' Meter';
  }else if(yardCount){
    totalLength=totalYard+' Yard';
  }

  return {
    success:true,
    totalRoll:results.length,
    totalLength:totalLength,
    overallTotals:{
      roll:allStockRoll,
      meter:allStockMeter,
      product:allStockProducts.size,
      category:allStockCategories.size
    },
    results:results
  };
}

function checkStockAccess(){
  const email=Session.getActiveUser().getEmail().trim().toLowerCase();
  if(!email)return {allowed:false,email:''};

  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('USER_ACCESS');
  if(!sheet)return {allowed:false,email:email};

  const data=sheet.getDataRange().getValues();
  if(data.length<2)return {allowed:false,email:email};

  for(let i=1;i<data.length;i++){
    const rowEmail=String(data[i][0]||'').trim().toLowerCase();
    const stock=String(data[i][1]||'').trim().toLowerCase();

    if(rowEmail===email&&['true','yes','1'].includes(stock)){
      return {allowed:true,email:email};
    }
  }

  return {allowed:false,email:email};
}
