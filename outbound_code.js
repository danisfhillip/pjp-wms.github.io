function getRollForOutbound(rollId){
  try{
    rollId=String(rollId||'').trim();
    if(!rollId)return {success:false,message:'Roll ID kosong.'};

    const cache=CacheService.getScriptCache();
    const cacheKey='outbound_roll_'+rollId.toLowerCase().replace(/[^a-z0-9_-]/g,'_');
    const cached=cache.get(cacheKey);

    if(cached){
      return JSON.parse(cached);
    }

    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName('ROLL');
    if(!sheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};

    const lastRow=sheet.getLastRow();
    const lastCol=sheet.getLastColumn();
    if(lastRow<2)return {success:false,message:'Data ROLL kosong.'};

    const headers=sheet.getRange(1,1,1,lastCol).getValues()[0].map(function(h){return String(h).trim();});
    const getIndex=function(name){return headers.findIndex(function(h){return h.toLowerCase()===name.toLowerCase();});};

    const indexes={
      rollId:getIndex('Roll ID'),
      sku:getIndex('SKU'),
      design:getIndex('Kode Desain'),
      product:getIndex('Nama Produk'),
      color:getIndex('Warna'),
      width:getIndex('Lebar'),
      gsm:getIndex('GSM'),
      meter:getIndex('Meter'),
      yard:getIndex('Yard'),
      grade:getIndex('Grade'),
      location:getIndex('Location ID')
    };

    if(indexes.rollId===-1)return {success:false,message:'Kolom "Roll ID" tidak ditemukan.'};

    const cell=sheet.getRange(2,indexes.rollId+1,lastRow-1,1)
      .createTextFinder(rollId)
      .matchCase(false)
      .matchEntireCell(true)
      .findNext();

    if(!cell)return {success:false,message:'Roll "'+rollId+'" tidak ditemukan.'};

    const row=sheet.getRange(cell.getRow(),1,1,lastCol).getValues()[0];
    const missing=[];

    const required=[
      ['SKU',indexes.sku],
      ['Nama Produk',indexes.product],
      ['Warna',indexes.color],
      ['Lebar',indexes.width],
      ['GSM',indexes.gsm],
      ['Grade',indexes.grade],
      ['Location ID',indexes.location]
    ];

    required.forEach(function(item){
      if(item[1]===-1||String(row[item[1]]||'').trim()==='')missing.push(item[0]);
    });

    const meterValue=indexes.meter!==-1?String(row[indexes.meter]||'').trim():'';
    const yardValue=indexes.yard!==-1?String(row[indexes.yard]||'').trim():'';
    if(!meterValue&&!yardValue)missing.push('Meter/Yard');

    if(missing.length){
      const result={success:false,incomplete:true,message:'Roll '+row[indexes.rollId]+' belum bisa diproses. Data belum lengkap: '+missing.join(', ')+'.'};
      cache.put(cacheKey,JSON.stringify(result),600);
      return result;
    }

    const result={
      success:true,
      data:{
        rollId:String(row[indexes.rollId]||''),
        sku:String(row[indexes.sku]||''),
        design:indexes.design!==-1?String(row[indexes.design]||''):'',
        product:String(row[indexes.product]||''),
        color:String(row[indexes.color]||''),
        width:String(row[indexes.width]||''),
        gsm:String(row[indexes.gsm]||''),
        meter:indexes.meter!==-1?Number(row[indexes.meter]||0):0,
        yard:indexes.yard!==-1?Number(row[indexes.yard]||0):0,
        grade:String(row[indexes.grade]||''),
        location:String(row[indexes.location]||'')
      }
    };

    cache.put(cacheKey,JSON.stringify(result),600);
    return result;
  }catch(error){
    return {success:false,message:'Error getRollForOutbound: '+error.message};
  }
}



function getOutboundRollIndex(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('ROLL');
  if(!sheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};
  const lastRow=sheet.getLastRow();
  const lastCol=sheet.getLastColumn();
  if(lastRow<2)return {success:false,message:'Data ROLL kosong.'};

  const data=sheet.getRange(1,1,lastRow,lastCol).getValues();
  const headers=data[0].map(function(h){return String(h).trim();});
  const getIndex=function(name){return headers.findIndex(function(h){return h.toLowerCase()===name.toLowerCase();});};

  const idx={
    rollId:getIndex('Roll ID'),
    sku:getIndex('SKU'),
    design:getIndex('Kode Desain'),
    product:getIndex('Nama Produk'),
    color:getIndex('Warna'),
    width:getIndex('Lebar'),
    gsm:getIndex('GSM'),
    meter:getIndex('Meter'),
    yard:getIndex('Yard'),
    grade:getIndex('Grade'),
    location:getIndex('Location ID')
  };

  if(idx.rollId===-1)return {success:false,message:'Kolom "Roll ID" tidak ditemukan.'};

  const rolls={};

  for(let i=1;i<data.length;i++){
    const row=data[i];
    const rollId=String(row[idx.rollId]||'').trim();
    if(!rollId)continue;

    const missing=[];
    const required=[
      ['SKU',idx.sku],
      ['Nama Produk',idx.product],
      ['Warna',idx.color],
      ['Lebar',idx.width],
      ['GSM',idx.gsm],
      ['Grade',idx.grade],
      ['Location ID',idx.location]
    ];

    required.forEach(function(item){
      if(item[1]===-1||String(row[item[1]]||'').trim()==='')missing.push(item[0]);
    });

    const meter=idx.meter!==-1?Number(row[idx.meter]||0):0;
    const yard=idx.yard!==-1?Number(row[idx.yard]||0):0;
    if(!meter&&!yard)missing.push('Meter/Yard');

    rolls[rollId.toLowerCase()]={
      valid:missing.length===0,
      missing:missing,
      data:{
        rollId:rollId,
        sku:idx.sku!==-1?String(row[idx.sku]||''):'',
        design:idx.design!==-1?String(row[idx.design]||''):'',
        product:idx.product!==-1?String(row[idx.product]||''):'',
        color:idx.color!==-1?String(row[idx.color]||''):'',
        width:idx.width!==-1?String(row[idx.width]||''):'',
        gsm:idx.gsm!==-1?String(row[idx.gsm]||''):'',
        meter:meter,
        yard:yard,
        grade:idx.grade!==-1?String(row[idx.grade]||''):'',
        location:idx.location!==-1?String(row[idx.location]||''):''
      }
    };
  }

  return {success:true,rolls:rolls,count:Object.keys(rolls).length};
}



function saveOutbound(rollIds,destination,customer,documentNo,note,token){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const session=getLoginSession(token);
    if(!session.success)return {success:false,message:'Session login tidak valid. Silakan login kembali.'};
    const user=session.userId;

    if(!Array.isArray(rollIds)||!rollIds.length)return {success:false,message:'Tidak ada Roll yang dipilih.'};
    destination=String(destination||'').trim();
    customer=String(customer||'').trim();
    documentNo=String(documentNo||'').trim();
    note=String(note||'').trim();
    if(!destination)return {success:false,message:'Tujuan belum dipilih.'};

    const uniqueIds=[];
    const duplicateIds=[];
    rollIds.forEach(function(id){
      const value=String(id||'').trim();
      const key=value.toLowerCase();
      if(!value)return;
      if(uniqueIds.some(function(x){return x.toLowerCase()===key;}))duplicateIds.push(value);
      else uniqueIds.push(value);
    });

    if(duplicateIds.length)return {success:false,message:'Roll duplikat: '+duplicateIds.join(', ')};

    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const rollSheet=ss.getSheetByName('ROLL');
    if(!rollSheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};

    const lastRow=rollSheet.getLastRow();
    const lastCol=rollSheet.getLastColumn();
    if(lastRow<2)return {success:false,message:'Data ROLL kosong.'};

    const data=rollSheet.getRange(1,1,lastRow,lastCol).getValues();
    const headers=data[0].map(function(h){return String(h).trim();});
    const getIndex=function(name){return headers.findIndex(function(h){return h.toLowerCase()===name.toLowerCase();});};

    const idx={
      rollId:getIndex('Roll ID'),
      sku:getIndex('SKU'),
      design:getIndex('Kode Desain'),
      product:getIndex('Nama Produk'),
      color:getIndex('Warna'),
      width:getIndex('Lebar'),
      gsm:getIndex('GSM'),
      meter:getIndex('Meter'),
      yard:getIndex('Yard'),
      grade:getIndex('Grade'),
      status:getIndex('Status'),
      location:getIndex('Location ID')
    };

    if(idx.rollId===-1)return {success:false,message:'Kolom "Roll ID" tidak ditemukan.'};
    if(idx.status===-1)return {success:false,message:'Kolom "Status" tidak ditemukan di ROLL.'};

    const found={};
    for(let i=1;i<data.length;i++){
      const id=String(data[i][idx.rollId]||'').trim();
      if(id)found[id.toLowerCase()]={row:i+1,data:data[i]};
    }

    const missingRolls=[];
    const incomplete=[];
    const logs=[];
    const rowsToDelete=[];

    uniqueIds.forEach(function(requestedId){
      const item=found[requestedId.toLowerCase()];

      if(!item){
        missingRolls.push(requestedId);
        return;
      }

      const row=item.data;
      const missing=[];
      const required=[
        ['SKU',idx.sku],
        ['Nama Produk',idx.product],
        ['Warna',idx.color],
        ['Lebar',idx.width],
        ['GSM',idx.gsm],
        ['Grade',idx.grade],
        ['Status',idx.status],
        ['Location ID',idx.location]
      ];

      required.forEach(function(field){
        if(field[1]===-1||String(row[field[1]]||'').trim()==='')missing.push(field[0]);
      });

      const meter=idx.meter!==-1?Number(row[idx.meter]||0):0;
      const yard=idx.yard!==-1?Number(row[idx.yard]||0):0;
      if(!meter&&!yard)missing.push('Meter/Yard');

      if(missing.length){
        incomplete.push(String(row[idx.rollId]||requestedId)+' ('+missing.join(', ')+')');
        return;
      }

      logs.push([
        null,
        null,
        String(row[idx.rollId]||''),
        idx.sku!==-1?String(row[idx.sku]||''):'',
        idx.design!==-1?String(row[idx.design]||''):'',
        idx.product!==-1?String(row[idx.product]||''):'',
        idx.color!==-1?String(row[idx.color]||''):'',
        idx.width!==-1?row[idx.width]:'',
        idx.gsm!==-1?row[idx.gsm]:'',
        idx.meter!==-1?row[idx.meter]:'',
        idx.yard!==-1?row[idx.yard]:'',
        idx.grade!==-1?String(row[idx.grade]||''):'',
        idx.status!==-1?String(row[idx.status]||''):'',
        destination,
        customer,
        documentNo,
        note,
        user
      ]);

      rowsToDelete.push(item.row);
    });

    if(missingRolls.length)return {success:false,message:'Roll tidak ditemukan: '+missingRolls.join(', ')};
    if(incomplete.length)return {success:false,message:'Data Roll belum lengkap: '+incomplete.join('; ')};

    const logSS=getLogSpreadsheet();
    let logSheet=logSS.getSheetByName('OUTBOUND_LOG');

    const logHeaders=['Tanggal','Outbound ID','Roll ID','SKU','Kode Desain','Nama Produk','Warna','Lebar','GSM','Meter','Yard','Grade','Status','Tujuan','Customer','No. Dokumen','Keterangan','User'];

    if(!logSheet){
      logSheet=logSS.insertSheet('OUTBOUND_LOG');
      logSheet.getRange(1,1,1,logHeaders.length).setValues([logHeaders]);
    }

    const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd');
    const props=PropertiesService.getScriptProperties();
    const seqKey='OUTBOUND_SEQ_'+today;
    let sequence=Number(props.getProperty(seqKey)||0)+1;
    props.setProperty(seqKey,String(sequence));

    const outboundId='OUT-'+today+'-'+String(sequence).padStart(4,'0');
    const now=new Date();

    logs.forEach(function(row){
      row[0]=now;
      row[1]=outboundId;
    });

    const logStartRow=logSheet.getLastRow()+1;
    logSheet.getRange(logStartRow,1,logs.length,logHeaders.length).setValues(logs);
    logSheet.getRange(logStartRow,1,logs.length,1).setNumberFormat('MM/dd/yyyy HH:mm');
    SpreadsheetApp.flush();

    rowsToDelete.sort(function(a,b){return b-a;});
    rowsToDelete.forEach(function(row){
      rollSheet.deleteRow(row);
    });

    SpreadsheetApp.flush();

    return {
      success:true,
      outboundId:outboundId,
      count:logs.length,
      message:'Barang keluar berhasil disimpan.'
    };
  }catch(error){
    return {success:false,message:'Error saveOutbound: '+error.message};
  }finally{
    lock.releaseLock();
  }
}
