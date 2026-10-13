function getRackMaster(){
  let sheet;

  try{
    sheet=ensureLocationStatusColumn_();
  }catch(error){
    return {success:false,message:error.message};
  }

  const data=sheet.getDataRange().getValues();
  if(data.length<2)return {success:true,racks:[]};

  const rackMap={};

  for(let i=1;i<data.length;i++){
    const type=String(data[i][0]||'').trim();
    const rack=String(data[i][1]||'').trim();
    const position=String(data[i][2]||'').trim();
    const level=String(data[i][3]||'').trim();
    const locationId=String(data[i][4]||'').trim();
    const status=String(data[i][5]||'Kosong').trim();

    if(type!=='Rak'||!rack||!position||!level||!locationId)continue;

    if(!rackMap[rack]){
      rackMap[rack]={rack:rack,positions:[],levels:[],locations:[]};
    }

    const item=rackMap[rack];

    if(!item.positions.includes(position))item.positions.push(position);
    if(!item.levels.includes(level))item.levels.push(level);

    item.locations.push({
      id:locationId,
      position:position,
      level:level,
      status:status||'Kosong'
    });
  }

  const racks=Object.values(rackMap);

  racks.forEach(function(item){
    item.positions.sort(function(a,b){return Number(a)-Number(b);});
    item.levels.sort();
    item.locations.sort(function(a,b){
      return Number(a.position)-Number(b.position)||a.level.localeCompare(b.level);
    });
    item.startPosition=item.positions.length?item.positions[0]:'';
    item.endPosition=item.positions.length?item.positions[item.positions.length-1]:'';
    item.totalLocation=item.locations.length;
  });

  racks.sort(function(a,b){
    return a.rack.localeCompare(b.rack,undefined,{numeric:true});
  });

  return {success:true,racks:racks};
}



function updateRackConfiguration(rack,endPosition,levels,token){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);

  try{
    const session=getLoginSession(token);
    if(!session.success){
      return {success:false,message:'Session login tidak valid. Silakan login kembali.'};
    }

    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const locationSheet=ensureLocationStatusColumn_();
    const rollSheet=ss.getSheetByName('ROLL');

    if(!rollSheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};

    rack=String(rack||'').trim();
    endPosition=Number(endPosition);
    levels=Array.isArray(levels)?levels.map(function(item){
      return String(item).trim().toUpperCase();
    }):[];
    levels=[...new Set(levels)];

    if(!rack)return {success:false,message:'Rack harus diisi.'};
    if(!Number.isInteger(endPosition)||endPosition<1){
      return {success:false,message:'Jumlah posisi tidak valid.'};
    }
    if(!levels.length)return {success:false,message:'Minimal pilih satu level.'};

    const allowedLevels=['A','B','C','D','E'];
    if(levels.some(function(level){return !allowedLevels.includes(level);})){
      return {success:false,message:'Level tidak valid.'};
    }

    const locationData=locationSheet.getDataRange().getValues();
    const existing={};
    const existingRows={};

    for(let i=1;i<locationData.length;i++){
      const type=String(locationData[i][0]||'').trim();
      const currentRack=String(locationData[i][1]||'').trim();
      const position=String(locationData[i][2]||'').trim();
      const level=String(locationData[i][3]||'').trim();
      const locationId=String(locationData[i][4]||'').trim();

      if(type==='Rak'&&currentRack===rack&&locationId){
        existing[locationId]={row:i+1,position:position,level:level};
        existingRows[locationId]=i+1;
      }
    }

    if(!Object.keys(existing).length){
      return {success:false,message:'Rack tidak ditemukan.'};
    }

    const rollHeaders=rollSheet.getRange(1,1,1,rollSheet.getLastColumn())
      .getValues()[0].map(function(header){return String(header).trim();});

    const locationIndex=rollHeaders.indexOf('Location ID');
    if(locationIndex===-1){
      return {success:false,message:'Kolom "Location ID" tidak ditemukan di ROLL.'};
    }

    const rollData=rollSheet.getDataRange().getValues();
    const usedLocations=new Set();

    for(let i=1;i<rollData.length;i++){
      const locationId=String(rollData[i][locationIndex]||'').trim();
      if(locationId)usedLocations.add(locationId);
    }

    const targetLocations=new Set();

    for(let pos=1;pos<=endPosition;pos++){
      const position=String(pos).padStart(2,'0');
      levels.forEach(function(level){
        targetLocations.add(rack+' '+position+'-'+level);
      });
    }

    const locationsToDelete=Object.keys(existing).filter(function(locationId){
      return !targetLocations.has(locationId);
    });

    const blockedLocations=locationsToDelete.filter(function(locationId){
      return usedLocations.has(locationId);
    });

    if(blockedLocations.length){
      return {
        success:false,
        message:'Perubahan tidak dapat dilakukan karena ada location yang masih digunakan oleh Roll.',
        blockedLocations:blockedLocations
      };
    }

    const locationsToAdd=[];

    targetLocations.forEach(function(locationId){
      if(!existing[locationId]){
        const match=locationId.match(/^(.+) (\d{2})-([A-Z])$/);

        if(match){
          locationsToAdd.push([
            'Rak',
            rack,
            match[2],
            match[3],
            locationId,
            'Kosong'
          ]);
        }
      }
    });

    if(locationsToAdd.length){
      locationSheet.getRange(
        locationSheet.getLastRow()+1,1,locationsToAdd.length,6
      ).setValues(locationsToAdd);
    }

    if(locationsToDelete.length){
      locationsToDelete.map(function(id){return existingRows[id];})
        .sort(function(a,b){return b-a;})
        .forEach(function(row){locationSheet.deleteRow(row);});
    }

    SpreadsheetApp.flush();

    return {
      success:true,
      rack:rack,
      added:locationsToAdd.map(function(row){return row[4];}),
      deleted:locationsToDelete,
      totalLocation:targetLocations.size
    };
  }finally{
    lock.releaseLock();
  }
}



function generateLocations(rack,startPosition,endPosition,levels,token){
  const session=getLoginSession(token);
  if(!session.success){
    return {success:false,message:'Session login tidak valid. Silakan login kembali.'};
  }

  let sheet;
  try{
    sheet=ensureLocationStatusColumn_();
  }catch(error){
    return {success:false,message:error.message};
  }

  rack=String(rack||'').trim();
  startPosition=Number(startPosition);
  endPosition=Number(endPosition);
  levels=Array.isArray(levels)?levels:[];

  if(!rack)return {success:false,message:'Rack harus diisi.'};
  if(!Number.isInteger(startPosition)||!Number.isInteger(endPosition)||startPosition<1||endPosition<startPosition){
    return {success:false,message:'Range posisi tidak valid.'};
  }
  if(!levels.length)return {success:false,message:'Minimal pilih satu level.'};

  levels=[...new Set(levels.map(function(level){
    return String(level).trim().toUpperCase();
  }))];

  if(levels.some(function(level){return !['A','B','C','D','E'].includes(level);})){
    return {success:false,message:'Level tidak valid.'};
  }

  const data=sheet.getDataRange().getValues();
  const existing=new Set();

  for(let i=1;i<data.length;i++){
    const locationId=String(data[i][4]||'').trim();
    if(locationId)existing.add(locationId);
  }

  const rows=[];
  const created=[];
  const skipped=[];

  for(let pos=startPosition;pos<=endPosition;pos++){
    const position=String(pos).padStart(2,'0');

    levels.forEach(function(level){
      const locationId=rack+' '+position+'-'+level;

      if(existing.has(locationId)){
        skipped.push(locationId);
        return;
      }

      rows.push(['Rak',rack,position,level,locationId,'Kosong']);
      created.push(locationId);
      existing.add(locationId);
    });
  }

  if(rows.length){
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,6).setValues(rows);
  }

  SpreadsheetApp.flush();

  return {success:true,created:created,skipped:skipped};
}



function updateLocationStatus(locationId,status,token){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);

  try{
    const session=getLoginSession(token);
    if(!session.success){
      return {success:false,message:'Session login tidak valid. Silakan login kembali.'};
    }

    locationId=String(locationId||'').trim();
    status=String(status||'').trim();

    const allowedStatuses=['Kosong','Tersedia','Hampir Penuh','Penuh'];

    if(!locationId)return {success:false,message:'Location ID harus diisi.'};
    if(!allowedStatuses.includes(status)){
      return {success:false,message:'Status lokasi tidak valid.'};
    }

    const sheet=ensureLocationStatusColumn_();
    const lastRow=sheet.getLastRow();

    if(lastRow<2)return {success:false,message:'Location belum tersedia.'};

    const data=sheet.getRange(2,1,lastRow-1,6).getValues();
    let targetRow=0;

    for(let i=0;i<data.length;i++){
      if(String(data[i][4]||'').trim()===locationId){
        targetRow=i+2;
        break;
      }
    }

    if(!targetRow)return {success:false,message:'Location tidak ditemukan.'};

    sheet.getRange(targetRow,6).setValue(status);
    SpreadsheetApp.flush();

    return {
      success:true,
      locationId:locationId,
      status:status
    };
  }finally{
    lock.releaseLock();
  }
}



function getLocationStockDetails(locationId,token){
  const session=getLoginSession(token);
  if(!session.success){
    return {success:false,message:'Session login tidak valid. Silakan login kembali.'};
  }

  locationId=String(locationId||'').trim();
  if(!locationId)return {success:false,message:'Location ID harus diisi.'};

  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const rollSheet=ss.getSheetByName('ROLL');

  if(!rollSheet)return {success:false,message:'Sheet ROLL tidak ditemukan.'};

  const data=rollSheet.getDataRange().getValues();
  if(!data.length)return {success:false,message:'Sheet ROLL belum memiliki header.'};

  const headers=data[0].map(function(header){return String(header).trim();});
  const index={};

  headers.forEach(function(header,i){index[header]=i;});

  if(index['Location ID']===undefined){
    return {success:false,message:'Kolom "Location ID" tidak ditemukan di ROLL.'};
  }

  const pick=function(row,header){
    return index[header]===undefined?'':row[index[header]];
  };

  const rolls=[];
  let totalMeter=0;
  let totalYard=0;

  for(let i=1;i<data.length;i++){
    const row=data[i];

    if(String(pick(row,'Location ID')||'').trim()!==locationId)continue;

    const meter=Number(pick(row,'Meter'))||0;
    const yard=Number(pick(row,'Yard'))||0;

    totalMeter+=meter;
    totalYard+=yard;

    rolls.push({
      rollId:pick(row,'Roll ID'),
      sku:pick(row,'SKU'),
      category:pick(row,'Kategori'),
      fabricType:pick(row,'Jenis Kain'),
      design:pick(row,'Kode Desain'),
      product:pick(row,'Nama Produk'),
      color:pick(row,'Warna'),
      grade:pick(row,'Grade'),
      status:pick(row,'Status'),
      meter:meter,
      yard:yard
    });
  }

  return {
    success:true,
    locationId:locationId,
    rollCount:rolls.length,
    totalMeter:totalMeter,
    totalYard:totalYard,
    rolls:rolls
  };
}



function getRackStockDetails(rackName,token){
  const session=getLoginSession(token);
  if(!session.success){
    return {success:false,message:'Session login tidak valid. Silakan login kembali.'};
  }

  rackName=String(rackName||'').trim();
  if(!rackName)return {success:false,message:'Nama rack harus diisi.'};

  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const locationSheet=ss.getSheetByName('LOCATION_MASTER');
  const rollSheet=ss.getSheetByName('ROLL');

  if(!locationSheet||!rollSheet){
    return {success:false,message:'Sheet LOCATION_MASTER atau ROLL tidak ditemukan.'};
  }

  const locationData=locationSheet.getDataRange().getValues();
  const rollData=rollSheet.getDataRange().getValues();

  if(!rollData.length)return {success:false,message:'Header ROLL tidak ditemukan.'};

  const headers=rollData[0].map(function(h){return String(h).trim();});
  const idx={};
  headers.forEach(function(h,i){idx[h]=i;});

  if(idx['Location ID']===undefined){
    return {success:false,message:'Kolom Location ID tidak ditemukan.'};
  }

  const locations={};

  for(let i=1;i<locationData.length;i++){
    const row=locationData[i];
    if(String(row[0]||'').trim()!=='Rak')continue;
    if(String(row[1]||'').trim()!==rackName)continue;

    const id=String(row[4]||'').trim();
    if(!id)continue;

    locations[id]={
      locationId:id,
      status:String(row[5]||'Kosong').trim()||'Kosong',
      rollCount:0,
      totalMeter:0,
      totalYard:0,
      rolls:[]
    };
  }

  const get=function(row,key){
    return idx[key]===undefined?'':row[idx[key]];
  };

  for(let i=1;i<rollData.length;i++){
    const row=rollData[i];
    const id=String(get(row,'Location ID')||'').trim();
    const location=locations[id];

    if(!location)continue;

    const meter=Number(get(row,'Meter'))||0;
    const yard=Number(get(row,'Yard'))||0;

    location.rollCount++;
    location.totalMeter+=meter;
    location.totalYard+=yard;

    location.rolls.push({
      rollId:get(row,'Roll ID'),
      sku:get(row,'SKU'),
      category:get(row,'Kategori'),
      fabricType:get(row,'Jenis Kain'),
      design:get(row,'Kode Desain'),
      product:get(row,'Nama Produk'),
      color:get(row,'Warna'),
      gsm:get(row,'GSM'),
      width:get(row,'Lebar'),
      grade:get(row,'Grade'),
      status:get(row,'Status'),
      meter:meter,
      yard:yard
    });
  }

  return {
    success:true,
    rack:rackName,
    locations:locations
  };
}



function ensureLocationStatusColumn_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('LOCATION_MASTER');

  if(!sheet)throw new Error('Sheet LOCATION_MASTER tidak ditemukan.');

  const header=String(sheet.getRange(1,6).getValue()||'').trim();

  if(header&&header!=='Status Lokasi'){
    throw new Error('Kolom F sudah digunakan oleh "'+header+'". Periksa struktur LOCATION_MASTER terlebih dahulu.');
  }

  if(!header)sheet.getRange(1,6).setValue('Status Lokasi');

  const lastRow=sheet.getLastRow();
  if(lastRow<2)return sheet;

  const statuses=sheet.getRange(2,6,lastRow-1,1).getValues();
  let changed=false;

  for(let i=0;i<statuses.length;i++){
    if(!String(statuses[i][0]||'').trim()){
      statuses[i][0]='Kosong';
      changed=true;
    }
  }

  if(changed)sheet.getRange(2,6,statuses.length,1).setValues(statuses);

  return sheet;
}
