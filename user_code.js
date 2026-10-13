function getUsers(token){
  if(!hasPermission(token,'Admin'))return {success:false,message:'Akses ditolak.'};
  const sheet=getUserAccessSheet_();
  const data=sheet.getDataRange().getValues();
  if(data.length<2)return {success:true,users:[]};
  const headers=data[0].map(function(h){return String(h).trim();});
  const users=[];
  for(let i=1;i<data.length;i++){
    const userId=String(data[i][0]||'').trim();
    if(!userId)continue;
    const user={userId:userId,hasPin:!!data[i][1],status:String(data[i][12]||'Aktif').trim(),permissions:{}};
    for(let j=2;j<headers.length;j++){
      if(headers[j])user.permissions[headers[j]]=data[i][j]===true;
    }
    users.push(user);
  }
  return {success:true,users:users};
}

function addUser(token,userId,pin,permissions){
  if(!hasPermission(token,'Admin'))return {success:false,message:'Akses ditolak.'};

  const sheet=getUserAccessSheet_();
  const data=sheet.getDataRange().getValues();

  const target=String(userId||'').trim();

  if(!target)return {success:false,message:'ID wajib diisi.'};
  if(!pin||String(pin).length<4)return {success:false,message:'PIN minimal 4 digit.'};

  for(let i=1;i<data.length;i++){
    if(String(data[i][0]||'').trim().toLowerCase()===target.toLowerCase()){
      return {success:false,message:'ID sudah terdaftar.'};
    }
  }

  const headers=data[0].map(function(h){
    return String(h).trim();
  });

  const row=new Array(headers.length).fill(false);

  row[0]=target;
  row[1]=hashPin_(pin);

  const statusIndex=headers.findIndex(function(h){
    return h.toLowerCase()==='status';
  });

  if(statusIndex===-1){
    return {success:false,message:'Kolom Status tidak ditemukan.'};
  }

  row[statusIndex]='Aktif';

  for(let i=2;i<headers.length;i++){
    if(!headers[i])continue;
    if(headers[i].toLowerCase()==='status')continue;
    row[i]=permissions&&permissions[headers[i]]===true;
  }

  let targetRow=-1;

  for(let i=1;i<data.length;i++){
    if(String(data[i][0]||'').trim()===''){
      targetRow=i+1;
      break;
    }
  }

  if(targetRow===-1){
    targetRow=Math.max(sheet.getLastRow()+1,2);
  }

  sheet.getRange(targetRow,1,1,row.length).setValues([row]);
  SpreadsheetApp.flush();

  return {
    success:true,
    message:'User berhasil ditambahkan.'
  };
}

function updateUser(token,userId,permissions){
  if(!hasPermission(token,'Admin'))return {success:false,message:'Akses ditolak.'};
  const sheet=getUserAccessSheet_();
  const data=sheet.getDataRange().getValues();
  const target=String(userId||'').trim().toLowerCase();
  if(!target)return {success:false,message:'ID wajib diisi.'};
  const headers=data[0].map(function(h){return String(h).trim();});
  for(let i=1;i<data.length;i++){
    if(String(data[i][0]||'').trim().toLowerCase()===target){
      for(let j=2;j<headers.length;j++){
        if(!headers[j]||headers[j].toLowerCase()==='status')continue;
        sheet.getRange(i+1,j+1).setValue(permissions&&permissions[headers[j]]===true);
      }
      return {success:true,message:'Permission berhasil diperbarui.'};
    }
  }
  return {success:false,message:'ID tidak ditemukan.'};
}

function resetUserPin(token,userId,pin){
  if(!hasPermission(token,'Admin'))return {success:false,message:'Akses ditolak.'};
  return setUserPin(userId,pin);
}

function deleteUser(token,userId){
  const session=getLoginSession(token);
  if(!session.success)return {success:false,message:'Belum login.'};
  if(session.permissions.Admin!==true)return {success:false,message:'Akses ditolak.'};
  const target=String(userId||'').trim().toLowerCase();
  if(!target)return {success:false,message:'ID wajib diisi.'};
  if(target===String(session.userId||'').trim().toLowerCase()){
    return {success:false,message:'User yang sedang login tidak dapat dihapus.'};
  }
  const sheet=getUserAccessSheet_();
  const data=sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0]||'').trim().toLowerCase()===target){
      sheet.deleteRow(i+1);
      return {success:true,message:'User berhasil dihapus.'};
    }
  }
  return {success:false,message:'ID tidak ditemukan.'};
}

function setUserStatus(token,userId,status){
  if(!hasPermission(token,'Admin'))return {success:false,message:'Akses ditolak.'};
  const sheet=getUserAccessSheet_();
  const data=sheet.getDataRange().getValues();
  const headers=data[0].map(function(h){return String(h).trim();});
  const statusIndex=headers.findIndex(function(h){return h.toLowerCase()==='status';});
  if(statusIndex===-1)return {success:false,message:'Kolom Status tidak ditemukan.'};
  const target=String(userId||'').trim().toLowerCase();
  const newStatus=String(status||'').trim();
  if(!target)return {success:false,message:'ID wajib diisi.'};
  if(newStatus!=='Aktif'&&newStatus!=='Nonaktif')return {success:false,message:'Status tidak valid.'};
  const session=getLoginSession(token);
  if(!session.success)return {success:false,message:'Belum login.'};
  if(target===String(session.userId||'').trim().toLowerCase()&&newStatus==='Nonaktif'){
    return {success:false,message:'User yang sedang login tidak dapat dinonaktifkan.'};
  }
  for(let i=1;i<data.length;i++){
    if(String(data[i][0]||'').trim().toLowerCase()===target){
      sheet.getRange(i+1,statusIndex+1).setValue(newStatus);
      return {success:true,message:'Status user berhasil diubah menjadi '+newStatus+'.'};
    }
  }
  return {success:false,message:'ID tidak ditemukan.'};
}
