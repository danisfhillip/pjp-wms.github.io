/********* AUTH **********/
function getUserAccessSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('USER_ACCESS');
  if(!sheet)throw new Error('Sheet USER_ACCESS tidak ditemukan.');
  return sheet;
}

function getAuthSalt_(){
  const props=PropertiesService.getScriptProperties();
  let salt=props.getProperty('AUTH_SALT');
  if(!salt){
    salt=Utilities.getUuid();
    props.setProperty('AUTH_SALT',salt);
  }
  return salt;
}

function hashPin_(pin){
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(pin)+getAuthSalt_(),Utilities.Charset.UTF_8);
  return bytes.map(function(b){return ('0'+(b&255).toString(16)).slice(-2);}).join('');
}

function setUserPin(userId,pin){
  const sheet=getUserAccessSheet_();
  const data=sheet.getDataRange().getValues();
  const target=String(userId||'').trim().toLowerCase();
  if(!target)return {success:false,message:'ID wajib diisi.'};
  if(!pin||String(pin).length<4)return {success:false,message:'PIN minimal 4 digit.'};

  for(let i=1;i<data.length;i++){
    const rowId=String(data[i][0]||'').trim().toLowerCase();
    if(rowId===target){
      sheet.getRange(i+1,2).setValue(hashPin_(pin));
      return {success:true,message:'PIN berhasil disimpan.'};
    }
  }
  return {success:false,message:'ID tidak ditemukan di USER_ACCESS.'};
}

function loginUser(userId,pin){
  const sheet=getUserAccessSheet_();
  const data=sheet.getDataRange().getValues();
  if(data.length<2)return {success:false,message:'USER_ACCESS belum memiliki user.'};

  const headers=data[0].map(function(h){return String(h).trim();});
  const target=String(userId||'').trim().toLowerCase();
  const inputPin=String(pin||'');

  if(!target||!inputPin)return {success:false,message:'ID dan PIN wajib diisi.'};

  let userRow=null;

  for(let i=1;i<data.length;i++){
    const rowId=String(data[i][0]||'').trim().toLowerCase();
    if(rowId===target){
      userRow=data[i];
      break;
    }
  }

  if(!userRow)return {success:false,message:'ID atau PIN salah.'};

  const userStatus=String(userRow[12]||'Aktif').trim();
  if(userStatus.toLowerCase()!=='aktif')return {success:false,message:'Akun Anda sedang dinonaktifkan. Silakan hubungi Admin.'};

  const storedPin=String(userRow[1]||'').trim();
  if(!storedPin)return {success:false,message:'PIN user belum dibuat.'};

  if(hashPin_(inputPin)!==storedPin)return {success:false,message:'ID atau PIN salah.'};

  const permissions={};

  for(let i=2;i<headers.length;i++){
    if(!headers[i])continue;
    permissions[headers[i]]=userRow[i]===true;
  }

  const token=Utilities.getUuid()+'-'+Utilities.getUuid();
  const session={
    userId:target,
    permissions:permissions,
    loginAt:new Date().getTime()
  };

  CacheService.getScriptCache().put('AUTH_'+token,JSON.stringify(session),21600);

  return {
    success:true,
    token:token,
    userId:target,
    permissions:permissions
  };
}

function getLoginSession(token){
  if(!token)return {success:false};

  const cache=CacheService.getScriptCache();
  const cacheKey='AUTH_'+token;
  const value=cache.get(cacheKey);
  if(!value)return {success:false};

  const session=JSON.parse(value);

  // Cek tanggal login

  const timeZone=Session.getScriptTimeZone();
  const loginDate=Utilities.formatDate(new Date(session.loginAt),timeZone,'yyyy-MM-dd');
  const today=Utilities.formatDate(new Date(),timeZone,'yyyy-MM-dd');

  if(loginDate!==today){
    cache.remove(cacheKey);
    return {
      success:false,
      reason:'SESSION_EXPIRED',
      message:'Session telah berakhir. Silakan login kembali.'
    };
  }

  // Cek user

  const sheet=getUserAccessSheet_();
  const data=sheet.getDataRange().getValues();

  if(data.length<2){
    cache.remove(cacheKey);
    return {success:false};
  }

  const headers=data[0].map(function(h){return String(h).trim();});
  let userRow=null;

  for(let i=1;i<data.length;i++){
    if(String(data[i][0]||'').trim().toLowerCase()===session.userId){
      userRow=data[i];
      break;
    }
  }

  if(!userRow){
    cache.remove(cacheKey);
    return {
      success:false,
      reason:'USER_DELETED',
      message:'Akun Anda sudah dihapus.'
    };
  }

  // Cek status user

  const userStatus=String(userRow[12]||'Aktif').trim();

  if(userStatus.toLowerCase()!=='aktif'){
    cache.remove(cacheKey);
    return {
      success:false,
      reason:'USER_DISABLED',
      message:'Akun Anda telah dinonaktifkan oleh Admin.'
    };
  }

  // Update permission terbaru

  const permissions={};

  for(let i=2;i<headers.length;i++){
    if(headers[i])permissions[headers[i]]=userRow[i]===true;
  }

  session.permissions=permissions;

  cache.put(cacheKey,JSON.stringify(session),21600);

  return {
    success:true,
    userId:session.userId,
    permissions:permissions,
    loginAt:session.loginAt
  };
}

function getCurrentLoginUser(token){
  const session=getLoginSession(token);
  if(!session.success)return {success:false,message:'Session tidak valid.'};
  return {success:true,userId:session.userId,permissions:session.permissions};
}

function logoutUser(token){
  if(token)CacheService.getScriptCache().remove('AUTH_'+token);
  return {success:true};
}

function hasPermission(token,feature){
  const session=getLoginSession(token);
  if(!session.success)return false;
  return session.permissions[feature]===true;
}

function setupFirstPin(){
  setUserPin('PJP.Danis','123456');
  return setUserPin('PJP.RNW','123456');
}

/********** Cek Akses **********/
function checkPageAccess(token,feature){
  const session=getLoginSession(token);
  if(!session.success)return {success:false,allowed:false,message:'Belum login.'};
  const allowed=session.permissions[feature]===true;
  return {success:true,allowed:allowed,userId:session.userId,message:allowed?'Akses diizinkan.':'Akses ditolak.'};
}
