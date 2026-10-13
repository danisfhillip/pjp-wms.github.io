function getRollForPrint(rollId){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName('ROLL');

    if(!sheet)return{success:false,message:'Sheet ROLL tidak ditemukan.'};

    rollId=String(rollId).trim().toLowerCase();

    if(!rollId)return{success:false,message:'Roll ID kosong.'};

    const data=sheet.getDataRange().getValues();

    if(data.length<2)return{success:false,message:'Sheet ROLL belum memiliki data.'};

    const headers=data[0].map(function(header){
      return String(header).trim();
    });

    function getIndex(name){
      return headers.findIndex(function(header){
        return header.toLowerCase()===name.toLowerCase();
      });
    }

    const rollIndex=getIndex('Roll ID');
    const fabricTypeIndex=getIndex('Jenis Kain');
    const productIndex=getIndex('Nama Produk');
    const designIndex=getIndex('Kode Desain');
    const colorIndex=getIndex('Warna');
    const widthIndex=getIndex('Lebar');
    const gsmIndex=getIndex('GSM');
    const meterIndex=getIndex('Meter');
    const yardIndex=getIndex('Yard');

    if(rollIndex===-1){
      return{
        success:false,
        message:'Kolom "Roll ID" tidak ditemukan.'
      };
    }

    for(let i=1;i<data.length;i++){
      const currentRoll=String(data[i][rollIndex]).trim().toLowerCase();

      if(currentRoll===rollId){

        const meter=
          meterIndex!==-1&&data[i][meterIndex]!==''
            ?Number(data[i][meterIndex])
            :'';

        const yard=
          yardIndex!==-1&&data[i][yardIndex]!==''
            ?Number(data[i][yardIndex])
            :'';

        let length='';
        let lengthUnit='';

        if(meter!==''){
          length=meter;
          lengthUnit='M';
        }else if(yard!==''){
          length=yard;
          lengthUnit='Yd';
        }

        return{
          success:true,
          data:{
            'Roll ID':String(data[i][rollIndex]),
            'Jenis Kain':fabricTypeIndex!==-1
              ?formatCellValue(data[i][fabricTypeIndex])
              :'',
            'Nama Produk':productIndex!==-1
              ?formatCellValue(data[i][productIndex])
              :'',
            'Kode Desain':designIndex!==-1
              ?formatCellValue(data[i][designIndex])
              :'',
            'Warna':colorIndex!==-1
              ?formatCellValue(data[i][colorIndex])
              :'',
            'Lebar':widthIndex!==-1
              ?formatCellValue(data[i][widthIndex])
              :'',
            'GSM':gsmIndex!==-1
              ?formatCellValue(data[i][gsmIndex])
              :'',
            'Length':length,
            'Length Unit':lengthUnit,
            'Print Date':Utilities.formatDate(
              new Date(),
              Session.getScriptTimeZone(),
              'dd/MMM/yyyy'
            )
          }
        };
      }
    }

    return{
      success:false,
      message:'Roll ID "'+rollId+'" tidak ditemukan.'
    };

  }catch(error){
    return{
      success:false,
      message:'Error getRollForPrint: '+error.message
    };
  }
}

function getRollsForPrint(rollIds){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName('ROLL');

    if(!sheet){
      return{
        success:false,
        message:'Sheet ROLL tidak ditemukan.'
      };
    }

    if(!Array.isArray(rollIds)||!rollIds.length){
      return{
        success:false,
        message:'Tidak ada Roll ID yang diproses.'
      };
    }

    const data=sheet.getDataRange().getValues();

    if(data.length<2){
      return{
        success:false,
        message:'Sheet ROLL belum memiliki data.'
      };
    }

    const headers=data[0].map(function(header){
      return String(header).trim();
    });

    function getIndex(name){
      return headers.findIndex(function(header){
        return header.toLowerCase()===name.toLowerCase();
      });
    }

    const rollIndex=getIndex('Roll ID');
    const fabricTypeIndex=getIndex('Jenis Kain');
    const productIndex=getIndex('Nama Produk');
    const designIndex=getIndex('Kode Desain');
    const colorIndex=getIndex('Warna');
    const widthIndex=getIndex('Lebar');
    const gsmIndex=getIndex('GSM');
    const meterIndex=getIndex('Meter');
    const yardIndex=getIndex('Yard');

    if(rollIndex===-1){
      return{
        success:false,
        message:'Kolom "Roll ID" tidak ditemukan.'
      };
    }

    const requested={};

    rollIds.forEach(function(rollId){
      const key=String(rollId).trim().toLowerCase();
      if(key)requested[key]=String(rollId).trim();
    });

    const found={};

    for(let i=1;i<data.length;i++){
      const currentRoll=String(data[i][rollIndex]||'').trim();
      const key=currentRoll.toLowerCase();

      if(!requested[key])continue;

      const meter=
        meterIndex!==-1&&data[i][meterIndex]!==''
          ?Number(data[i][meterIndex])
          :'';

      const yard=
        yardIndex!==-1&&data[i][yardIndex]!==''
          ?Number(data[i][yardIndex])
          :'';

      let length='';
      let lengthUnit='';

      if(meter!==''){
        length=meter;
        lengthUnit='M';
      }else if(yard!==''){
        length=yard;
        lengthUnit='Yd';
      }

      found[key]={
        success:true,
        data:{
          'Roll ID':currentRoll,
          'Jenis Kain':fabricTypeIndex!==-1
            ?formatCellValue(data[i][fabricTypeIndex])
            :'',
          'Nama Produk':productIndex!==-1
            ?formatCellValue(data[i][productIndex])
            :'',
          'Kode Desain':designIndex!==-1
            ?formatCellValue(data[i][designIndex])
            :'',
          'Warna':colorIndex!==-1
            ?formatCellValue(data[i][colorIndex])
            :'',
          'Lebar':widthIndex!==-1
            ?formatCellValue(data[i][widthIndex])
            :'',
          'GSM':gsmIndex!==-1
            ?formatCellValue(data[i][gsmIndex])
            :'',
          'Length':length,
          'Length Unit':lengthUnit,
          'Print Date':Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            'dd/MMM/yyyy'
          )
        }
      };
    }

    const results=rollIds.map(function(rollId){
      const original=String(rollId).trim();
      const key=original.toLowerCase();

      if(found[key]){
        return{
          rollId:original,
          success:true,
          data:found[key].data
        };
      }

      return{
        rollId:original,
        success:false,
        message:'Roll ID "'+original+'" tidak ditemukan.'
      };
    });

    return{
      success:true,
      results:results
    };

  }catch(error){
    return{
      success:false,
      message:'Error getRollsForPrint: '+error.message
    };
  }
}

function formatCellValue(value){
  if(value instanceof Date){
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'dd MMM yyyy, HH:mm'
    );
  }

  return String(value??'');
}
