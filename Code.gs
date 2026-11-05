const TZ='Asia/Kolkata';
const PRODUCTS='Products', TX='Transactions', SETTINGS='Settings';

function doGet(e) {
  const p=(e && e.parameter) || {};
  if (p.action) {
    let result;
    try {
      if (p.action==='lookup') result=lookup_(p.barcode);
      else if (p.action==='save') result=save_(p);
      else result={ok:false,error:'Unknown action'};
    } catch(err) { result={ok:false,error:err.message}; }
    const body=JSON.stringify(result);
    const cb=String(p.callback||'').replace(/[^\w.$]/g,'');
    return ContentService.createTextOutput(cb ? cb+'('+body+');' : body)
      .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutput('LIQD LAB API is running.');
}

function setup(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const defs=[
    [PRODUCTS,['Barcode','SKU','Product Name','Variant / Size','Cost Price','Selling Price','Opening Stock','Active']],
    [TX,['Timestamp (IST)','Barcode','SKU','Product Name','Variant / Size','Transaction Type','Quantity','Unit Price','Total Amount','Stock After','Latitude','Longitude','Google Maps Link','Device','Notes']],
    [SETTINGS,['Setting','Value']]
  ];
  defs.forEach(([n,h])=>{let s=ss.getSheetByName(n)||ss.insertSheet(n);if(s.getLastRow()===0)s.appendRow(h);s.setFrozenRows(1);s.getRange(1,1,1,h.length).setFontWeight('bold');});
  const s=ss.getSheetByName(SETTINGS); if(s.getLastRow()===1){s.appendRow(['Timezone',TZ]);s.appendRow(['Business Name','LIQD LAB']);}
  const pr=ss.getSheetByName(PRODUCTS);
  if(pr.getLastRow()===1){pr.appendRow(['890000000001','LIQD-AVN-30','Creed Aventus','30ml',100,249,10,'Yes']);pr.appendRow(['890000000002','LIQD-OMB-30','Ombre Leather','30ml',100,249,10,'Yes']);}
}

function lookup_(barcode){
  barcode=String(barcode||'').trim(); if(!barcode)throw Error('Barcode required');
  const sh=SpreadsheetApp.getActive().getSheetByName(PRODUCTS),v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++){let r=v[i];if(String(r[0]).trim()===barcode&&String(r[7]).toLowerCase()!=='no'){
    return {ok:true,product:{barcode:String(r[0]),sku:String(r[1]),name:String(r[2]),variant:String(r[3]),cost:Number(r[4])||0,price:Number(r[5])||0,stock:stock_(barcode,Number(r[6])||0)}};
  }} throw Error('Barcode not found. Add it to Products first.');
}
function stock_(barcode,opening){
  const rows=SpreadsheetApp.getActive().getSheetByName(TX).getDataRange().getValues();let n=opening;
  for(let i=1;i<rows.length;i++){let r=rows[i];if(String(r[1])===barcode){let q=Number(r[6])||0,t=String(r[5]);if(t==='SALE')n-=q;else if(t==='RESTOCK'||t==='RETURN')n+=q;else if(t==='ADJUSTMENT')n=q;}}
  return n;
}
function save_(p){
  const lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const b=String(p.barcode||'').trim(),type=String(p.type||'').toUpperCase(),qty=Math.floor(Number(p.qty));
    if(!['SALE','RESTOCK','RETURN','ADJUSTMENT'].includes(type))throw Error('Invalid transaction type');
    if(!Number.isInteger(qty)||qty<1)throw Error('Quantity must be at least 1');
    const found=lookup_(b);if(!found.ok)throw Error('Product not found');
    const x=found.product,before=x.stock;let after=before;
    if(type==='SALE')after-=qty;else if(type==='RESTOCK'||type==='RETURN')after+=qty;else after=qty;
    if(after<0)throw Error('Insufficient stock. Available: '+before);
    const price=Number(p.price)>=0?Number(p.price):x.price;
    const amount=(type==='SALE'||type==='RETURN')?price*qty:0;
    const lat=p.lat===''?'':Number(p.lat),lng=p.lng===''?'':Number(p.lng);
    const map=(lat!==''&&lng!=='')?'https://www.google.com/maps?q='+lat+','+lng:'';
    const now=new Date();
    SpreadsheetApp.getActive().getSheetByName(TX).appendRow([
      Utilities.formatDate(now,TZ,'yyyy-MM-dd HH:mm:ss'),b,x.sku,x.name,x.variant,type,qty,price,amount,after,lat,lng,map,String(p.device||'').slice(0,250),String(p.notes||'').slice(0,500)
    ]);
    return {ok:true,timestamp:Utilities.formatDate(now,TZ,'dd-MMM-yyyy HH:mm:ss'),stock:after,amount:amount};
  }finally{lock.releaseLock();}
}
