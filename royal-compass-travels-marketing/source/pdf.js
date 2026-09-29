const { chromium } = require('/opt/node22/lib/node_modules/playwright');
(async()=>{
 const b=await chromium.launch(); const p=await b.newPage();
 await p.goto('file://'+process.cwd()+'/playbook_out.html',{waitUntil:'networkidle'});
 await p.evaluate(()=>document.fonts.ready);
 await p.pdf({path:'../out/Royal_Compass_Internal_Marketing_Playbook.pdf',format:'A4',printBackground:true,preferCSSPageSize:true,displayHeaderFooter:true,
  headerTemplate:'<div></div>',
  footerTemplate:'<div style="width:100%;font-size:7.5pt;color:#8a94a0;padding:0 18mm;display:flex;justify-content:space-between;font-family:Helvetica"><span>Royal Compass Travels · Internal Playbook · Not for client distribution</span><span class="pageNumber"></span></div>'});
 await b.close();})();
