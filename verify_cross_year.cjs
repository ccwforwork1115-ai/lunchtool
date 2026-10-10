const {chromium}=require('C:/Users/asd11/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const fs=require('fs'),path=require('path'),assert=require('assert');
(async()=>{
 const expected=JSON.parse(fs.readFileSync('跨年度測試/expected.json','utf8'));
 const browser=await chromium.launch({headless:true,executablePath:'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'});
 try {
 const page=await browser.newPage();
 await page.route('**/*',route=>decodeURIComponent(route.request().url()).endsWith('午餐明細帳轉換工具.html')?route.fulfill({contentType:'text/html',body:fs.readFileSync('午餐明細帳轉換工具.html','utf8').replace('</script></body>', '\nwindow.demoApi={parsePdf,convertLedgerFiles,decodeB64,DEFAULT_TEMPLATE};\n</script></body>')}):route.continue());
 await page.goto('file:///'+path.resolve('午餐明細帳轉換工具.html').replace(/\\/g,'/'));
 await page.waitForFunction(()=>window.demoApi);
 await page.locator('#school-year').fill('115');
 await page.locator('.settings').evaluateAll(elements=>elements.forEach(e=>e.open=true));
 await page.locator('#editing-month').selectOption('1');
 assert((await page.locator('#cross-year-reminder').innerText()).includes('116 年 1 月'));
 await page.locator('#school-year').fill('116');
 assert((await page.locator('#cross-year-reminder').innerText()).includes('117 年 1 月'));
 await page.locator('#school-year').fill('115');
 await page.locator('#editing-month').selectOption('8');
 assert(!await page.locator('#cross-year-reminder').isVisible());
 await page.locator('.settings').evaluateAll(elements=>elements.forEach(e=>e.open=false));
 await page.locator('#pdfs').setInputFiles([expected.pdfs[1]]);
 await page.locator('#submit').click();
 await page.waitForFunction(()=>!document.querySelector('#download-result').hidden||document.querySelector('#status').textContent.includes('轉換未完成'));
 assert((await page.locator('#cross-year-reminder').innerText()).includes('目前未選取前年度同一子目帳冊：免費學生午餐補助款'));
 await page.locator('#pdfs').setInputFiles(expected.pdfs);
 const result=await page.evaluate(async()=>{
   const {parsePdf,convertLedgerFiles,decodeB64,DEFAULT_TEMPLATE}=window.demoApi;
   const files=[...document.querySelector('#pdfs').files],parsed=[];
   for(const file of files)parsed.push({file,...await parsePdf(file)});
   const output=await convertLedgerFiles(files,decodeB64(DEFAULT_TEMPLATE),{schoolName:'嘉義縣民雄鄉興中國民小學',schoolYear:115,autoSchoolFromPdf:true},parsed);
   return {parsed:parsed.map(p=>({header:p.header,entries:p.entries})),balances:output.balances,bytes:Array.from(new Uint8Array(await output.blob.arrayBuffer()))};
 });
 const first=result.parsed.find(p=>p.header.year===115),second=result.parsed.find(p=>p.header.year===116);
 assert.equal(first.entries.length,57);assert.equal(second.entries.length,19);
 assert(!second.entries.some(e=>e.憑單.includes('*')));
 assert.equal(second.entries.reduce((s,e)=>s+e.貸方,0),75250);
 assert.equal(second.entries.at(-1).餘額,883052);
 fs.writeFileSync('跨年度測試/測試示意_115學年度產出.xlsx',Buffer.from(result.bytes));
 delete result.bytes;fs.writeFileSync('跨年度測試/解析結果.json',JSON.stringify(result,null,2));
 // Exercise the actual form and verify reverse file order is also supported.
 await page.locator('#pdfs').setInputFiles(expected.pdfs.slice().reverse());
 await page.locator('#submit').click();
 await page.waitForFunction(()=>!document.querySelector('#download-result').hidden||!document.querySelector('#classification-review').hidden||document.querySelector('#status').textContent.includes('轉換未完成'));
 assert(await page.locator('#download-result').isVisible(),await page.locator('#status').innerText());
 assert((await page.locator('#cross-year-reminder').innerText()).includes('已選取前年度同一子目帳冊'));
 for(const width of [375,768,1440]){
   await page.setViewportSize({width,height:1000});
   await page.locator('#cross-year-help').evaluate(e=>e.open=true);
   assert(!await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth));
 }
 await page.screenshot({path:'跨年度測試/跨年度提醒預覽.png',fullPage:true});
 // Regression: existing real PDFs must still parse and convert after the row fix.
 await page.locator('#pdfs').setInputFiles(['L20022  代扣午餐費.pdf','L30076  免費學生午餐補助款.pdf','L30077  精進學校菜單方案補助款.pdf']);
 await page.evaluate(async()=>{
   const {parsePdf,convertLedgerFiles,decodeB64,DEFAULT_TEMPLATE}=window.demoApi;
   const files=[...document.querySelector('#pdfs').files],parsed=[];
   for(const file of files)parsed.push({file,...await parsePdf(file)});
   const output=await convertLedgerFiles(files,decodeB64(DEFAULT_TEMPLATE),{schoolName:'嘉義縣民雄鄉興中國民小學',schoolYear:115,autoSchoolFromPdf:true},parsed);
   if(!output.blob.size)throw Error('Real PDF regression produced empty workbook');
 });
 console.log('PASS: 76 parsed transactions; starred opening excluded; Jan income 75,250; Jan PDF balance 883,052; actual form succeeds in reverse order');
 } finally {await browser.close();}
})().catch(e=>{console.error(e);process.exit(1)});
