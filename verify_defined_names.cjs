const {chromium}=require('C:/Users/asd11/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const fs=require('fs'),path=require('path'),assert=require('assert');
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'});
 try {
 const page=await browser.newPage();
 const html=fs.readFileSync('index.html','utf8').replace('</script></body>','\nwindow.repairApi={cleanInvalidDefinedNames,convertLedgerFiles,decodeB64,DEFAULT_TEMPLATE};\n</script></body>');
 await page.route('**/*',route=>route.request().url().startsWith('file:')?route.fulfill({contentType:'text/html',body:html}):route.continue());
 await page.goto('file:///'+path.resolve('index.html').replace(/\\/g,'/'));
 await page.waitForFunction(()=>window.repairApi);
 const result=await page.evaluate(async()=>{
 const {cleanInvalidDefinedNames,convertLedgerFiles,decodeB64,DEFAULT_TEMPLATE}=window.repairApi;
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 const doc=new DOMParser().parseFromString(`<workbook xmlns="${ns}"><definedNames><definedName name="_xlfn_SINGLE"/><definedName name="empty"> </definedName><definedName name="_xlfn_SINGLE">none</definedName><definedName name="Valid">'基本資料'!$A$1</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">'基本資料'!$A$1:$H$20</definedName></definedNames></workbook>`,'application/xml');
 cleanInvalidDefinedNames(doc);
 const kept=[...doc.getElementsByTagNameNS(ns,'definedName')].map(n=>({name:n.getAttribute('name'),value:n.textContent,scope:n.getAttribute('localSheetId')}));
 const source=await JSZip.loadAsync(decodeB64(DEFAULT_TEMPLATE));
 const sourceDoc=new DOMParser().parseFromString(await source.file('xl/workbook.xml').async('string'),'application/xml');
 const invalidBefore=[...sourceDoc.getElementsByTagNameNS(ns,'definedName')].filter(n=>!n.textContent.trim()).length;
 const output=await convertLedgerFiles([],decodeB64(DEFAULT_TEMPLATE),{schoolName:'嘉義縣民雄鄉興中國民小學',schoolYear:115},[]);
 const zip=await JSZip.loadAsync(await output.blob.arrayBuffer());
 const xml=await zip.file('xl/workbook.xml').async('string');
 const outDoc=new DOMParser().parseFromString(xml,'application/xml');
 if(outDoc.getElementsByTagName('parsererror').length)throw Error('Invalid XML');
 return {kept,invalidBefore,namesAfter:outDoc.getElementsByTagNameNS(ns,'definedName').length,sheetCount:outDoc.getElementsByTagNameNS(ns,'sheet').length};
 });
 assert.deepStrictEqual(result.kept,[{name:'Valid',value:"'基本資料'!$A$1",scope:null},{name:'_xlnm.Print_Area',value:"'基本資料'!$A$1:$H$20",scope:'0'}]);
 assert(result.invalidBefore>0);assert.equal(result.namesAfter,0);assert(result.sheetCount>20);
 console.log('PASS: invalid embedded template names removed; valid names and print area preserved; workbook XML parses; '+result.sheetCount+' worksheets retained');
 }finally{await browser.close()}
})().catch(e=>{console.error(e);process.exit(1)});