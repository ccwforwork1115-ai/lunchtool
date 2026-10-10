const {chromium}=require('C:/Users/asd11/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const path=require('path');
const fs=require('fs');
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'});
 const page=await browser.newPage({viewport:{width:1440,height:1000}});
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto('file:///'+path.resolve('午餐明細帳轉換工具.html').replace(/\\/g,'/'));
 console.log('Startup errors:',errors);
 await page.evaluate(()=>localStorage.setItem('unrelated-test','keep'));
 await page.locator('#school-name').fill('測試學校');
 await page.locator('#save-school-preset').click();
 await page.locator('#storage-manager > summary').click();
 await page.locator('#storage-scope').selectOption('presets');
 page.once('dialog',d=>d.dismiss());
 await page.locator('#storage-clear').click();
 if(await page.locator('#school-presets option').count()!==2)throw Error('Cancel did not preserve data');
 page.once('dialog',d=>d.accept());
 await page.locator('#storage-clear').click();
 if(await page.locator('#school-presets option').count()!==1)throw Error('Presets were not cleared');
 if(await page.locator('#school-name').inputValue()!=='測試學校')throw Error('Preset clearing changed school');
 await page.locator('#storage-scope').selectOption('all');
 page.once('dialog',d=>d.accept());
 await page.locator('#storage-clear').click();
 const keys=await page.evaluate(()=>Object.keys(localStorage));
 if(keys.some(k=>k.startsWith('lunch-ledger-')))throw Error('Tool storage remains after clear all');
 if(await page.evaluate(()=>localStorage.getItem('unrelated-test'))!=='keep')throw Error('Unrelated storage cleared');
 if(await page.locator('#school-name').inputValue()!=='')throw Error('School not reset');
 await page.evaluate(()=>localStorage.removeItem('unrelated-test'));
 await page.locator('#storage-manager > summary').click();
 console.log('PASS: cancel clearing, selective presets reset, complete reset, unrelated storage preserved');
 await page.locator('#school-year').fill('115');
 await page.locator('#submit').click();
 if(!(await page.locator('#status').innerText()).includes('請先選取'))throw Error('Missing PDF feedback');
 if(await page.evaluate(()=>document.activeElement.id)!=='choose-pdfs')throw Error('Missing PDF focus');
 await page.locator('.settings > summary').filter({hasText:'月份與收費設定'}).click();
 await page.locator('#editing-month').selectOption('10');
 const visibleMonths=await page.locator('#c4-settings .month-card:visible').count();
 if(visibleMonths!==1)throw Error('Month switching failed: '+visibleMonths);
 await page.locator('.settings > summary').filter({hasText:'月份與收費設定'}).click();
 for(const width of [375,768,1024,1440]){
  await page.setViewportSize({width,height:1000});
  if(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth))throw Error('Page overflow at '+width);
  await page.locator('.settings > summary').filter({hasText:'月份與收費設定'}).click();
  if(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth))throw Error('Settings overflow at '+width);
  await page.locator('.settings > summary').filter({hasText:'月份與收費設定'}).click();
 }
 await page.setViewportSize({width:1440,height:1100});
 await page.screenshot({path:'介面預覽.png',fullPage:true});
 await page.locator('#pdfs').setInputFiles(['L20022  代扣午餐費.pdf','L30076  免費學生午餐補助款.pdf','L30077  精進學校菜單方案補助款.pdf']);
 await page.waitForFunction(()=>document.querySelector('#school-name').value.length>0,{timeout:30000});
 await page.locator('#school-year').fill('115');
 await page.locator('#submit').click();
 await page.waitForFunction(()=>!document.querySelector('#classification-review').hidden||!document.querySelector('#download-result').hidden||document.querySelector('#status').textContent.includes('轉換未完成'),{timeout:60000});
 const review=page.locator('#classification-review');
 if(await review.isVisible()){
  const selects=page.locator('#classification-rows select');
  console.log('Classification candidates:',await selects.count());
  for(const select of await selects.all()){
   const suggestion=await select.locator('xpath=../preceding-sibling::td[1]').innerText();
   const value=await select.locator('option').evaluateAll((options,suggestion)=>options.find(o=>o.value&&suggestion.includes(o.value))?.value||options.find(o=>o.value)?.value,suggestion);
   await select.selectOption(value);
  }
  await page.locator('#confirm-classifications').click();
  await page.waitForFunction(()=>!document.querySelector('#download-result').hidden||document.querySelector('#status').textContent.includes('轉換未完成'),{timeout:60000});
 }
 const status=await page.locator('#status').innerText();console.log(status);
 if(!await page.locator('#download-result').isVisible())throw Error('Conversion failed: '+status);
 const bytes=await page.evaluate(async()=>Array.from(new Uint8Array(await(await fetch(document.querySelector('#download-result').href)).arrayBuffer())));
 fs.writeFileSync('UI驗證產出.xlsx',Buffer.from(bytes));
 await page.emulateMedia({reducedMotion:'reduce'});
 if(await page.locator('#submit').evaluate(e=>getComputedStyle(e).transitionDuration)!=='0s')throw Error('Reduced motion failed');
 if(errors.length)throw Error(errors.join('\n'));
 console.log('PASS: initial state, missing-file focus, month switching, four viewport widths, real PDF conversion, reduced motion, no runtime errors');
 await browser.close();
})().catch(e=>{console.error(e);process.exit(1)});
