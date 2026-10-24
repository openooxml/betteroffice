import { chromium } from 'playwright';
import { createServer } from 'node:http';
import { access, readFile, writeFile, mkdir, readdir } from 'node:fs/promises';
import { resolve, extname, relative } from 'node:path';
const out = resolve(process.argv[2]);
const work = resolve(out, 'work');
const dist = resolve(import.meta.dirname, '../dist');
const audit = JSON.parse(await readFile(resolve(out, 'audit.json'), 'utf8'));
const ids = process.env.CAPTURE_PROBES?.split(',') ?? [...new Set(audit.boOnlyCorrect.map(x => x.id))];
await mkdir(resolve(out, 'shots'), { recursive: true });
const types={'.html':'text/html','.js':'text/javascript','.css':'text/css','.json':'application/json','.wasm':'application/wasm'};
const server=createServer(async(req,res)=>{
  const url=new URL(req.url,'http://localhost');
  const base=url.pathname.startsWith('/docs/')?work:dist;
  let suffix=url.pathname;
  if(base===work){
    suffix=suffix.slice(6);
    if(suffix.endsWith('.xlsx')) suffix=suffix.slice(0,-5)+'/input.xlsx';
  }
  const path=resolve(base,base===dist?'.'+suffix:suffix);
  if(relative(base,path).startsWith('..')){res.writeHead(400).end();return;}
  try{const body=await readFile(path);res.writeHead(200,{'Content-Type':types[extname(path)]??'application/octet-stream'}).end(body);}
  catch{res.writeHead(404).end();}
});
await new Promise(done=>server.listen(0,'127.0.0.1',done));
const base=`http://127.0.0.1:${server.address().port}`;
const directories=(await readdir('/opt/ms-playwright')).filter(n=>n.startsWith('chromium-')).sort().reverse();
let executablePath;
for(const d of directories){for(const n of ['chrome-linux/chrome','chrome-linux64/chrome']){const p=resolve('/opt/ms-playwright',d,n);try{await access(p);executablePath=p;}catch{}if(executablePath)break;}if(executablePath)break;}
const browser=await chromium.launch({executablePath,headless:true,args:['--disable-background-timer-throttling','--disable-renderer-backgrounding']});
const records=[];
for(const id of ids){
  const expected=JSON.parse(await readFile(resolve(work,'sp',id,'expected.json'),'utf8'));
  const targets=[expected.cells[0]];
  for(const arm of ['bo','univer']){
    const ctx=await browser.newContext({viewport:{width:1180,height:760},deviceScaleFactor:1});
    const page=await ctx.newPage();
    const record={id,arm,formula:targets[0].formula,cells:targets.map(c=>c.cell),errors:[]};
    page.on('pageerror',e=>record.errors.push(String(e).slice(0,200)));
    try{
      if(arm==='bo'){
        await page.goto(`${base}/bo-xlsx.html?wb=sp/${id}`,{timeout:60000});
        await page.waitForFunction(()=>window.cmp,undefined,{timeout:60000});
        await Promise.race([page.evaluate(()=>window.cmp.ready),new Promise((_,fail)=>setTimeout(()=>fail(new Error('BO ready timeout')),60000))]);
        record.before=await page.evaluate(cell=>window.cmp.read(cell),targets[0].cell);
        await page.evaluate(()=>window.cmp.setA1(0));
        await page.evaluate(f=>window.cmp.setA1('='+f),record.formula);
        await page.waitForTimeout(1500);
        record.value=[];
        for(const cell of targets)record.value.push(await page.evaluate(a1=>window.cmp.read(a1),cell.cell));
        record.errors.push(...await page.evaluate(()=>window.cmp.errors));
      }else{
        await page.goto(`${base}/univer-calc.html?wb=sp/${id}/univer&cap=60000`,{timeout:60000});
        await page.waitForFunction(()=>window.calc,undefined,{timeout:60000});
        record.ready=await page.evaluate(()=>window.calc.ready);
        record.value=await page.evaluate(cells=>window.calc.read([{sheet:'Sheet1',cells}]),targets.map(c=>[c.r,c.c]));
        record.errors.push(...await page.evaluate(()=>window.calc.errors));
      }
      await page.screenshot({path:resolve(out,'shots',`sheets-formula-${id}-${arm}.png`)});
    }catch(error){record.error=String(error).slice(0,300);}
    records.push(record);
    await ctx.close();
  }
}
await writeFile(resolve(out,'capture.json'),JSON.stringify({chromium:browser.version(),viewport:[1180,760],records},null,1));
await browser.close();
server.close();
console.log(`captured ${records.length} arms across ${ids.length} formulas`);
process.exit(0);
