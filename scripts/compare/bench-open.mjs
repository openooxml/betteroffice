import { chromium } from 'playwright';
import { access, mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import { resolve, extname } from 'node:path';
import { createHash } from 'node:crypto';
import os from 'node:os';
import { richDocument } from './fixtures.ts';

const here = import.meta.dirname, env = process.env;
const out = resolve(here, env.CMP_OUT ?? 'out/docxperf');
const docs = (env.CMP_DOCS ?? 'public-800pp,syn800,long1500,long3000').split(',');
const reps = Number(env.CMP_N ?? 6), burn = Number(env.CMP_BURN ?? 1), cap = Number(env.CMP_OPEN_CAP ?? 600000);
const viewport = { width: 1180, height: 760 };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const bounded = async (p, ms, label) => {
  let t;
  return Promise.race([p,new Promise((_,r) => {t=setTimeout(()=>r(new Error(`${label} timeout ${ms}ms`)),ms);})]).finally(()=>clearTimeout(t));
};
await mkdir(resolve(out,'shots'),{recursive:true});
await mkdir(resolve(here,'gen'),{recursive:true});
if (docs.includes('syn800')) await writeFile(resolve(here,'gen/syn800.docx'),(await richDocument(374)).docx);
const files = {};
for (const doc of docs) {
  let path = resolve(here,`docs/${doc}.docx`);
  try { await access(path); } catch { path = resolve(here,`gen/${doc}.docx`); }
  const bytes = await readFile(path);
  files[doc]={ path,bytes:bytes.length,sha256:createHash('sha256').update(bytes).digest('hex') };
}
const server = Bun.serve({ hostname:'127.0.0.1',port:0,async fetch(req) {
  const path = decodeURIComponent(new URL(req.url).pathname);
  if (path.includes('..')) return new Response('bad',{status:400});
  const doc = /^\/docs\/(.*)\.docx$/.exec(path)?.[1];
  const file = Bun.file(doc && files[doc] ? files[doc].path : resolve(here,`dist/.${path}`));
  if (!await file.exists()) return new Response('missing',{status:404});
  const types={'.html':'text/html','.js':'text/javascript','.css':'text/css','.wasm':'application/wasm'};
  return new Response(file,{headers:{'Content-Type':types[extname(path)]??'application/octet-stream'}});
}});
let executablePath=env.CHROME;
if (!executablePath) for(const d of (await readdir('/opt/ms-playwright')).filter(x=>x.startsWith('chromium-')).sort().reverse()) {
  for(const b of ['chrome-linux/chrome','chrome-linux64/chrome']) {
    const p=resolve('/opt/ms-playwright',d,b);try {await access(p);executablePath=p;break;}catch{}
  }
  if(executablePath)break;
}
const browser=await chromium.launch({executablePath,headless:true,args:['--disable-background-timer-throttling','--disable-renderer-backgrounding','--disable-backgrounding-occluded-windows']});
const root=await browser.newBrowserCDPSession();
const lock=JSON.parse(await readFile(resolve(here,'site/package-lock.json'),'utf8'));
const data={meta:{date:new Date().toISOString(),head:env.CMP_HEAD,chromium:browser.version(),hostname:os.hostname(),firstPageOnly:env.CMP_FIRST_ONLY==='1',selectorCorrected:true,fontPreset:env.CMP_FONT_PRESET??'eager 20 defaults',cpus:os.cpus().length,viewport,reps,burn,cap,files,versions:Object.fromEntries(['@docx-editor.dev/core','@docx-editor.dev/react','@docx-editor.dev/fonts'].map(k=>[k,lock.packages[`node_modules/${k}`]?.version ?? null]))},sessions:[]};
data.meta.harnessFingerprints={};
for(const name of ['bench-open.mjs','fixtures.ts','long_doc.py','bo/bo-docx.tsx','site/dxe-docx.tsx','site/package.json'])data.meta.harnessFingerprints[name]=createHash('sha256').update(await readFile(resolve(here,name))).digest('hex');
if(env.CMP_PREPARE_FONTS==='1')data.meta.harnessFingerprints['site/public/preloaded-fonts.json']=createHash('sha256').update(await readFile(resolve(here,'site/public/preloaded-fonts.json'))).digest('hex');
const save=()=>writeFile(resolve(out,'open.json'),JSON.stringify(data,null,2));
async function memory(cdp) {
  const metrics=await cdp.send('Performance.getMetrics');
  const processes=(await root.send('SystemInfo.getProcessInfo')).processInfo.filter(p=>p.type==='renderer');
  const pss=[];
  for(const p of processes)try{const s=await readFile(`/proc/${p.id}/smaps_rollup`,'utf8');pss.push({pid:p.id,kb:Number(/^Pss:\s+(\d+)/m.exec(s)?.[1])});}catch{}
  return {jsHeapUsed:metrics.metrics.find(x=>x.name==='JSHeapUsedSize')?.value,rendererPssKb:pss.reduce((s,p)=>s+p.kb,0),renderers:pss};
}
for(const doc of docs)for(let rep=-burn;rep<reps;rep++)for(const arm of ((rep+burn)%2?['dxe','bo']:['bo','dxe'])) {
  const row={doc,rep,arm,burner:rep<0,at:new Date().toISOString(),load:os.loadavg(),errors:[]};data.sessions.push(row);
  const ctx=await browser.newContext({viewport,deviceScaleFactor:1});
  await ctx.addInitScript(()=>{
    window.edge={gaps:[],keys:[],scrolls:[],raf:[],t:performance.now()};
    let last=performance.now();
    setInterval(()=>{const now=performance.now();window.edge.gaps.push([now,now-last]);last=now;},50);
    addEventListener('keydown',e=>window.edge.keys.push([performance.now(),e.key,e.timeStamp]),true);
    addEventListener('wheel',e=>window.edge.scrolls.push([performance.now(),e.timeStamp]),true);
  });
  const page=await ctx.newPage();page.on('pageerror',e=>row.errors.push(e.message.slice(0,200)));
  const cdp=await ctx.newCDPSession(page);await cdp.send('Performance.enable');
  let inputTimer,shotPromise,openCastListener,lastOpenFrame;
  try {
    if(rep<0 && env.CMP_NO_OPEN_SHOTS!=='1') {
      openCastListener=event=>{lastOpenFrame={data:event.data,ts:event.metadata.timestamp*1000};cdp.send('Page.screencastFrameAck',{sessionId:event.sessionId}).catch(()=>{});};
      cdp.on('Page.screencastFrame',openCastListener);
      await cdp.send('Page.startScreencast',{format:'png',everyNthFrame:1});
    }
    const nav=Date.now();row.navEpoch=nav;
    await page.goto(`http://127.0.0.1:${server.port}/${arm}-docx.html?doc=${doc}`,{waitUntil:'commit',timeout:30000});
    inputTimer=setInterval(()=>{
      if(Date.now()-nav>cap)return;
      cdp.send('Input.dispatchKeyEvent',{type:'keyDown',key:'Shift',code:'ShiftLeft',windowsVirtualKeyCode:16,timestamp:Date.now()/1000}).catch(()=>{});
      cdp.send('Input.dispatchKeyEvent',{type:'keyUp',key:'Shift',code:'ShiftLeft',windowsVirtualKeyCode:16,timestamp:Date.now()/1000}).catch(()=>{});
      cdp.send('Input.dispatchMouseEvent',{type:'mouseWheel',x:1100,y:650,deltaX:0,deltaY:1}).catch(()=>{});
    },250);
    if(rep<0 && env.CMP_NO_OPEN_SHOTS!=='1')shotPromise=(async()=>{
      row.shots=[];
      for(const target of [10000,60000]) {
        await sleep(Math.max(0,nav+target-Date.now()));
        const frame=lastOpenFrame;
        const record={targetMs:target,observedMs:Date.now()-nav,frameEpoch:frame?.ts??null,frameAgeMs:frame?Date.now()-frame.ts:null};
        row.shots.push(record);
        if(frame)await writeFile(resolve(out,`shots/docxperf-${doc}-${arm}-${target/1000}s.png`),Buffer.from(frame.data,'base64'));
      }
    })().catch(e=>{row.shotError=e.message;});
    await bounded(page.waitForFunction(firstOnly=>window.cmp?.firstPageMs!==undefined && (firstOnly || window.cmp?.openMs!==undefined),env.CMP_FIRST_ONLY==='1',{timeout:cap,polling:100}),cap+1000,'open');
    row.readyObservedMs=Date.now()-nav;
    if(env.CMP_FIRST_ONLY==='1') {
      clearInterval(inputTimer);
      await sleep(100);
      row.open=await page.evaluate(()=>{
        const start=cmp.t0Epoch-cmp.timeOrigin,end=start+cmp.firstPageMs;
        const tasks=cmp.longtasks.filter(([s])=>s>=start&&s<=end);
        const gaps=edge.gaps.filter(([s])=>s>=start&&s<=end+100);
        const keys=edge.keys.filter(([s])=>s>=start&&s<=end+100);
        return {firstPageMs:cmp.firstPageMs,pages:cmp.pageCount(),info:cmp.info,longTaskCount:tasks.length,longTaskMaxMs:Math.max(0,...tasks.map(x=>x[1])),longTaskSumMs:tasks.reduce((a,b)=>a+b[1],0),blockingBeyond50Ms:tasks.reduce((a,b)=>a+Math.max(0,b[1]-50),0),timerGapMaxMs:Math.max(0,...gaps.map(x=>x[1])),keyLagMaxMs:Math.max(0,...keys.map(x=>x[0]-x[2])),longtasks:tasks,timerGaps:gaps,keys,wheel:edge.scrolls};
      });
      if(shotPromise)await shotPromise;
      if(rep<0){await page.screenshot({path:resolve(out,`shots/docxperf-${doc}-${arm}-first-page.png`),timeout:30000});row.firstPageShotCompletedMs=Date.now()-nav;}
      await bounded(page.evaluate(()=>cmp.layoutDone),cap,'layout before healthy jump');
      row.memoryOpen=await memory(cdp);
      row.healthyJump=await bounded(page.evaluate(async()=>{
        const candidates=[...document.querySelectorAll('*')].filter(e=>e.scrollHeight-e.clientHeight>200 && ['auto','scroll'].includes(getComputedStyle(e).overflowY)).sort((a,b)=>b.clientWidth*b.clientHeight-a.clientWidth*a.clientHeight);
        const scroller=candidates[0]??document.scrollingElement;
        scroller.scrollTop=0;await new Promise(requestAnimationFrame);await new Promise(requestAnimationFrame);
        const total=cmp.pageCount(),t=performance.now();scroller.scrollTop=scroller.scrollHeight-scroller.clientHeight;
        let targetVisible=false,crash=false;
        while(performance.now()-t<10000){
          await new Promise(requestAnimationFrame);
          crash=document.body.innerText.includes('Something went wrong') && document.body.innerText.includes('An error occurred while rendering this component');
          if(crash)break;
          const target=document.querySelector(`[data-page-index="${total-1}"]`);
          const r=target?.getBoundingClientRect();
          if(r && r.width>0 && r.bottom>180 && r.top<innerHeight){targetVisible=true;break;}
        }
        if(targetVisible){await new Promise(requestAnimationFrame);await new Promise(requestAnimationFrame);}
        return {ms:targetVisible?performance.now()-t:null,targetVisible,crash,total,top:scroller.scrollTop,range:scroller.scrollHeight-scroller.clientHeight,visibleRects:cmp.pageRects(),criterion:'last indexed page intersects viewport plus two rAF; no pixel verification'};
      }),30000,'healthy jump');
      if(rep<0)await page.screenshot({path:resolve(out,`shots/docxperf-${doc}-${arm}-healthy-jump.png`),timeout:30000});
    } else {
    await bounded(page.evaluate(()=>cmp.layoutDone),cap,'layout');
    clearInterval(inputTimer);
    await sleep(100);
    row.open=await page.evaluate(()=>{
      const start=cmp.t0Epoch-cmp.timeOrigin,readyEnd=start+Math.max(cmp.openMs,cmp.firstPageMs),end=start+Math.max(cmp.openMs,cmp.firstPageMs,cmp.layoutDoneMs??0);
      const tasks=cmp.longtasks.filter(([s])=>s>=start && s<=end).map(([,d])=>d);
      const gaps=edge.gaps.filter(([s])=>s>=start && s<=end+100);
      return {openMs:cmp.openMs,firstPageMs:cmp.firstPageMs,layoutMs:cmp.layoutDoneMs,pages:cmp.pageCount(),readyLongTaskMaxMs:Math.max(0,...cmp.longtasks.filter(([s])=>s>=start&&s<=readyEnd).map(([,d])=>d)),longTaskCount:tasks.length,longTaskMaxMs:Math.max(0,...tasks),longTaskSumMs:tasks.reduce((a,b)=>a+b,0),blockingBeyond50Ms:tasks.reduce((a,b)=>a+Math.max(0,b-50),0),timerGapMaxMs:Math.max(0,...gaps.map(x=>x[1])),keyEvents:edge.keys.length,keyLagMaxMs:Math.max(0,...edge.keys.filter(x=>x[0]>=start&&x[0]<=end+100).map(x=>x[0]-x[2])),keys:edge.keys.slice(0,40),wheel:edge.scrolls.slice(0,40),errors:cmp.errors,info:cmp.info};
    });
    if(row.open?.info?.layoutState && row.open.info.layoutState!=='complete'){await sleep(60000);row.afterInputStop=await page.evaluate(()=>({waitedMs:60000,pages:cmp.pageCount(),pending:document.querySelectorAll('[data-docx-layout-pending]').length}));}
    row.memoryOpen=await memory(cdp);
    if(shotPromise)await shotPromise;
    if(openCastListener){await cdp.send('Page.stopScreencast');cdp.off('Page.screencastFrame',openCastListener);openCastListener=null;}
    const castFrames=[];
    const frameListener=event=>{castFrames.push(event.metadata.timestamp*1000);cdp.send('Page.screencastFrameAck',{sessionId:event.sessionId}).catch(()=>{});};
    cdp.on('Page.screencastFrame',frameListener);
    await cdp.send('Page.startScreencast',{format:'jpeg',quality:30,everyNthFrame:1});
    row.scroll=await bounded(page.evaluate(async()=>{
      const candidates=[...document.querySelectorAll('*')].filter(e=>e.scrollHeight-e.clientHeight>200 && ['auto','scroll'].includes(getComputedStyle(e).overflowY)).sort((a,b)=>b.clientWidth*b.clientHeight-a.clientWidth*a.clientHeight);
      const scroller=candidates[0]??document.scrollingElement;
      const start=performance.now(),times=[];let last=start;
      await new Promise(done=>{
        const frame=()=>{const now=performance.now();times.push(now-last);last=now;const f=Math.min(1,(now-start)/10000);scroller.scrollTop=f*(scroller.scrollHeight-scroller.clientHeight);if(f<1)requestAnimationFrame(frame);else done();};requestAnimationFrame(frame);
      });
      return {durationMs:performance.now()-start,frames:times.length,rafPerSecond:1000*times.length/(performance.now()-start),framesOver50:times.filter(x=>x>50).length,maxFrameMs:Math.max(...times),scroller:scroller.className,range:scroller.scrollHeight-scroller.clientHeight};
    }),60000,'scroll');
    await cdp.send('Page.stopScreencast');
    cdp.off('Page.screencastFrame',frameListener);
    const intervals=castFrames.slice(1).map((t,i)=>t-castFrames[i]);
    row.scroll.screencast={frames:castFrames.length,spanMs:castFrames.length>1?castFrames.at(-1)-castFrames[0]:null,intervalsOver50:intervals.filter(x=>x>50).length,maxIntervalMs:intervals.length?Math.max(...intervals):null};
    row.uiCrash=await page.evaluate(()=>document.body.innerText.includes('Something went wrong') && document.body.innerText.includes('An error occurred while rendering this component'));
    if(row.uiCrash)throw new Error('rendering error boundary during scroll; post-scroll measurements invalid');
    row.memoryScrolled=await memory(cdp);
    row.jump=await bounded(page.evaluate(async()=>{
      const candidates=[...document.querySelectorAll('*')].filter(e=>e.scrollHeight-e.clientHeight>200 && ['auto','scroll'].includes(getComputedStyle(e).overflowY)).sort((a,b)=>b.clientWidth*b.clientHeight-a.clientWidth*a.clientHeight);
      const scroller=candidates[0]??document.scrollingElement;
      scroller.scrollTop=0;await new Promise(requestAnimationFrame);await new Promise(requestAnimationFrame);
      const t=performance.now();scroller.scrollTop=scroller.scrollHeight-scroller.clientHeight;await new Promise(requestAnimationFrame);await new Promise(requestAnimationFrame);
      return {twoRafMs:performance.now()-t,top:scroller.scrollTop,range:scroller.scrollHeight-scroller.clientHeight,visibleRects:cmp.pageRects()};
    }),60000,'jump');
    if(rep<0)await page.screenshot({path:resolve(out,`shots/docxperf-${doc}-${arm}-end.png`),timeout:30000}).catch(e=>row.errors.push(e.message));
    }
  } catch(e) {row.errors.push(e.message);row.timedOut=/timeout/i.test(e.message);}
  finally {clearInterval(inputTimer);if(shotPromise)await bounded(shotPromise,10000,'shot finish').catch(()=>{});await bounded(ctx.close(),30000,'close').catch(()=>{});}
  await save();console.log('OPEN',JSON.stringify(row));
}
await save();await browser.close();server.stop(true);
