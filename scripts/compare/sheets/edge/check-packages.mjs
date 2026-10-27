import { readFile, readdir, writeFile, mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
const site=resolve(import.meta.dirname,'../site');
const out=resolve(process.argv[2]);
await mkdir(out,{recursive:true});
const packages=[];
const hits=[];
const interesting=/importXLSX|importCSV|importSheetToSnapshot|importSheetToUnitId|createWorkbook\(/;
async function walk(path){
  for(const ent of await readdir(path,{withFileTypes:true})){
    const p=resolve(path,ent.name);
    if(ent.isDirectory())await walk(p);
    else if(ent.name.endsWith('.d.ts')){
      const text=await readFile(p,'utf8');
      text.split('\n').forEach((line,i)=>{if(interesting.test(line))hits.push({file:p.slice(site.length+1),line:i+1,text:line.trim().slice(0,300)});});
    }
  }
}
for(const name of (await readdir(resolve(site,'node_modules/@univerjs'))).sort()){
  const root=resolve(site,'node_modules/@univerjs',name);
  const p=JSON.parse(await readFile(resolve(root,'package.json'),'utf8'));
  packages.push({name:p.name,version:p.version,license:p.license,dependencies:p.dependencies});
  await walk(resolve(root,'lib/types')).catch(error=>hits.push({package:p.name,error:String(error)}));
}
const result={date:new Date().toISOString(),packages,hits};
await writeFile(resolve(out,'package-surface.json'),JSON.stringify(result,null,1));
console.log(JSON.stringify({packages:packages.map(p=>`${p.name}@${p.version}`),hits},null,1));
