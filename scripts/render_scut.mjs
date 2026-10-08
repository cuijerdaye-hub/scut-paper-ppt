import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {pathToFileURL, fileURLToPath} from 'node:url';
import {createRequire} from 'node:module';
import {execFileSync} from 'node:child_process';

const argv=process.argv.slice(2);
const arg=(k,d)=>{const i=argv.indexOf(k);return i>=0?argv[i+1]:d};
if(!arg('--plan') || !arg('--output')) throw Error('Usage: render_scut.mjs --plan PLAN.json --output DRAFT.pptx [--previews DIR]');
const planPath=path.resolve(arg('--plan'));
const output=path.resolve(arg('--output'));
const previewDir=path.resolve(arg('--previews',path.join(path.dirname(output),'previews')));
const root=path.join(os.homedir(),'.cache/codex-runtimes/codex-primary-runtime/dependencies');
const pkg=path.join(arg('--node-modules',path.join(root,'node/node_modules')),'@oai/artifact-tool');
const {Presentation,PresentationFile}=await import(pathToFileURL(path.join(pkg,'dist/artifact_tool.mjs')).href);
const require=createRequire(path.join(pkg,'package.json'));
const {Canvas}=require('skia-canvas');
const ctx=new Canvas(1,1).getContext('2d');
const plan=JSON.parse((await fs.readFile(planPath,'utf8')).replace(/^\uFEFF/,''));
if(!argv.includes('--replace-draft') && await fs.stat(output).then(()=>true,()=>false))throw Error('Output already exists; choose a new draft path or explicitly use --replace-draft.');
const C={navy:'#2D4F8F',dark:'#163553',blue:'#76ABDC',pale:'#EDF3FA',ink:'#1B2C42',muted:'#5D6D80',line:'#D4DFEC',white:'#FFFFFF'};
const FONT='Microsoft YaHei';
const audit=[]; let currentSlide=0;
function measure(text,size,bold=false){ctx.font=`${bold?'bold ':''}${size}px "${FONT}"`; return ctx.measureText(text).width;}
function wrap(text,width,size,bold=false){
  return String(text).split('\n').flatMap(p=>{
    const toks=p.match(/[A-Za-z0-9][A-Za-z0-9.,:;()\/_–—+%-]*|\s+|[^\s]/gu)||[''];
    const out=[];let line='';
    for(const tok of toks){if(line&&measure(line+tok,size,bold)>width){ if(/^[，。；：、！？）】》”’]$/.test(tok)){line+=tok;out.push(line.trim());line='';}else{out.push(line.trim());line=tok.trimStart();}}else line+=tok;}
    if(line||!out.length)out.push(line.trim());return out;
  });
}
function text(s,txt,x,y,w,h,size=28,{bold=false,color=C.ink,align='left',name='text',leading=1.30}={}){
  if(typeof txt!=='string' || !txt.trim())throw Error(`Empty text: ${name}`);
  const lines=wrap(txt,w-4,size,bold); const needed=lines.length*size*leading;
  if(needed>h+1)throw Error(`Text overflow ${name}: ${needed.toFixed(1)} > ${h}; ${lines.join(' | ')}`);
  if(x<0||y<0||x+w>1281||y+h>721)throw Error(`Out of bounds: ${name}`);
  const sh=s.shapes.add({geometry:'textbox',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
  sh.text=lines.join('\n');sh.text.style={typeface:FONT,fontSize:size,bold,color,alignment:align,verticalAlignment:'top',autoFit:'none',wrap:'none',lineSpacing:1.10,insets:{left:0,right:0,top:0,bottom:0}};
  audit.push({slide:currentSlide,name,text:txt,lines:lines.length,fontPx:size,position:{x,y,w,h},estimatedHeight:needed});return sh;
}
function rule(s,x,y,w,color=C.line){s.shapes.add({geometry:'line',name:'divider',position:{left:x,top:y,width:w,height:0},fill:'none',line:{style:'solid',fill:color,width:1}});}
const assetBase=path.resolve(path.dirname(planPath),plan.assets?.directory||'assets');
const logo=await fs.readFile(path.join(assetBase,plan.assets?.logo||'scut-logo-blue.png'));
const campus=await fs.readFile(path.join(assetBase,plan.assets?.cover||'scut-campus.jpg'));
function brand(s){s.images.add({blob:logo,contentType:'image/png',fit:'contain',alt:'华南理工大学校名与校徽，取自用户华工蓝模板',position:{left:989,top:33,width:220,height:61}});}
function base(d,i){
 const s=p.slides.add();s.background.fill=C.white;brand(s);
 text(s,d.title,72,58,892,74,42,{bold:true,color:C.navy,name:'title'});
 if(d.subtitle)text(s,d.subtitle,74,143,1110,55,24,{color:C.muted,name:'subtitle'});
 rule(s,72,645,1136);
 text(s,d.citation||d.source,74,659,1040,35,14,{color:C.muted,name:'source'});
 text(s,String(i+1).padStart(2,'0'),1156,657,50,37,17,{color:C.navy,align:'right',name:'page'});
 return s;
}
function note(s,d){s.speakerNotes.textFrame.setText(`${d.notes||''}\n\n来源：${d.source}\n${plan.reference||''}\n${plan.doi||''}`);}
function table(s,{values,x=72,y=216,w=1136,h=350,widths,font=24,header=true}){
 if(!values?.length||values.some(r=>r.length!==values[0].length))throw Error('Invalid table matrix');
 const rows=values.length,cols=values[0].length;
 if(values.some(r=>r.some(c=>c==null || !String(c).trim())))throw Error('Empty table cell');
 const ws=widths||Array(cols).fill(w/cols);if(ws.length!==cols)throw Error('Table column count does not match layout');
 const tb=s.tables.add({rows,columns:cols,left:x,top:y,width:w,height:h,columnWidths:ws,values});
 const headerH=header?62:h/rows,bodyH=header?(h-headerH)/(rows-1):h/rows;
 for(let r=0;r<rows;r++){
   tb.rows[r].height=header&&r===0?headerH:bodyH;
   for(let c=0;c<cols;c++){
     const isHead=header&&r===0;const sz=isHead?font:font;
     const lines=wrap(String(values[r][c]),ws[c]-28,sz,isHead||c===0);
     if(lines.length*sz*1.30>(isHead?headerH:bodyH)-16)throw Error(`Table overflow r${r} c${c}: ${values[r][c]}`);
     const cell=tb.getCell(r,c);cell.value=lines.join('\n');
     cell.fill=isHead?C.navy:(r%2?C.white:C.pale);
     cell.text.style={typeface:FONT,fontSize:sz,bold:isHead||c===0,color:isHead?C.white:C.ink,autoFit:'none',lineSpacing:1.08};
   }
 }
 tb.cells.block({row:0,column:0,rowCount:rows,columnCount:cols}).assign({margins:{left:14,right:14,top:10,bottom:8},anchor:'center'});
 // Apply final font and fills after range styling, which can set defaults.
 for(let r=0;r<rows;r++)for(let c=0;c<cols;c++){
  const cell=tb.getCell(r,c);const hd=header&&r===0;
  cell.fill=hd?C.navy:(r%2?C.white:C.pale);cell.text.style={typeface:FONT,fontSize:font,bold:hd||c===0,color:hd?C.white:C.ink,autoFit:'none',lineSpacing:1.08};
 }
 tb.borders.assign({style:'solid',fill:C.line,width:1});
 return tb;
}
function takeaway(s,txt,y=586){text(s,txt,74,y,1130,49,26,{bold:true,color:C.navy,name:'takeaway'});}
function rows(s,d,{start=225,gap=124,labelW=240,bodyX=372,bodyW=818,sz=27}={}){
 d.items.forEach((a,j)=>{const y=start+j*gap; text(s,a.label,74,y,labelW,60,30,{bold:true,color:C.navy,name:`row-${j}-label`});text(s,a.body,bodyX,y+2,bodyW,gap-26,sz,{name:`row-${j}-body`});if(j<d.items.length-1)rule(s,74,y+gap-20,1130);});
}
if(!Array.isArray(plan.slides)||!plan.slides.length)throw Error('Missing slides');
const allowed=new Set(['cover','intro','comparison','decision','process','screening','synthesis','quality','contribution','summary']);
for(const d of plan.slides){if(!allowed.has(d.layout))throw Error(`Unknown layout ${d.layout}`);if(!d.source?.trim())throw Error(`Missing source for ${d.title}`);if(!d.title?.trim())throw Error('Missing title');}
const itemCounts={intro:3,decision:3,process:4,contribution:3,summary:3};
for(const d of plan.slides){if(itemCounts[d.layout] && d.items?.length!==itemCounts[d.layout])throw Error('Layout item count mismatch: '+d.layout);if(d.layout==='screening' && (d.leftItems?.length!==2 || d.rightItems?.length!==2))throw Error('screening requires two items per column');}
if(new Set(plan.slides.map(s=>s.title)).size!==plan.slides.length)throw Error('Duplicate slide titles');
const p=Presentation.create({slideSize:{width:1280,height:720}});
for(const [i,d] of plan.slides.entries()){
 currentSlide=i+1;
 let s;
 if(d.layout==='cover'){
  s=p.slides.add();s.background.fill=C.white;
  s.images.add({blob:campus,contentType:'image/jpeg',fit:'cover',alt:'原华工蓝模板校园建筑照片',position:{left:813,top:0,width:467,height:720}});
  s.images.add({blob:logo,contentType:'image/png',fit:'contain',alt:'华南理工大学',position:{left:72,top:53,width:285,height:79}});
  text(s,d.title,72,238,690,198,64,{bold:true,color:C.navy,name:'cover-title'});
  text(s,d.subtitle,76,451,652,92,26,{color:C.muted,name:'cover-subtitle'});
  text(s,d.authorLine,76,591,650,47,22,{color:C.navy,name:'cover-author'});
  text(s,d.journalLine,76,639,665,37,17,{color:C.muted,name:'cover-citation'});
 }else{
  s=base(d,i);
  switch(d.layout){
  case 'intro':
   text(s,d.statement,74,235,470,212,48,{bold:true,color:C.navy,name:'statement'});
   d.items.forEach((a,j)=>{const y=224+j*119;text(s,a.label,640,y,555,47,28,{bold:true,color:C.navy,name:`intro-${j}-label`});text(s,a.body,640,y+49,555,69,25,{name:`intro-${j}-body`});if(j<d.items.length-1)rule(s,640,y+107,554);});
   break;
  case 'comparison':table(s,{values:d.values,y:215,h:355,widths:[140,332,332,332],font:23});takeaway(s,d.takeaway,587);break;
  case 'decision':rows(s,d,{start:223,gap:113,labelW:288,bodyX:401,bodyW:791,sz:27});takeaway(s,d.takeaway,586);break;
  case 'process':
   d.items.forEach((a,j)=>{const x=74+j*289;text(s,String(j+1).padStart(2,'0'),x,226,243,86,61,{color:C.blue,bold:true,name:`step-${j}-num`});text(s,a.label,x,324,248,50,32,{bold:true,color:C.navy,name:`step-${j}-label`});text(s,a.body,x,394,243,166,26,{name:`step-${j}-body`});});takeaway(s,d.takeaway,586);break;
  case 'screening':
   text(s,d.leftTitle,74,225,503,59,31,{bold:true,color:C.navy,name:'left-title'});
   d.leftItems.forEach((a,j)=>{text(s,a.label,75,305+j*132,507,41,26,{bold:true,name:`left-${j}-label`});text(s,a.body,75,350+j*132,507,80,25,{color:C.muted,name:`left-${j}-body`});});
   text(s,d.rightTitle,660,225,536,59,31,{bold:true,color:C.navy,name:'right-title'});
   d.rightItems.forEach((a,j)=>{text(s,a.label,660,305+j*132,537,41,26,{bold:true,name:`right-${j}-label`});text(s,a.body,660,350+j*132,537,80,25,{color:C.muted,name:`right-${j}-body`});});break;
  case 'synthesis':table(s,{values:d.values,y:219,h:327,widths:[245,446,445],font:24});takeaway(s,d.takeaway,582);break;
  case 'quality':table(s,{values:d.values,y:216,h:386,widths:[209,463,464],font:24});break;
  case 'contribution':rows(s,d,{start:224,gap:117,labelW:271,bodyX:394,bodyW:798,sz:27});takeaway(s,d.takeaway,586);break;
  case 'summary':
   text(s,d.statement,74,224,1105,126,47,{bold:true,color:C.navy,name:'summary-statement'});
   d.items.forEach((a,j)=>{text(s,a.label,75,389+j*70,242,47,28,{bold:true,color:C.navy,name:`summary-${j}-label`});text(s,a.body,360,389+j*70,829,53,27,{name:`summary-${j}-body`});});break;
  }
 }
 note(s,d);
}
await fs.mkdir(path.dirname(output),{recursive:true});await fs.mkdir(previewDir,{recursive:true});
await (await PresentationFile.exportPptx(p)).save(output);
execFileSync(path.join(root,'python/python.exe'),[path.join(path.dirname(fileURLToPath(import.meta.url)),'style_native_tables.py'),output],{stdio:'inherit'});
await fs.writeFile(path.join(previewDir,'text-layout-audit.json'),JSON.stringify(audit,null,2));
for(let i=0;i<p.slides.items.length;i++){
 const slide=p.slides.items[i];const png=await p.export({slide,format:'png',scale:1.5});
 await fs.writeFile(path.join(previewDir,`slide-${String(i+1).padStart(2,'0')}.png`),new Uint8Array(await png.arrayBuffer()));
 const l=await slide.export({format:'layout'});await fs.writeFile(path.join(previewDir,`slide-${String(i+1).padStart(2,'0')}.layout.json`),await l.text());
}
console.log(JSON.stringify({output,previewDir,slides:p.slides.items.length}));
