import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {pathToFileURL,fileURLToPath} from 'node:url';
import {createRequire} from 'node:module';
import {execFileSync} from 'node:child_process';
const args=process.argv.slice(2);const arg=(k)=>{const i=args.indexOf(k);return i>=0?args[i+1]:undefined;};
if(!arg('--plan')||!arg('--output'))throw Error('Usage: render_template.mjs --plan PLAN.json --output DRAFT.pptx');
const planPath=path.resolve(arg('--plan')),output=path.resolve(arg('--output'));
if(!args.includes('--replace-draft') && await fs.stat(output).then(()=>true,()=>false))throw Error('Output exists; use a new draft path.');
const root=path.join(os.homedir(),'.cache/codex-runtimes/codex-primary-runtime/dependencies');
const pkg=path.join(root,'node/node_modules/@oai/artifact-tool');
const {PresentationFile,FileBlob}=await import(pathToFileURL(path.join(pkg,'dist/artifact_tool.mjs')).href);
const req=createRequire(path.join(pkg,'package.json'));const {Canvas}=req('skia-canvas');
const measure=new Canvas(1,1).getContext('2d');
function fit(text,w,h,size,bold){
 measure.font=`${bold?'bold ':''}${size}px "Microsoft YaHei"`;
 let total=0;
 for(const para of text.split('\n')){
  let line='',n=1;
  for(const token of para.match(/[A-Za-z0-9][A-Za-z0-9.,:;()/_–—+%-]*|\s+|[^\s]/gu)||[]){
   if(line && measure.measureText(line+token).width>w+4){n++;line=token;}else line+=token;
  }total+=n;
 }
 if(total*size*1.25>h+3)throw Error(`Text capacity exceeded: ${total} lines at ${size}px need ${(total*size*1.25).toFixed(1)}px; slot ${h}px. ${text}`);
}
const privateDir=path.join(path.dirname(output),'.template-authoring',path.basename(output,'.pptx'));
await fs.mkdir(privateDir,{recursive:true});const editedPath=path.join(privateDir,'text-edited.pptx');
const plan=JSON.parse(await fs.readFile(planPath,'utf8'));
if(!plan.title?.trim() || !plan.reference?.trim())throw Error('Plan title and complete reference are required.');
if(!Array.isArray(plan.slides) || !plan.slides.length)throw Error('Plan must contain slides.');
plan.template=path.resolve(path.dirname(planPath),plan.template);
if(output===plan.template)throw Error('Output must differ from the source template.');
if(new Set(plan.slides.map(s=>s.templateSlide)).size!==plan.slides.length)throw Error('Duplicate templateSlide: duplicate the source page in a template copy first, then map the new page.');
const p=await PresentationFile.importPptx(await FileBlob.load(plan.template));
const inspected=(await p.inspect({kind:'textbox',maxChars:900000})).ndjson.split('\n').filter(Boolean).map(x=>JSON.parse(x));
for(const [index,d] of plan.slides.entries()){
 if(!d.source?.trim() || !d.notes?.trim())throw Error('Every slide requires a source and speaker notes.');
 const boxes=inspected.filter(x=>x.kind==='textbox'&&x.slide===d.templateSlide);
 const planned=new Set(d.edits.map(e=>e.name));
 if(planned.size!==d.edits.length)throw Error('A template slot may be edited only once per page.');
 const missing=boxes.filter(x=>x.text.trim()&&!planned.has(x.name));
 if(missing.length)throw Error('Unmapped template text '+d.templateSlide+': '+missing.map(x=>x.name).join(','));
 for(const e of d.edits){
  const found=boxes.filter(x=>x.name===e.name);if(found.length!==1)throw Error('Template slot not unique '+d.templateSlide+'/'+e.name);
  if(!Number.isFinite(e.fontSize)||e.fontSize<=0)throw Error('fontSize must be a positive number.');
  if(!e.text?.trim())throw Error('Empty planned text '+e.name);
  const box=found[0],sh=p.resolve(box.id);
  fit(e.text,e.position?.width||box.bbox[2],e.position?.height||e.height||box.bbox[3],e.fontSize,e.bold);sh.text=e.text;
  const style={typeface:'Microsoft YaHei',fontSize:e.fontSize,bold:e.bold??false,autoFit:'none',wrap:'square',lineSpacing:1.13,insets:{left:0,right:0,top:0,bottom:0}};
  if(e.color)style.color=e.color;if(e.align)style.alignment=e.align;if(e.vertical)style.verticalAlignment=e.vertical;
  sh.text.style=style;
  if(e.position||e.height)sh.position={left:box.bbox[0],top:box.bbox[1],width:box.bbox[2],height:e.height||box.bbox[3],...e.position};
 }
 p.slides.items[d.templateSlide-1].speakerNotes.textFrame.setText(`${d.notes}\n\n来源：${d.source}\n${plan.reference}\n${plan.doi||''}`);
}
await (await PresentationFile.exportPptx(p)).save(editedPath);
execFileSync(path.join(root,'python/python.exe'),[path.join(path.dirname(fileURLToPath(import.meta.url)),'preserve_template.py'),'--plan',planPath,'--edited',editedPath,'--output',output],{stdio:'inherit'});
console.log(JSON.stringify({output,slides:plan.slides.length,textSlots:plan.slides.reduce((n,s)=>n+s.edits.length,0)}));
