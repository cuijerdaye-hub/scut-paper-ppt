"""Prepare a native template copy with one independently editable source page per plan page."""
import argparse,copy,io,json,re,zipfile,importlib.util,contextlib
from pathlib import Path
from lxml import etree as E
ap=argparse.ArgumentParser();ap.add_argument('--plan',required=True);ap.add_argument('--output-dir',required=True);ap.add_argument('--pptx-skill',required=True);args=ap.parse_args()
plan_path=Path(args.plan).resolve();out=Path(args.output_dir).resolve();out.mkdir(parents=True,exist_ok=True)
plan=json.loads(plan_path.read_text(encoding='utf-8-sig'));source=(plan_path.parent/plan['template']).resolve();unpacked=out/'template-parts'
if unpacked.exists():raise FileExistsError('Use a fresh preparation directory')
unpacked.mkdir()
with zipfile.ZipFile(source) as z:
 for info in z.infolist():
  destination=(unpacked/info.filename).resolve()
  if not destination.is_relative_to(unpacked):raise ValueError('Unsafe package member')
 z.extractall(unpacked)
helper=Path(args.pptx_skill)/'scripts/add_slide.py';spec=importlib.util.spec_from_file_location('scut_add_slide',helper);mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
pres_path=unpacked/'ppt/presentation.xml';rels_path=unpacked/'ppt/_rels/presentation.xml.rels'
initial=E.parse(str(pres_path));ids=list(initial.find('p:sldIdLst',ns));rels=E.parse(str(rels_path));targets={r.get('Id'):r.get('Target') for r in rels.getroot()}
seen=set();mapping=[]
for index,slide in enumerate(plan['slides'],1):
 n=slide['templateSlide']
 if not isinstance(n,int) or n<1 or n>len(ids):raise ValueError('Invalid source slide')
 slide['originalTemplateSlide']=n
 if n not in seen:seen.add(n);mapping.append({'output':index,'original':n,'prepared':n});continue
 src=Path(targets[ids[n-1].get('{'+ns['r']+'}id')]).name
 captured=io.StringIO()
 with contextlib.redirect_stdout(captured):mod.duplicate_slide(unpacked,src)
 match=re.search(r'<p:sldId id="(\d+)" r:id="([^"]+)"/>',captured.getvalue())
 if not match:raise ValueError('Unable to parse bundled duplicate-slide result')
 pres=E.parse(str(pres_path));lst=pres.find('p:sldIdLst',ns);E.SubElement(lst,'{'+ns['p']+'}sldId',id=match.group(1),attrib={'{'+ns['r']+'}id':match.group(2)})
 pres.write(str(pres_path),encoding='UTF-8',xml_declaration=True,standalone=True)
 slide['templateSlide']=len(lst);mapping.append({'output':index,'original':n,'prepared':len(lst)})
prepared_template=out/'template.pptx'
with zipfile.ZipFile(prepared_template,'w',zipfile.ZIP_DEFLATED) as z:
 for f in unpacked.rglob('*'):
  if f.is_file():z.write(f,f.relative_to(unpacked).as_posix())
plan['template']=str(prepared_template)
prepared_plan=out/'plan.json';prepared_plan.write_text(json.dumps(plan,ensure_ascii=False,indent=2),encoding='utf-8')
(out/'source-map.json').write_text(json.dumps(mapping,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'plan':str(prepared_plan),'template':str(prepared_template),'slides':len(plan['slides'])},ensure_ascii=False))
