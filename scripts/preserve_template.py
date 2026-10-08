import json,zipfile,posixpath,copy,re,sys
from pathlib import Path
from lxml import etree as E
import argparse
ap=argparse.ArgumentParser();ap.add_argument('--plan',required=True);ap.add_argument('--edited',required=True);ap.add_argument('--output',required=True);args=ap.parse_args()
plan_path=Path(args.plan).resolve();out=Path(args.output).resolve();base=out.parent
plan=json.loads(plan_path.read_text(encoding='utf-8-sig'))
plan['template']=str((plan_path.parent/plan['template']).resolve())
if out==Path(plan['template']):raise ValueError('Output cannot overwrite source template')
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
rel_ns='http://schemas.openxmlformats.org/package/2006/relationships'
qn=lambda p,n:'{'+ns[p]+'}'+n
xml=lambda b:E.fromstring(b)
ser=lambda e:E.tostring(e,encoding='UTF-8',xml_declaration=True,standalone=True)
with zipfile.ZipFile(plan['template']) as z:parts={n:z.read(n) for n in z.namelist()}
with zipfile.ZipFile(args.edited) as z:authored={n:z.read(n) for n in z.namelist()}
def relname(part):return posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
def resolve(part,target):return posixpath.normpath(posixpath.join(posixpath.dirname(part),target)).lstrip('/')
def note_target(d,part):
 r=xml(d[relname(part)])
 for x in r:
  if x.get('Type','').endswith('/notesSlide'):return resolve(part,x.get('Target'))
 return None

def geometry(shape,desired):
 xf=shape.find('p:spPr/a:xfrm',ns)
 if xf is None:raise ValueError('Text slot has no transform')
 ax=ay=1.;bx=by=0.
 ancestors=[g for g in shape.iterancestors() if g.tag==qn('p','grpSp')][::-1]
 for g in ancestors:
  tr=g.find('p:grpSpPr/a:xfrm',ns)
  off,ext,co,ce=[tr.find('a:'+x,ns) for x in ['off','ext','chOff','chExt']]
  sx=float(ext.get('cx'))/float(ce.get('cx'));sy=float(ext.get('cy'))/float(ce.get('cy'))
  bx+=ax*(float(off.get('x'))-float(co.get('x'))*sx);by+=ay*(float(off.get('y'))-float(co.get('y'))*sy)
  ax*=sx;ay*=sy
 off=xf.find('a:off',ns);ext=xf.find('a:ext',ns)
 if 'left' in desired:off.set('x',str(round((desired['left']*9525-bx)/ax)))
 if 'top' in desired:off.set('y',str(round((desired['top']*9525-by)/ay)))
 if 'width' in desired:ext.set('cx',str(round(desired['width']*9525/ax)))
 if 'height' in desired:ext.set('cy',str(round(desired['height']*9525/ay)))
report=[]
for page in plan['slides']:
 n=page['templateSlide'];part=f'ppt/slides/slide{n}.xml';src=xml(parts[part]);edited=xml(authored[part]);
 originals={s.find('p:nvSpPr/p:cNvPr',ns).get('name'):s for s in src.findall('.//p:sp',ns)}
 updated={s.find('p:nvSpPr/p:cNvPr',ns).get('name'):s for s in edited.findall('.//p:sp',ns)}
 for edit in page['edits']:
  name=edit['name'];target=originals[name];replacement=updated[name].find('p:txBody',ns)
  if replacement is None:raise ValueError('Missing authored text '+name)
  target.replace(target.find('p:txBody',ns),copy.deepcopy(replacement))
  desired=dict(edit.get('position',{}))
  if 'height' in edit:desired['height']=edit['height']
  if desired:geometry(target,desired)
  written='\n'.join(''.join(p.itertext()) for p in target.findall('p:txBody/a:p',ns))
  actual=''.join(target.xpath('./p:txBody//a:t/text()',namespaces=ns))
  if actual.replace('\n','')!=edit['text'].replace('\n',''):raise ValueError('Text transfer mismatch '+name)
 # Preserve source images and all their transforms/crops verbatim.
 pictures_before=[E.tostring(x) for x in xml(parts[part]).findall('.//p:pic',ns)]
 pictures_after=[E.tostring(x) for x in src.findall('.//p:pic',ns)]
 if pictures_before!=pictures_after:raise ValueError('Template pictures changed')
 parts[part]=ser(src)
 oldnote,newnote=note_target(parts,part),note_target(authored,part)
 if not newnote:raise ValueError('Missing authored note relationship')
 if not oldnote:
  oldnote=f'ppt/notesSlides/scutNotes{n}.xml'
  relroot=xml(parts[relname(part)])
  E.SubElement(relroot,'{'+rel_ns+'}Relationship',Id=f'scutNotes{n}',Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',Target=f'../notesSlides/scutNotes{n}.xml')
  parts[relname(part)]=ser(relroot)
  nr=E.Element('{'+rel_ns+'}Relationships',nsmap={None:rel_ns})
  E.SubElement(nr,'{'+rel_ns+'}Relationship',Id='rId1',Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',Target=f'../slides/slide{n}.xml')
  E.SubElement(nr,'{'+rel_ns+'}Relationship',Id='rId2',Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',Target='../notesMasters/notesMaster1.xml')
  parts[relname(oldnote)]=ser(nr)
  ct0=xml(parts['[Content_Types].xml'])
  E.SubElement(ct0,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName='/'+oldnote,ContentType='application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml')
  parts['[Content_Types].xml']=ser(ct0)
 parts[oldnote]=authored[newnote]
 report.append({'source_slide':n,'edited_text_slots':len(page['edits']),'preserved_picture_objects':len(pictures_before)})
# Reorder existing slide references; source slide parts retain their native relationships.
pres=xml(parts['ppt/presentation.xml']);lst=pres.find('p:sldIdLst',ns);old=list(lst)
selected=[old[d['templateSlide']-1] for d in plan['slides']]
for x in list(lst):lst.remove(x)
for x in selected:lst.append(x)
parts['ppt/presentation.xml']=ser(pres)
pr=xml(parts['ppt/_rels/presentation.xml.rels']);selected_ids={x.get(qn('r','id')) for x in selected}
for x in list(pr):
 if x.get('Type','').endswith('/slide') and x.get('Id') not in selected_ids:pr.remove(x)
parts['ppt/_rels/presentation.xml.rels']=ser(pr)
rootrels=xml(parts['_rels/.rels'])
for x in list(rootrels):
 if x.get('Type','').endswith('/thumbnail'):rootrels.remove(x)
parts['_rels/.rels']=ser(rootrels)
if 'docProps/core.xml' in parts:
 core=xml(parts['docProps/core.xml']);dc='http://purl.org/dc/elements/1.1/'
 title=core.find('{'+dc+'}title')
 if title is None:title=E.SubElement(core,'{'+dc+'}title')
 title.text=plan['title'];parts['docProps/core.xml']=ser(core)
if 'docProps/app.xml' in parts:
 app=xml(parts['docProps/app.xml'])
 for x in list(app):
  local=E.QName(x).localname
  if local=='Slides':x.text=str(len(plan['slides']))
  if local in ('HeadingPairs','TitlesOfParts'):app.remove(x)
 parts['docProps/app.xml']=ser(app)
# Keep exactly the package graph reachable from root relationships.
keep={'_rels/.rels'};pending=[]
for r in rootrels:
 if r.get('TargetMode')!='External':pending.append(posixpath.normpath(r.get('Target')).lstrip('/'))
while pending:
 part=pending.pop()
 if part in keep:continue
 if part not in parts:raise ValueError('Missing relationship target '+part)
 keep.add(part);rn=relname(part)
 if rn in parts:
  keep.add(rn)
  for r in xml(parts[rn]):
   if r.get('TargetMode')!='External':pending.append(resolve(part,r.get('Target')))
ct=xml(parts['[Content_Types].xml'])
for e in list(ct):
 if E.QName(e).localname=='Override' and e.get('PartName','').lstrip('/') not in keep:ct.remove(e)
parts['[Content_Types].xml']=ser(ct);keep.add('[Content_Types].xml')
out.parent.mkdir(parents=True,exist_ok=True)
with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
 for n in sorted(keep):z.writestr(n,parts[n])
out.with_suffix('.preservation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'output':str(out),'slides':len(selected),'parts':len(keep),'preservation':report},ensure_ascii=False))
