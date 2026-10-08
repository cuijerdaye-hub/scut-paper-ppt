from pathlib import Path
from lxml import etree as E
import zipfile,json,posixpath,re,copy
import argparse,hashlib
ap=argparse.ArgumentParser(description='Replace selected pages in a base PPTX with matching page numbers from a designed PPTX. Both decks must have the same total page count.')
ap.add_argument('--base',required=True);ap.add_argument('--designed',required=True);ap.add_argument('--pages',required=True,help='JSON file containing 1-based page numbers');ap.add_argument('--output',required=True);args=ap.parse_args()
source=Path(args.base).resolve();donor=Path(args.designed).resolve();out=Path(args.output).resolve()
if out in (source,donor) or out.exists():raise ValueError('Output must be a new file distinct from the inputs')
selected=json.loads(Path(args.pages).read_text(encoding='utf-8-sig'))
if not selected or len(set(selected))!=len(selected):raise ValueError('Page list must be non-empty and unique')
prefix='layout_'+hashlib.sha256(donor.read_bytes()).hexdigest()[:10]+'_'
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'};rel_ns='http://schemas.openxmlformats.org/package/2006/relationships';ct_ns='http://schemas.openxmlformats.org/package/2006/content-types'
xml=E.fromstring;ser=lambda t:E.tostring(t,encoding='UTF-8',xml_declaration=True,standalone=True)
with zipfile.ZipFile(source) as z:a={n:z.read(n) for n in z.namelist()}
with zipfile.ZipFile(donor) as z:b={n:z.read(n) for n in z.namelist()}
def relname(part):return posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
def owner(rn):return posixpath.join(posixpath.dirname(posixpath.dirname(rn)),posixpath.basename(rn)[:-5])
def resolve(part,t):return posixpath.normpath(posixpath.join(posixpath.dirname(part),t)).lstrip('/')
max_slide=max(int(m.group(1)) for n in a if (m:=re.fullmatch(r'ppt/slides/slide(\d+)\.xml',n)))
map_parts={}
for name in b:
 if not name.startswith('ppt/') or '/_rels/' in name:continue
 m=re.fullmatch(r'ppt/slides/slide(\d+)\.xml',name)
 new=f'ppt/slides/slide{max_slide+int(m.group(1))}.xml' if m else posixpath.join(posixpath.dirname(name),prefix+posixpath.basename(name))
 if new in a:raise ValueError('Part collision: '+new)
 map_parts[name]=new
for name in b:
 if name.startswith('ppt/') and '/_rels/' in name:
  part=owner(name)
  if part in map_parts:map_parts[name]=relname(map_parts[part])
for name,new in map_parts.items():
 data=b[name]
 if name.endswith('.rels'):
  rt=xml(data);part=owner(name)
  for r in rt:
   if r.get('TargetMode')!='External':
    target=resolve(part,r.get('Target'))
    if target not in map_parts:raise ValueError('Unmapped donor relation: '+target)
    r.set('Target','/'+map_parts[target])
  data=ser(rt)
 a[new]=data
pres=xml(a['ppt/presentation.xml']);pr=xml(a['ppt/_rels/presentation.xml.rels']);ids=list(pres.find('p:sldIdLst',ns));dr=xml(b['ppt/_rels/presentation.xml.rels']);dtargets={r.get('Id'):resolve('ppt/presentation.xml',r.get('Target')) for r in dr if r.get('TargetMode')!='External'}
di=list(xml(b['ppt/presentation.xml']).find('p:sldIdLst',ns));ridq='{'+ns['r']+'}id'
if len(ids)!=len(di) or any(not isinstance(i,int) or i<1 or i>len(ids) for i in selected):raise ValueError('Both decks must have matching page counts and valid selected pages')
for i in selected:
 target=dtargets[di[i-1].get(ridq)];newid=f'{prefix}slide{i}'
 if any(r.get('Id')==newid for r in pr):raise ValueError('Relationship collision')
 ids[i-1].set(ridq,newid)
 E.SubElement(pr,'{'+rel_ns+'}Relationship',Id=newid,Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',Target='/'+map_parts[target])
used={e.get(ridq) for e in ids}
for r in list(pr):
 if r.get('Type','').endswith('/slide') and r.get('Id') not in used:pr.remove(r)
# Register donor slide masters, retaining their original layouts and theme graph.
ml=pres.find('p:sldMasterIdLst',ns)
if ml is None:ml=E.SubElement(pres,'{'+ns['p']+'}sldMasterIdLst')
next_id=max([int(x.get('id')) for x in ml] or [2147483648])+1
for j,r in enumerate(dr):
 if r.get('Type','').endswith('/slideMaster'):
  target=resolve('ppt/presentation.xml',r.get('Target'));newid=f'{prefix}master{j}'
  E.SubElement(pr,'{'+rel_ns+'}Relationship',Id=newid,Type=r.get('Type'),Target='/'+map_parts[target])
  E.SubElement(ml,'{'+ns['p']+'}sldMasterId',id=str(next_id),attrib={ridq:newid});next_id+=1
# Copy table style definitions when the source has none; cells also carry explicit formatting.
if not any(r.get('Type','').endswith('/tableStyles') for r in pr):
 for r in dr:
  if r.get('Type','').endswith('/tableStyles'):
   target=resolve('ppt/presentation.xml',r.get('Target'));E.SubElement(pr,'{'+rel_ns+'}Relationship',Id=prefix+'tableStyles',Type=r.get('Type'),Target='/'+map_parts[target])
a['ppt/presentation.xml']=ser(pres);a['ppt/_rels/presentation.xml.rels']=ser(pr)
ct=xml(a['[Content_Types].xml']);existing_defaults={x.get('Extension') for x in ct if E.QName(x).localname=='Default'}
for e in xml(b['[Content_Types].xml']):
 if E.QName(e).localname=='Default':
  if e.get('Extension') not in existing_defaults:ct.append(copy.deepcopy(e));existing_defaults.add(e.get('Extension'))
 else:
  old=e.get('PartName','').lstrip('/')
  if old in map_parts:E.SubElement(ct,'{'+ct_ns+'}Override',PartName='/'+map_parts[old],ContentType=e.get('ContentType'))
# Remove package parts no longer reachable from the final presentation graph.
keep={'_rels/.rels'};pending=[]
for r in xml(a['_rels/.rels']):
 if r.get('TargetMode')!='External':pending.append(posixpath.normpath(r.get('Target')).lstrip('/'))
while pending:
 part=pending.pop()
 if part in keep:continue
 if part not in a:raise ValueError('Missing package target '+part)
 keep.add(part);rn=relname(part)
 if rn in a:
  keep.add(rn)
  for r in xml(a[rn]):
   if r.get('TargetMode')!='External':pending.append(resolve(part,r.get('Target')))
for e in list(ct):
 if E.QName(e).localname=='Override' and e.get('PartName','').lstrip('/') not in keep:ct.remove(e)
a['[Content_Types].xml']=ser(ct);keep.add('[Content_Types].xml')
out.parent.mkdir(parents=True,exist_ok=True)
with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
 for n in sorted(keep):z.writestr(n,a[n])
print(json.dumps({'output':str(out),'pages':len(ids),'replaced':selected,'preserved':[n for n in range(1,len(ids)+1) if n not in selected]},ensure_ascii=False))
