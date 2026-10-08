"""Apply the selected SCUT table border style to an exported native PPTX table."""
import sys, zipfile, xml.etree.ElementTree as ET, os
from pathlib import Path
p=Path(sys.argv[1]).resolve()
ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
for k,u in {'a':ns['a'],'p':'http://schemas.openxmlformats.org/presentationml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}.items(): ET.register_namespace(k,u)
q=lambda n:'{'+ns['a']+'}'+n
parts=[];count=0
with zipfile.ZipFile(p) as z:
 for info in z.infolist():
  data=z.read(info.filename)
  if info.filename.startswith('ppt/slides/slide') and info.filename.endswith('.xml'):
   root=ET.fromstring(data);changed=False
   for table in root.findall('.//a:tbl',ns):
    for props in table.findall('.//a:tcPr',ns):
     for edge in ('lnL','lnR','lnT','lnB'):
      line=props.find(q(edge))
      if line is None: line=ET.SubElement(props,q(edge))
      line.set('w','6350')
      for c in list(line): line.remove(c)
      fill=ET.SubElement(line,q('solidFill'));ET.SubElement(fill,q('srgbClr'),{'val':'D4DFEC'})
      ET.SubElement(line,q('prstDash'),{'val':'solid'});changed=True;count+=1
   if changed:data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
  parts.append((info,data))
tmp=p.with_suffix('.styled.pptx')
with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as z:
 for info,data in parts:z.writestr(info,data)
os.replace(tmp,p)
print('Styled native table borders:',count)
