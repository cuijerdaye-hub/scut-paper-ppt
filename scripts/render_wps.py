import sys, json
from pathlib import Path
import win32com.client
pptx=Path(sys.argv[1]).resolve(); out=Path(sys.argv[2]).resolve(); out.mkdir(parents=True,exist_ok=True)
if not pptx.is_file(): raise FileNotFoundError(pptx)
app=win32com.client.Dispatch('KWPP.Application')
deck=None
try:
    deck=app.Presentations.Open(str(pptx), True, False, False)
    results=[]
    for i in range(1,deck.Slides.Count+1):
        dest=out/f'slide-{i:02d}.png';deck.Slides.Item(i).Export(str(dest),'PNG',1600,900)
        results.append(str(dest))
    print(json.dumps({'slides':len(results),'previews':str(out)},ensure_ascii=False))
finally:
    if deck is not None: deck.Close()
