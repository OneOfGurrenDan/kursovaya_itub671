python3 sheet.py >/dev/null && python3 - <<'P'
from PIL import Image
import glob,numpy as np
fs=sorted(glob.glob('/home/claude/v2/report/build/pg/p-*.png'))
for i,f in enumerate(fs,1):
    a=np.array(Image.open(f).convert('L'));h=a.shape[0]
    rows=np.where((a<200).sum(axis=1)>0)[0];body=rows[rows<h*0.93]
    print(i, round(body.max()/h*100), end=' | ')
P
pdftotext -layout "/home/claude/v2/report/build/Отчёт_курсовая_GM_GYM_v2.pdf" - | grep "\.\.\.\.\.\." | sed 's/\.\.\.*/ /' | tr '\n' ';'
