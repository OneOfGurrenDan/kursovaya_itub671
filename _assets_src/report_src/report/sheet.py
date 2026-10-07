import subprocess,glob,os,sys
from PIL import Image
b='/home/claude/v2/report/build'
subprocess.run('rm -rf %s/pg; mkdir %s/pg; pdftoppm -r 45 -png "%s/Отчёт_курсовая_GM_GYM_v2.pdf" %s/pg/p'%(b,b,b,b),shell=True,check=True)
fs=sorted(glob.glob(b+'/pg/p-*.png'))
ims=[Image.open(f) for f in fs]
w,h=ims[0].size
cols=9
rows=(len(ims)+cols-1)//cols
sheet=Image.new('RGB',(cols*w,rows*h),'white')
for i,im in enumerate(ims): sheet.paste(im,((i%cols)*w,(i//cols)*h))
sheet.save(b+'/sheet.png');print(len(ims),sheet.size)
