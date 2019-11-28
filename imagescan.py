from PIL import Image
import numpy as np
import colorsys
#from asprise_ocr_api import *

sp_n = ['0','1','2','3','4','5','6','7','8','9','.']
N = []

for i in range(len(sp_n)):
    n_im = Image.open("Numbers/"+sp_n[i]+".png")
    n_im = n_im.convert('L')
    n_im = n_im.point(lambda x: 0 if x<128 else 255, '1')
    dt = np.asarray(n_im)
    N.append(dt)


""" ld = im.load()
width, height = im.size
for y in range(height):
    for x in range(width):
        r,g,b = ld[x,y]
        h,s,v = colorsys.rgb_to_hsv(r/255., g/255., b/255.)

        if s>0.5:                     
           ld[x,y] = (0,0,0)
        else:
           ld[x,y] = (255,255,255) """


""" pixels = list(im.getdata())

pixels = [pixels[i * width:(i + 1) * width] for i in range(height)] """

def Scan(Name):
    im = Image.open(Name)
    im = im.convert('L')
    im = im.point(lambda x: 0 if x<128 else 255, '1')
    data = np.asarray(im)
    width, height = im.size
    for i in range(width-1,-1,-1):
        st = data[:,i]
        black = 0
        for j in range(height):
            if st[j]==False:
                black+=1
        if black/height>0.95:
            px=i
            break

    for i in range(height):
        st=data[i,px+1:]
        black = 0
        for j in range(0,width-(px+1)):
            if st[j]==False:
                black+=1
        if black/(width-(px+1))>0.95:
            py=i
            break
    #print(px,py)

    for i in range(px,-1,-1):
        if not data[py,i] == False:
            px2=i
            break

    rows =[]
    for i in range(py+1,height):
        st=data[i,px2:px]
        black = 0
        for j in range(0,px-px2):
            if st[j]==False:
                black+=1
        if black/(px-px2)>0.90:
            rows.append(i)

    #im.show()
    cols=[]
    for i in range(px+1,width):
        st=data[:py,i]
        black = 0
        for j in range(0,py):
            if st[j]==False:
                black+=1
        if black/py>0.80:
            cols.append(i)


    if py != rows[0]:
        rows=[py]+rows
    if rows[len(rows)-1] != height:
        rows.append(height)
    if px != cols[0]:
        cols=[px]+cols
    if cols[len(cols)-1] != width:
        cols.append(width)

    #print(rows)
    #print(cols)

    sp_image = []
    for i in range(1,len(rows)):
        for j in range(1,len(cols)):
            st=data[rows[i-1]+2:rows[i],cols[j-1]+2:cols[j]]
            if np.sum(st) == np.shape(st)[0]*np.shape(st)[1]:
                break
            im1=im.crop((cols[j-1]+2,rows[i-1]+2,cols[j]-1,rows[i]-1))
            im1.save("interim_image/%r_%r.png" % (i,j))
            sp_image.append((i,j))
            #break
        #break

    #im.save("tt.png")


    sp_f = []
    for i in sp_image:
        Im = Image.open("interim_image/%r_%r.png" % i)
        Im = Im.convert('L')
        Im = Im.point(lambda x: 0 if x<128 else 255, '1')
        dt = np.asarray(Im)
        d_n = {}
        for numb in range(len(N)):
            if numb != 10:
                a,b = 10,6
            else:
                a,b = 10,2
            for m in range(np.shape(dt)[0]-a):
                for n in range(np.shape(dt)[1]-b):
                    tr = dt[m:m+a,n:n+b] == N[numb]
                    if np.sum(tr) ==a*b:
                        #print(numb)
                        d_n[n]=numb
        #print(d_n)
        sr = sorted(list(d_n.keys()))
        s=""
        for l in sr:
            s+=sp_n[d_n[l]]
        #print(s)
        sp_f.append(float(s))

    t = [["","",""] for i in range(sp_image[len(sp_image)-1][0])]
    for i in range(len(sp_image)):
        t[sp_image[i][0]-1][sp_image[i][1]-1] = sp_f[i]

    #print(t)
    return t

#Scan("слои.jpg")

""" Ocr.set_up() # one time setup
ocrEngine = Ocr()
ocrEngine.start_engine("eng")

s = ocrEngine.recognize("%r_%r.png" % i, -1, -1, -1, -1, -1,
                        OCR_RECOGNIZE_TYPE_TEXT, OCR_OUTPUT_FORMAT_PLAINTEXT,PROP_IMG_PREPROCESS_TYPE="custom",
                            PROP_IMG_PREPROCESS_CUSTOM_CMDS="scale(%r);default()" % round(j/100,2))

#print ("Result: " + s)
# recognizes more images here ..
ocrEngine.stop_engine() """

#im.show()