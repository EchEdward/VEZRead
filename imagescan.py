from PIL import Image
import numpy as np
import colorsys
#from asprise_ocr_api import *
from pytesseract import image_to_string
import os.path
 
def check_program(progname):
    path = os.getenv('PATH')
    for p in path.split(';'):
        if os.path.exists(os.path.join(p, progname)):
            return True
    return False


tess_state = check_program("tesseract.exe")

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
            tr = False
            for k,m in zip(range(len(rows)-1,-1,-1),range(1,len(rows)+1)):
                if i == rows[k]+m:
                    tr = True
                    break
            if not tr:
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
            tr = False
            for k,m in zip(range(len(cols)-1,-1,-1),range(1,len(cols)+1)):
                if i == cols[k]+m:
                    tr = True
                    break
            if not tr:
                cols.append(i)


    if py != rows[0]:
        rows=[py]+rows
    if rows[len(rows)-1] != height:
        rows.append(height)
    if px != cols[0]:
        cols=[px]+cols
    if cols[len(cols)-1] != width:
        cols.append(width)

    """ print(Name)
    print("rows: ",rows)
    print("cols: ",cols)
    print("-"*10) """

    sp_image = []
    #h_standart = 16
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
    #print(Name)
    custom_config = r'--oem 3' # --psm 6
    #print("-"*13)
    sp_f = []
    st_dgt = []
    for i in sp_image:
        Im = Image.open("interim_image/%r_%r.png" % i)
        Im = Im.convert('L')
        Im = Im.point(lambda x: 0 if x<128 else 255, '1')
        dt = np.asarray(Im)
        d_n = {}
        for numb in range(len(N)):
            if numb != 10:
                a,b = 11,6
            else:
                a,b = 11,2
            for m in range(np.shape(dt)[0]-a):
                for n in range(np.shape(dt)[1]-b):
                    tr = dt[m:m+a,n:n+b] == N[numb]
                    if np.sum(tr) >=a*b*0.97:
                        d_n[n]=numb
                        
        #print(d_n)
        sr = sorted(list(d_n.keys()))
        s=""
        for l in sr:
            s+=sp_n[d_n[l]]
        #print("result: ",s)

        try: 
            dgt1 = float(s)
        except Exception:
            dgt1=-1

        if dgt1 == -1:
            try: 
                if not tess_state: raise Exception
                tes = image_to_string(Im, config = custom_config).strip() #
                #print("tes: ", tes)
                dgt2 = float(tes)
            except Exception:
                dgt2=-1
        else:
            dgt2 = -1

                
        if dgt1 !=-1 and dgt2 !=-1:
            sp_f.append(dgt1)
            st_dgt.append(3)
            #print("two: ",dgt1)
        elif dgt1 !=-1 and dgt2 ==-1:
            sp_f.append(dgt1)
            st_dgt.append(3)
            #print("two: ",dgt1)
        elif dgt2 != -1 and dgt1 ==-1:
            sp_f.append(dgt2)
            st_dgt.append(2)
            #print("one: ",dgt2)
        else:
            sp_f.append("")
            st_dgt.append(1)

    t = [["","",""] for i in range(sp_image[len(sp_image)-1][0])]
    state = [[[0,False],[0,False],[0,False]] for i in range(sp_image[len(sp_image)-1][0])]
    for i in range(len(sp_image)):
        t[sp_image[i][0]-1][sp_image[i][1]-1] = sp_f[i]
        state[sp_image[i][0]-1][sp_image[i][1]-1] = [st_dgt[i],False]

    #print(t)
    #print(state)
    return t, state

def common_state(lst):
    rez = 3
    for i in lst:
        for j in i:
            if j != []:
                if j[0] < rez and j[0] > 0:
                    rez = j[0]
    return rez

def define_state(arr):
    
    state2 = []
    a = arr[-2:]
    a_len = len(a)
    if a_len >=2:
        if  (a[1][0] != '' or a[1][0] == '') and a[1][1] == '' and a[1][2] == '':
            state2.append([3 if a[0][0] != '' else 1,3 if a[0][1] != '' else 1,3 if a[0][2] != '' else 1])
            state2.append([3 if a[1][0] != '' else 1,0,0])

        elif a[1][0] != '' and (a[1][1] != '' or a[1][2] != ''):
            state2.append([3 if a[0][0] != '' else 1,3 if a[0][1] != '' else 1,3 if a[0][2] != '' else 1])
            state2.append([3 if a[1][0] != '' else 1,3 if a[1][1] != '' else 1,3 if a[1][2] != '' else 1])
            state2.append([1,0,0])
    else:
        state2.append([3 if a[0][0] != '' else 1,3 if a[0][1] != '' else 1,3 if a[0][2] != '' else 1])
        state2.append([1,0,0])

    state1 = []
    if len(arr) >= 3:
        for i in range(len(arr)-2):
            state1.append([3 if arr[0][0] != '' else 1,3 if arr[0][1] != '' else 1,3 if arr[0][2] != '' else 1])

    state = state1 + state2
    
    
    return [[[i,False],[j,False],[k,False]] for i,j,k in state]

def new_state(old,new,i,j):
    old_len = len(old)
    new_len = len(new)
    common = []
    #print(old)
    #print("+"*10)
    #print(new)
    #print("-"*10)
    for m in range(new_len): #max(old_len,new_len)
        c = [[0,False],[0,False],[0,False]]
        for n in range(3):
            if m > old_len-1 and m <= new_len-1:
                if i==m and j==n:
                    c[n][0] = new[m][n][0]
                    c[n][1] = True
                else:
                    c[n][0] = new[m][n][0]
                    c[n][1] = True

            elif m <= old_len-1 and m > new_len-1:
                if i==m and j==n:
                    pass
                else:
                    if old[m][n][0]==1:
                        c[n][0] = 1
                        c[n][1] = old[m][n][1]
            
            elif m <= old_len-1 and m <= new_len-1:
                if i==m and j==n:
                    if new[m][n][0]==3 and old[m][n][0]==2:
                        c[n][0] = 3
                    elif new[m][n][0]==3 and old[m][n][0]==1:
                        c[n][0] = 3
                    else:
                        c[n][0] = new[m][n][0]
                    c[n][1] = True
                else:
                    if new[m][n][0]==3 and old[m][n][0]==2 and old[m][n][1]:
                        c[n][0] = 3
                    elif new[m][n][0]==3 and old[m][n][0]==2 and not old[m][n][1]:
                        c[n][0] = 2
                    elif new[m][n][0]==0 and old[m][n][0]==1 and old[m][n][1]:
                        c[n][0] = 0
                    elif new[m][n][0]==0 and old[m][n][0]==1 and not old[m][n][1]:
                        c[n][0] = 1
                    elif new[m][n][0]==3 and old[m][n][0]==1 and not old[m][n][1]:
                        c[n][0] = 1
                    elif new[m][n][0]==3 and old[m][n][0]==1 and old[m][n][1]:
                        c[n][0] = 3
                    else:
                        c[n][0] = new[m][n][0]
                    c[n][1] = old[m][n][1]
                

            """ elif m == old_len and m == new_len:
                4 """
        if not (c[0][0] == 0 and c[1][0] == 0 and c[2][0] == 0):
            common.append(c)
        
    #print(common)

    return common


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