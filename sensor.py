import docx2txt
import pandas as pd
tf=docx2txt.process("C:\\Users\\xxx\\Desktop\\cor.docx")
tf=tf.split()
tf=tf[4:]
c=tf.count('SENSOR')
for i in range(c):
    tf.remove('SENSOR')
ind=tf.index('Y')
tf1=tf[:ind+1]
tf2=tf[ind+1:]
tf2=tf2[3:]
tf1+=tf2[:2]
tf2=tf2[5:]
tf1+=tf2[:2]
tf2=tf2[4:]
tf1+=tf2[:2]
tf2=tf2[4:]
tf1+=tf2[:2]
tf2=tf2[4:]
tf1=tf1+tf2

d={}
for i in range(0,len(tf1),2):
    d[tf1[i+1]]=tf1[i]
p=sorted(d)

kt={}
scord={}
for i in p:
    k=d[i]
    if(len(i)==1):
        nu=ord(i)-64
    else:
        if(i[0]=='A'):
            nu=26
            nu+=(ord(i[1])-64)
        else:
            nu=52
            nu+=(ord(i[1])-64)
    k=k.split(',')
    k[0]=int(k[0])
    k[1]=int(k[1])
    kt[nu]=k
for i in sorted(kt):
    scord[i]=kt[i]


    

    
    
    
    
    
    



import pandas as pd
df=pd.read_excel("C:\\Users\\bellapukonda\\Downloads\\data.xlsx")
df=df[1101:]
df=df.drop(["Unnamed: 7","Unnamed: 8","Unnamed: 9","Unnamed: 10","Unnamed: 11","Unnamed: 12","Unnamed: 13","Unnamed: 14","Unnamed: 15"],axis=1)
df=df[0:12448]
df.index=range(0,len(df))
df.columns = df.iloc[0]
df=df.drop(0,axis=0)
df.columns=['sid','acount','pcount','aid','x','y','t']

"""
ar=df['acount']==df['pcount']
arr=[]
for i in range(1,len(ar)+1):
    arr.append(ar[i])
"""
df=df.drop(["acount"],axis=1)
df.index=range(0,12447)
# no of rows =12448
#time starts from 0 to 4900
d={}
aid65c={}
aid66c={}
aid67c={}
aid68c={}

aid65={}
aid66={}
aid67={}
aid68={}

for i in range(0,len(df)):
    if(df['aid'][i]==65):
        aid65c[df['t'][i]]=[df['x'][i],df['y'][i]]
    if(df['aid'][i]==66):
        aid66c[df['t'][i]]=[df['x'][i],df['y'][i]]
    if(df['aid'][i]==67):
        aid67c[df['t'][i]]=[df['x'][i],df['y'][i]]
    if(df['aid'][i]==68):
        aid68c[df['t'][i]]=[df['x'][i],df['y'][i]]


arr=[]
a=df['pcount']
for i in range(len(a)):
    arr.append(a[i])

k=arr[2]
p=arr.count(k)
for i in range(p):
    arr.remove(k)
packet_count=arr
TOTAL_PACKETS=sum(packet_count)

si=[]
dd=df["sid"]
for i in range(len(dd)):
    si.append(dd[i])
import math
for i in range(len(si)):
    if(math.isnan(si[i])):
        si[i]=0

p=si[0]        
for i in range(1,len(si)):
    if(si[i]==0):
        si[i]=p
    else:
        p=si[i]

df["sid"]=si


for i in range(len(df)):
    
    if(df["aid"][i]==65):
        if(df["t"][i] in aid65):
            aid65[df["t"][i]].append(df["sid"][i])
        else:
            aid65[df["t"][i]]=[df["sid"][i]]
    if(df["aid"][i]==66):
        if(df["t"][i] in aid66):
            aid66[df["t"][i]].append(df["sid"][i])
        else:
            aid66[df["t"][i]]=[df["sid"][i]]
    if(df["aid"][i]==67):
        if(df["t"][i] in aid67):
            aid67[df["t"][i]].append(df["sid"][i])
        else:
            aid67[df["t"][i]]=[df["sid"][i]]
    if(df["aid"][i]==68):
        if(df["t"][i] in aid68):
            aid68[df["t"][i]].append(df["sid"][i])
        else:
            aid68[df["t"][i]]=[df["sid"][i]]
    



def dis(a,b):
    s=(a[0]-b[0])**2+(a[1]-b[1])**2
    s=s**(1/2)
    return s
    
    
times=aid65.keys()
min65={}
min66={}
min67={}
min68={}
for i in times:
    min65[i]=[]
    min66[i]=[]
    min67[i]=[]
    min68[i]=[]



for i in times:
    for j in aid65[i]:
        d=dis(aid65c[i],scord[j])
        min65[i].append((d,j))
    for j in aid66[i]:
        d=dis(aid66c[i],scord[j])
        min66[i].append((d,j))
    for j in aid67[i]:
        d=dis(aid67c[i],scord[j])
        min67[i].append((d,j))
    for j in aid68[i]:
        d=dis(aid68c[i],scord[j])
        min68[i].append((d,j))

for i in times:
    min65[i]=min(min65[i])
    min66[i]=min(min66[i])
    min67[i]=min(min67[i])
    min68[i]=min(min68[i])


    
    
# min65 contains the sensor node nearer to agent node 65 at each time  (0,10000000,.....ms)
    

    

final_min_to_65=min(min65.values())
final_min_to_66=min(min66.values())
final_min_to_67=min(min67.values())
final_min_to_68=min(min68.values())




print("SENSOR NODE "+str(final_min_to_65[1])+" is nearer to agent node65 by "+str(final_min_to_65[0]))
print("SENSOR NODE "+str(final_min_to_66[1])+" is nearer to agent node66 by "+str(final_min_to_66[0]))
print("SENSOR NODE "+str(final_min_to_67[1])+" is nearer to agent node67 by "+str(final_min_to_67[0]))
print("SENSOR NODE "+str(final_min_to_68[1])+" is nearer to agent node68 by "+str(final_min_to_68[0]))

