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
sensor_nodes={}
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
    sensor_nodes[i]=kt[i]

print(sensor_nodes)
df=pd.DataFrame(sensor_nodes)
