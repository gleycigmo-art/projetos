import sys,base64,struct,zipfile,io,re
raw=open(sys.argv[1],'rb').read().decode('utf-16')
b64=re.search(r'>([A-Za-z0-9+/=\s]+)</DataMashup>',raw).group(1)
data=base64.b64decode(b64)
ver,plen=struct.unpack('<II',data[:8])
z=zipfile.ZipFile(io.BytesIO(data[8:8+plen]))
for n in z.namelist():
    print('##',n)
    if n.endswith('.m'):
        open(sys.argv[2]+'/Section1.m','wb').write(z.read(n))
