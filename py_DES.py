from pyDes import *
import base64
import win32api

CVolumeSerialNumber = win32api.GetVolumeInformation("C:\\")[1]
print("请将该字符串发给管理员获取激活码：",CVolumeSerialNumber)
k = des("DESCRYPT", CBC, "\1\0\1\0\0\1\0\0", pad=None, padmode=PAD_PKCS5)

CVolumeSerialNumber="-462462504"


d_calcu = k.encrypt(str(CVolumeSerialNumber))
print("Encrypted: " ,d_calcu)
d_input = input('please input your register code: ')
#print(d_input)
if str(d_calcu)==str(d_input):
    print("Ture")
#print("Decrypted: ", k.decrypt(d))
'''
data="20103"
sec1=base64.b64encode(data)
print(sec1)
sec2=base64.b64decode(sec1)
print(sec2)
'''


