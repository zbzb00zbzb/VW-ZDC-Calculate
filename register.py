from pyDes import *
import base64
import win32api


k = des("DESCRYPT", CBC, "\1\0\1\0\0\1\0\0", pad=None, padmode=PAD_PKCS5)
CVolumeSerialNumber=input('please input your code: ')
d_calcu = k.encrypt(str(CVolumeSerialNumber))
print("Encrypted register code is: " ,d_calcu)
input ("Thankyou for using.")
