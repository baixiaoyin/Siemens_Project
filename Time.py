import time

localtime1 = time.localtime()
print ("Localtime is:", localtime1)

localtime2 = time.asctime( time.localtime() )
print ("Localtime is:", localtime2)

localtime3 = time.strftime("%Y%m%d", time.localtime())
print ("Localtime is:", localtime3)

localtime4 = time.strftime("%a %b %d %H:%M:%S %Y", time.localtime())
print ("Localtime is:", localtime4)
  

a = "Sat Mar 28 22:24:24 2016"
print (time.mktime(time.strptime(a,"%a %b %d %H:%M:%S %Y")))

str = input("Press any key to Exit");
