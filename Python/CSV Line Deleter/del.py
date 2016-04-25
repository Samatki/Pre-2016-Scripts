import glob, os
os.chdir("MainDir")
for file in glob.glob("*.ppl"):
	with open(file,"r") as textobj:
		alist = list(textobj)
	
	for i in range(1,28480,1):
		del alist[257-1]
	
	with open(file,"w") as textobj:
		for n in alist:
			textobj.write(n)

