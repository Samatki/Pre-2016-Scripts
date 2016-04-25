
import random
import csv


print('How many results to generate? Up to 1 million')

# Tensile Strength South (Min)
def Tnmn (x,y,z,d):
	T= 4.593000017*(10**-5) *z**3 - 2.972871202*(10**-3) *z**2 *w - 6.771783206*(10**-3) *z**2 *x - 7.685446512*(10**-3) *z**2 *y + 2.681558456*(10**-2) *z *w**2 - 0.312422328 *z *w *x - 1.819140574*(10**-1) *z *w *y + 1.927507734*(10**-1) *z *x**2 + 4.284409994*(10**-1) *z *x *y + 2.213791898*(10**-1) *z *y**2 + 1.246672801 *w**3 - 3.158125862 *w**2 *x - 10.95083337 *w**2 *y - 3.261344559 *w *x**2 + 3.086376604 *w *x *y+ 17.12965775 *w *y**2 - 4.991828877*(10**-1) *x**3 - 4.71012395*(10**-2) *x**2 *y - 1.999811894 *x *y**2 - 10.55311748 *y**3 + 0.220967161 *z**2 + 7.041791539 *z *w - 2.27702082 *z *x - 5.348021328 *z *y+ 97.28111762 *w**2 + 249.7714125 *w *x - 11.80092946 *w *y + 43.60449264 *x**2 - 97.02224476 *x *y + 67.72597727 *y**2 - 62.63062115 *z - 4292.20839 *w - 2154.14126 *x + 421.2367518 *y+ 40527.47197
	#print(T)
	return(T)

numCases = int(input())

with open('Tnmn.csv', 'w', newline='') as csvfile:
	Twriter = csv.writer(csvfile)
	


	for i in range (0,numCases,1):

	### 	All input data are Gaussian Distributed ###

	# w = Hmx = Wave Height(m), mean 20.91, STDEV 4.27
		w = random.lognormvariate(20.91,4.27)
	# x = Offs = Platform Offet (m), mean 13.00, STDEV 2.60
		x = random.lognormvariate(13.00,2.60)
	# y = Tass = Wave Period (sec), mean 13.49, STDEV 2.24
		y = random.lognormvariate(13.49,2.24)
	# z = Dwc = Wave Current Direction (deg), mean 186.56, STDEV 93.58
		z = random.uniform(92.98,280.14)

	# d = D = common link diameter - Uniform Distribution between negative tolerance of -2mm and positive tolerance of 5%, Mean value of 0.076m, (m)
		d = random.uniform(0.074,1.05*0.076)
		
		T = Tnmn(x,y,z,d)
		Twriter.writerow([T])
		
		i = i+1

