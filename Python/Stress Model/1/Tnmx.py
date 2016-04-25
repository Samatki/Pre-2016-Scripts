import math
import random
import csv


print('How many results to generate? Up to 1 million')

# Tensile Strength South (Min)
def Tnmx (x,y,z,d):
	T= -8.695774258*(10**-5) *z**3 + 2.307202491*(10**-3) *z**2 *w + 7.566908253*(10**-3) *z**2 *x + 8.931672636*(10**-3) *z**2 *y - 3.016752583*(10**-2) *z *w**2 + 1.983718049*(10**-1) *z *w *x+ 1.054002433*(10**-1) *z *w *y - 2.404538246*(10**-1) *z *x**2 - 5.086026589*(10**-1) *z *x *y - 2.101809308*(10**-1) *z *y**2 - 0.796828685 *w**3 + 3.251038523 *(w**2) *x + 7.223539691 *w**2 *y + 1.978303249 *w *x**2- 6.856465712 *w *x *y - 6.53451066 *w *y**2 + 4.036033321*10**-1 *x**3 + 1.437798708 *x**2 *y + 5.400582882 *x *y**2 - 2.698617708*(10**-2) *y**3 - 2.177228046*(10**-1) *z**2 - 4.015103448 *z *w + 6.97487421 *z *x + 7.460240712 *z *y - 76.54981976 *w**2 - 146.0363381 *w *x - 61.75027102 *w *y - 23.17849948 *x**2 + 64.37888639 *x *y + 109.3382879 *y**2 - 16.99902755 *z + 3409.645453 *w+ 523.203409 *x - 2169.54498 *y - 12398.47684
	#print(T)
	return(T)

def lnmean (a,b):
	mu = math.log(a/((1+(b**2)/(a**2))**0.5))
	return(mu)
	
def lnstdev (a,b):	
	sigma = (math.log(1+(b**2)/(a**2)))**0.5
	return(sigma)
	
numCases = int(input())

with open('J:/DATA/J80193/A/03 Base Information/Stella Reliability Study/Working/Stress Modelling Calc//Tnmx.csv', 'w', newline='') as csvfile:
	Twriter = csv.writer(csvfile)
	
	
	for i in range (0,numCases,1):

	### 	All input data are log normal Distributed     ###

	# w = Hmx = Wave Height(m), mean 20.91, STDEV 4.27
		a = 20.91
		b = 4.27
		w = random.lognormvariate(lnmean(a,b),lnstdev(a,b))
	# x = Offs = Platform Offet (m), mean 13.00, STDEV 2.60
		a = 13.00
		b = 2.60
		x = random.lognormvariate(lnmean(a,b),lnstdev(a,b))
	# y = Tass = Wave Period (sec), mean 13.49, STDEV 2.24
		a = 13.49
		b = 2.24
		y = random.lognormvariate(lnmean(a,b),lnstdev(a,b))
	# z = Dwc = Wave Current Direction (deg), mean 186.56, STDEV 93.58, Range: +- 1 STDEV around mean
		z = random.uniform(92.98,280.14)

	# d = D = common link diameter - Uniform Distribution between negative tolerance of -2mm and positive tolerance of 5%, Mean value of 0.076m, (m)
		d = random.uniform(0.074,1.05*0.076)
		
		T = Tnmx(x,y,z,d)
		Twriter.writerow([T])
		
		i = i+1

