import math
import random
import csv


print('How many results to generate? Up to 1 million')

# Tensile Strength South (Min)
def Tsmn (x,y,z,d):
	T= 5.308162608*10**(-5)*z**3-1.687050876*(10**-3)*(z**2)*w-6.056602129*(10**-3)*(z**2)*x-6.971130984*(10**-3)*(z**2)*y+1.055460727*(10**-1)*z*(w**2)-4.184562742*(10**-1)*z*w*x-2.857464996*(10**-1)*z*w*y+3.621574824*(10**-1)*z*(x**2)+7.041553535*(10**-1)*z*x*y+2.820898881*(10**-1)*z*(y**2)+1.623212837*(w**3)-6.459205796*(w**2)*x-17.9686895*(w**2)*y-3.028391521*w*x**2+11.05889484*w*x*y+29.57078225*w*(y**2)+1.019472753*(x**3)-0.992914457*(x**2)*y-12.58091118*x*(y**2)-22.06064806*(y**3)+1.770248315*(10**-1)*(z**2)+5.516673416*z*w-8.41977397*z*x-8.605772328*z*y+187.5028464*(w**2)+302.4066488*w*x-135.4465869*w*y-44.59651028*x**2-9.392994963*x*y+407.0368648*(y**2)+26.94638321*z-5380.148824*w-1519.911803*x-3168.516881*y+53924.02039
	#print(T)
	return(T)

def lnmean (a,b):
	mu = math.log(a/((1+(b**2)/(a**2))**0.5))
	return(mu)
	
def lnstdev (a,b):	
	sigma = (math.log(1+(b**2)/(a**2)))**0.5
	return(sigma)
	
numCases = int(input())

with open('J:/DATA/J80193/A/03 Base Information/Stella Reliability Study/Working/Stress Modelling Calc//Tsmn.csv', 'w', newline='') as csvfile:
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
		
		T = Tsmn (x,y,z,d)
		Twriter.writerow([T])
		
		i = i+1

