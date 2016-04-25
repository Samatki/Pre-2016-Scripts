import random
import csv

print('How many results to generate? Up to 1 million')

# Tensile Strength South (Max)
def Tsmx (x,y,z,d):
	T = -1.528141851*(10**(-5))*(z**3)+2.065775685*(10**(-3))*(z**2)*w+5.436597386*(10**(-3))*(z**2)*x+7.326777103*(10**(-3))*(z**2)*y-8.771130271*(10**(-2))*z*(w**2)+0.554091815*z*w*x+3.844746346*(10**(-1))*z*w*y-3.981768863*(10**(-1))*z*(x**2)-7.838463029*(10**(-1))*z*x*y-3.528449983*(10**(-1))*z*(y**2)-1.928378104*(w**3)+6.282623043*(w**2)*x+19.65018603*(w**2)*y+4.129417133*w*(x**2)-9.94251569*w*x*y-34.75871007*w*(y**2)-1.17628472*(x**3)+1.275738571*(x**2)*y+13.4067192*x*(y**2)+27.02955258*(y**3)-1.959782564*(10**-1)*z**2-9.665826051*z*w+7.578542238*z*x+9.216165582*z*y-191.9746806*(w**2)-366.6722482*w*x+171.0700032*w*y+29.9161642*(x**2)-30.0808935*x*y-502.5741669*(y**2)+23.51598799*z+6099.080884*w+2784.254363*x+4336.709359*y-72645.85398
	#print(T)
	return(T)

numCases = int(input())

with open('Tsmx.csv', 'w', newline='') as csvfile:
	Twriter = csv.writer(csvfile)
	


	for i in range (0,numCases,1):

	### All input data are Gaussian Distributed ###

	# w = Hmx = Wave Height(m), mean 20.91, STDEV 4.27
		w = random.gauss(20.91,4.27)
	# x = Offs = Platform Offset (m), mean 13.00, STDEV 2.60
		x = random.gauss(13.00,2.60)
	# y = Tass = Wave Period (sec), mean 13.49, STDEV 2.24
		y = random.gauss(13.49,2.24)
	# z = Dwc = Wave Current Direction (deg), mean 186.56, STDEV 93.58
		z = random.gauss(186.56,93.58)

	# d = D = common link diameter - Uniform Distribution between negative tolerance of -2mm and positive tolerance of 5%, Mean value of 0.076m, (m)
		d = random.uniform(0.074,1.05*0.076)
		
		T = Tsmx (x,y,z,d)
		Twriter.writerow([T])
		
		i = i+1
		
		
