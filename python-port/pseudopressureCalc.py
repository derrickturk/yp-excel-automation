#!Python3
#!pseudopressureCalc - a simple python program to calculate pseudo pressure in a well.
viscosity=0.05
compressibility=0.9
steps=100
def PseudoPressure(pRef,p):
	PPressure=2.0*TrapezodialRule(pRef,p,steps)
	return PPressure 

def TrapezodialRule(lower,upper,nSteps):
	integral=0.0

	step=(upper-lower)/nSteps
	
	l=lower
	u=lower+step
	
	fAtL=Integrand(l)
	fAtU=Integrand(u)
	
	for i in range(0,nSteps):
		integral=integral+0.5*step*(fAtL+fAtU)
		l=u
		fAtL=fAtU
		u=u+step
		fAtU=Integrand(u)
		i=i+1
	return integral

def Integrand(p):
	Ingnd=p/(viscosity*compressibility)
	return Ingnd

Output1=PseudoPressure(100,100)
Output2=PseudoPressure(100,1000)
Output3=PseudoPressure(1000,100)
print(str(Output1)+'\n'+str(Output2)+'\n'+str(Output3)+'\n')
