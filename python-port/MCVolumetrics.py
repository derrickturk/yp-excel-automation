#!python3
#MCVolumetrics.py - a simple MonteCarlo program

import copy, random, math #built-in python libraries

MaxAllowedPorosity=0.4
#Make a couple of empty dictionaries that'll be copied and used in the program.
#This is done to minimize mistakes in the program.

Realization={'Area':[],'Height':[],'Porosity':[],'OilSaturation':[],'FVF':[]}

DistParams={'Area':[],
'MinHeight':[],
'MaxHeight':[],
'MeanPorosity':[],
'SDPorosity':[],
'MinOilSaturation':[],
'MaxOilSaturation':[],
'MinFVF':[],
'MaxFVF':[]}


def MCAverageOOIP(trials,seed):
	rngseed=random.Random(seed)
	params={'Area':40,
	'MinHeight':10,
	'MaxHeight':50,
	'MeanPorosity':0.1,
	'SDPorosity':0.01,
	'MinOilSaturation':0.3,
	'MaxOilSaturation':0.5,
	'MinFVF':1.2,
	'MaxFVF':1.5}

	MCAverageOOIP=Average(OOIPForSamples(MonteCarlo(trials,params,rngseed)))
	return MCAverageOOIP
def Average(values): #takes a list of doubles, returns a double value.
	runningAvg=0
	
	for i in range(len(values)):
		runningAvg=(values[i]/len(values))+runningAvg
	return runningAvg

def OOIPForSamples(samples): #takes a list of Realization dic, returns a list.
	result=[]
	for i in range(len(samples)):
		result.append(OOIP(samples[i]))
	return result

def OOIP(sample): #takes a Realization dic and returns a double value.
	OOIP=7758*sample['Area']*sample['Height']*sample['Porosity']*sample['OilSaturation']/sample['FVF']
	return OOIP 
	
def MonteCarlo(trials,params,rngseed): #takes an int, DistParams dic, int and returns a list of dictionaries.
	result=[]
	for i in range(trials):
		result.append(MonteCarloRealization(params,rngseed))
	return result
	
def MonteCarloRealization(params,rngseed): #takes a DistParams dic and rngseed, returns a dictionary.
	result=copy.deepcopy(Realization) #make a deepcopy of DistParams
	result['Area']=params['Area']
	
	result['Height']=BoundedRandUniform(params['MinHeight'],params['MaxHeight'],rngseed)
	result['Porosity']=BoundedRandNormal(params['MeanPorosity'],params['SDPorosity'],rngseed)
	if result['Porosity']<0:
		result['Porosity']= 0
	elif result['Porosity'] > MaxAllowedPorosity:
		result['Porosity']=MaxAllowedPorosity
	
	result['OilSaturation']=BoundedRandUniform(params['MinOilSaturation'],params['MaxOilSaturation'],rngseed)
	result['FVF']=BoundedRandUniform(params['MinFVF'],params['MaxFVF'],rngseed)
	return result
	
def BoundedRandNormal(mean,sd,rngseed): #takes two doubles and a rngseed, returns a single value.
	u1=rngseed.uniform(0,1)
	u2=rngseed.uniform(0,1)

	r=math.sqrt(-2.0*math.log(u1))
	theta=2.0*math.pi*u2
	
	BRN=r*math.cos(theta)
	
	return BRN

def BoundedRandUniform(lower,upper,rngseed): #takes two doulbes and a rngseed, returns a single value.
	return rngseed.uniform(0,1)*(upper-lower)+lower

	
	
	