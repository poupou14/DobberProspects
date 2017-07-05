#!/usr/bin/python 
import string, sys
from DPParser import DPParser
#from DPWriter import DPWriter

def main():
	myDP = DPParser()
#	myDPWriter = DPWriter("../INPUT/DP_Players.xls")
	sourceFile_l = "/home/poulnais//Developpement/DobberProspects/INPUT/TEST-2017.xls"

	targetFile_l = "/home/poulnais//Developpement/DobberProspects/OUTPUT/TEST-2017-out.xls"
	myDP.readDP(sourceFile_l, targetFile_l)

main()
