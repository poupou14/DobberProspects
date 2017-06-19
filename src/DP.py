#!/usr/bin/python 
import string, sys
from DPParser import DPParser
#from DPWriter import DPWriter

def main():
	myDP = DPParser()
#	myDPWriter = DPWriter("../INPUT/DP_Players.xls")
	sourceFile_l = "/home/poupou/Development/DobberProspects/INPUT/Free-Agent-2016.xls"

	targetFile_l = "/home/poupou/Development/DobberProspects/OUTPUT/Free-Agent-2016-out.xls"
	myDP.readDP(sourceFile_l, targetFile_l)

main()
