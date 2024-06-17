all: 
	make test3

test3:
	node /Users/ajdalesandro/Applications/NodeChaseSpreadsheetAnalyzer/src/index.js /Users/ajdalesandro/Desktop/process/Chase7285_Activity_20240517.CSV "/Users/ajdalesandro/Google Drive/My Drive/04_Finances/Analysis.xlsx"

test1:
	node src/index.js ../../../Desktop/process/*.CSV

test2:
	node test/test.js ../../../Desktop/process/*.xlsx

