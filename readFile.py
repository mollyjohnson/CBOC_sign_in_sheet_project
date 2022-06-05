# Author: Molly Johnson
# Date:
# Description:

# CBOC file name
CBOC = 'CBOCs.txt'
TRAILER = ', SMP'

####################################################################
### Function Title:
### Arguments:
### Returns:
### Description: 
###################################################################
# Check whether the CBOC needs an SMP check line or not


####################################################################
### Function Title:
### Arguments:
### Returns:
### Description: 
###################################################################
# open CBOC.txt file
file = open(CBOC,'r')

# use readlines to read all lines in the text file and
# return the file contents as a list of strings
lines = []
lines = file.readlines()

# count num CBOCs from the file
numCBOCs = 0

# go through each line and strip the added newline character and increment number of CBOCs
for line in lines:
    print(line.rstrip('\n'))
    numCBOCs += 1

print('\n' + str(numCBOCs))

# close .txt file
file.close()