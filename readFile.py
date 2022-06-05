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
def isSMPCBOC(line):
    if(TRAILER in line):
        return True
    return False

####################################################################
### Function Title:
### Arguments:
### Returns:
### Description: 
###################################################################
def getCBOClists(smpCBOCs, noSMPCBOCs):
    # open CBOC.txt file
    file = open(CBOC,'r')

    # use readlines to read all lines in the text file and
    # return the file contents as a list of strings
    lines = file.readlines()

    # go through each line and strip the added newline character and increment number of CBOCs
    for line in lines:
        # strip the added newline character readlines() adds
        line = line.rstrip('\n')

        # check if CBOC needs an smp line. if so remove trailer indicating is an SMP
        # clinic from the string before adding to list
        if(isSMPCBOC(line) == True):
            smpCBOCs.append(line.rstrip(TRAILER))
        # if CBOC does not need an smp line, add it to other list as is
        else:
            noSMPCBOCs.append(line)

    # close .txt file
    file.close()

    # return list of SMP CBOCs and non SMP cbocs
    return smpCBOCs, noSMPCBOCs