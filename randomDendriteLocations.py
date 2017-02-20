# -*- coding: utf-8 -*-
"""
Created on Wed Feb 15 11:36:40 2017

@author: Steve Ryan

Purpose - 
We don't currently have a complete sterelogical solution for analyzing 
dendritic spines of individuall filled neurons.  This script allows us to 
import an excel sheet representing the dendrogram and generate a set of 
pseudorandomly selected points on the dendritic arbor for followup analysis.

Requirements - 
1: Import dendrogram data from a reconstructed neuron in an Excel / CSV format
2: Generate a set of random positions in the dendrogram for followup imaging 
   at 100X
 2a:  This will be conducted iteratively so each successive random number can 
      be compared to the existing list of numbers to reject any that would 
      overlap.  For past projects we've always kept the overlap threshold at 
      50 microns, so for now I'm just hard-coding this to do the same.'
 2b:  This function will also accept as INPUT a list of pre-existing random 
      locations so that, if we're modifying the set from an existing neuron, 
      we don't have to totally start over from scratch.
      
      This will involve a few basic steps
      1: import workbook
      2: identify whether there are pre-existing values
      3a: if there are, extract them
      3b: if there are not, carry on
      4: generate new random values
      5: format output
      6: write output to spreadsheet.
"""

import xlrd
import random
import math
from numpy import array

def main():
    filename = "BLA 1-14 #36-2 10x 488.xlsx" # need to get this from interface
    dend_workbook = xlrd.open_workbook(filename)
    first_sheet = dend_workbook.sheet_by_index(0)
    segment_column = first_sheet.col_values(1)
    
#    if first_sheet.ncols > 2:
#        fourth_column = first_sheet.col_values(3)
#        fifth_column = first_sheet.col_values(4)
        # third_column and fourth_column don't have obviously cleaner names
        # they refer to the columns added to the spreadsheet by the previous v
        # version of this function.  They store the locations selected by that 
        # script.  For example, next to a segment of length 294 which has been 
        # selected may be, in column_three, the numeral 1098.  The Thousands place
        # is superfluous, and the 98 indicates 98 microns from the proximal point
        # of origin for this segment.  We'll need to convert these to numerical
        # values that make sense in the schema used here.
    
    running_total_column = []
    running_total = 0
    for i in segment_column:
        running_total += i
        running_total_column.append(running_total)

    existing_list = [12, 150, 400, 2009] # toy list, just for testing
    number_existing = len(existing_list) # establish number of pre-existing points
    number_to_randomize = 10 - number_existing # hardcodes 10 segments as desired output

    updated_list = randomLocation(existing_list, number_to_randomize, segment_column, running_total_column)
    print(updated_list)


def randomLocation(existing_list, number_of_locations, dendrite_segment_list, running_total_column):
    # Input - existing_list
    #           List of random points already selected that are being kept, if
    #           any.  This can be empty.  
    # Input - number_of_locations
    #           Total number of locations to generate
    # Input - dendrite_segment_list
    #           scaling factor for random numbers, so the results can be 
    #           returned in terms of the cell's dendritic arbor instead of 
    #           abstractly on the scale from 0 to 1
    # Output - new_list
    #           new rnandomly generated points

    new_list = existing_list
    number_to_generate = 10 - len(new_list)
    if number_to_generate <= 0:
        return new_list

    total_length = running_total_column[len(running_total_column)-1]

    for i in range(number_to_generate):
        new_random_is_valid = False
        while True:
            # generate random location
            # test validity, must be more than 50 um from existing points
            # repeat or break based on outcome of validity test
            new_location = random.randint(1, math.floor(total_length))
            new_list_array = array(new_list)
            new_list_array = abs(new_list_array - new_location)
            if min(new_list_array) < 50:
                print(new_list)
#                print(new_list_array)
                print("{} is Invalid selection, regenerating".format(new_location))
                new_random_is_valid = False
            else:
                new_random_is_valid = True
            
            if new_random_is_valid:
                break
        new_list.append(new_location)

    return new_list


def convertOldList(segment_column, third_column, fourth_column):
    # stuff
    return 0
    
    

# 
if __name__ == "__main__":
    main()