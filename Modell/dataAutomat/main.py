# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      msommer
#
# Created:     13.01.2015
# Copyright:   (c) msommer 2015
# Licence:     <your licence>
#-------------------------------------------------------------------------------

from getDEIOdata import GetSUT_IPTS
import os

def main():

    path_sut= "K:\\Projekt_DEIO\\Data\\IO_SUT\\"
    name_cntry= "AUT_hard.xlsx"
    timespan=["08","09","10"]

    gdd = GetSUT_IPTS(path_sut,name_cntry,timespan)


##    # GET SINGLE ROWS/COLUMNS
##    selection=[]
##    selection.append(["use","pp","D94:BY94","QN","Sector Gross Output"])
##    selection.append(["tls","","D82:BY82","TLS","Sector Taxes Less Subsidy"])
##    selection.append(["use","bp","D91:BY91","OPTX","Other net taxes on production"])
##    selection.append(["use","bp","D94:BY94","SURPLUS","Sector Surplus gross"])
##    selection.append(["use","bp","D89:BY89","L","Sector Labour Compensation"])
##    selection.append(["use","bp","D89:BY89","WAGE","Sector Wages"])
##    selection.append(["sup","bp","BZ8:BZ81","QG","Commodity Production"])
##    selection.append(["useimp","bp","CP8:CP81","MG","Commodity Imports"])
##    selection.append(["useimp","bp","D82:BY82","MN","Sector Imports"])
##
##
##    for info in selection:
##        data=gdd.get_array(info[0],info[1],info[2])  # Extract Data from XLSX
##        gdd.write_array(info[3],data,info[4])       # WRITE IN new XLSX


    # GET AGGREGATED ROWS/COLUMNS
    # EN
    data1=gdd.get_array("use","bp","D11:by11")
    data2=gdd.get_array("use","bp","D21:by21")
    data3=gdd.get_array("use","bp","D36:by36")

    print data1
    print data2
    print data3

    en = gdd.add_arrays(data1,data2)
    print en
    en = gdd.add_arrays(en,data3)

    print en
    gdd.write_array("EN",en,"Energy Commodity Inputs")


#    print gdd.add_arrays([1,2],[3,4])

if __name__ == '__main__':
    main()
