# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      msommer
#
# Created:     21.01.2015
# Copyright:   (c) msommer 2015
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import os, profile
from DEIO_datahandler import IPTS
#from DEIO_helper import DeioHelper  #nur zum testen
def main():


    path_sut= "K:\\Projekt_DEIO\\Data\\IO_SUT\\"
    dir_output='K:\\Projekt_DEIO\\Data\\_InputData'
    name_cntry= "AUT_hard.xlsx"
    timespan=["2008","2009","2010"]

    # initialize IPTS Extraction
    ipts=IPTS(path_sut,name_cntry,timespan,dir_output,nace="2008")


    #print ipts.main_tables_data["USE_BP"][0]
    #trg= ipts.get_QN()
    #trg=ipts.get_VA()

    ipts.genr_data(varname="-",all=True)



    #ipts.genr_r_TLS_imed()
    #ipts.genr_r_TTM_imed()

    ipts.aggreg_SUT_data()


if __name__ == '__main__':
    main()


#profile.run('print main(); print',sort=1)