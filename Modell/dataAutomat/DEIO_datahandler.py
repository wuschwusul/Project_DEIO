# -*- coding: utf-8 -*-

# EXCEL read write librarys
from openpyxl import load_workbook
from openpyxl.workbook import Workbook

##import numpy as np

from DEIO_helper import DeioHelper
from DEIO_config import Config

import os
import random as rnd

class IPTS:

    filepath  =  ""                         # full path of SUT
##    wb  =  0                                # active workbook object
    timespan  =  ""                         # eg["2008","2009","2010"]
    dir_output  =  ""
    variable_dict = {}
    va_detail = {}
    hlpr  =  ""                             #helper instance

    #NACE-Sectors
    nace  =  ""                             #classification nace  =  2003 or 2008
    sectors  =  ""                        # eg.['01', '02', '03', '05',...
    sectors_engy  =  ""                     # eg. ['05','19','35']
    sectors_nonengy  =  ""                  # autom. generated

    base_areas  =  {}                        # BASE areas - depending on SUT size
                                             #eg.: {"USE_SUP": "D8:BY81", "FD": "CA8:CP81", "VA":  =  ..

    main_tables_info  =  {}                 # describes where the Main Tables can be found in the Excel Sheet
                                            # Main Tables are : SUP_BP, USE_FD_PP, USE_BP, USEDOM_FD_BP, TTM

                                            #e.g. "SUP_PP":["BS_SUP_USE",["sup","pp"]],
                                            #       "TTM": ["BS_SUP_USE",["ttm",""]],...

    main_tables_data  =  {}                 # DATA - dictionary containing main data matrixes

                                            #eg. {"SUP_BP": [2013.121,0,0,101.2 ..],
                                            #     "USE_PP": [95.2,0,7,45.2 ..],
                                            #       "TTM": [0,0.0.1,1021,32,0,...]

                                            # by e.g. main_tables_data["VA"]  -> returns 3d matrix [yr][row][col]


    variable_names  = {}


    def __init__(self,path_sut,name_cntry,timespan,dir_output,nace):

        self.nace  =  nace                                      # classification (NACe)
        self.filepath  =  os.path.join(path_sut,name_cntry)   # path source
        self.timespan  =  timespan                              # timespan
        self.dir_output  =  dir_output                          # path output


        #init config file
        cf  =  Config()

        #init sectors
        self.sectors  =  cf.get_sectors(nace)
        self.sectors_engy  =  cf.get_engy_sectors(nace)
        self.sectors_nonengy  =  cf.get_nonengy_sectors(nace)


        # matrix-helper instance
        self.hlpr  =  DeioHelper(self.filepath,name_cntry,timespan,dir_output,self.sectors)

        #init Base Positions/Areas
        self.base_areas  =  cf.get_base(nace)

        #init main Tables
        self.main_tables_info  =  cf.get_main_tables_info()          # assign Variables to MainAreas

        #init Range of Variables
        self.variable_names = cf.get_IPTS_variable_names()
        self.variable_dict = cf.get_IPTS_variable_dict()
        self.va_detail = cf.get_va_detail(nace)



#-----------LOAD DATA METHODS------------------------------------------------------------

    def load(self,label):
        if label=="all":
            self.load_main_data_all()

        else:
            self.load_main_data(label)


#---- Sub Loading methods----
    def load_main_data_all(self):
        for ma in self.main_tables_info:
            self.load_main_data(ma)


#---- Sub Loading methods----

    def load_main_data(self,main_label):   # load main data in class variables in style [yr][row][col]
        # main_label  =  SUP_PP

        #---prepare information---
        base_pos  =  self.main_tables_info[main_label][0]
        sheetname  =  self.main_tables_info[main_label][1]          #SUP_PP->["sup","pp"] ; USE_CONS_PP -> ["use","pp"]

        data_area  =  self.base_areas[base_pos]               # eg. D8:BY81
        #print ("Load - %s " %(main_label)),

        #---get data from files---
        tmp  =  []
        for yr in self.timespan:   #eg yr =2008,2009,2010
            tmp.append(self.hlpr.get_data(sheetname[0]+yr[2:]+sheetname[1],data_area))  # sup08pp , A21:B22

        self.main_tables_data[main_label]  =  tmp
        #print ("...OK")


##    def genr(self,name):
##        #self.genr_QN()
##        pass
##


##        # D         -> share of rowsum
##        # XD        -> sum over selected rows
##        # SXD       -> delegint rows, punctual division

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#-----------GENERATE DATA META METHOD ------------------------------------------------------------
# Generate = Load, extract and write to excel
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    def genr_data(self,varname="-", all=False):
        # varname = Name of Variable (i.e. QN,DN...)
        # all -> generate all variables

        if varname=="-" and all==False:    # error catching
            return False

        # Artifical Functinos

        # get all
        if all==True:
            for vn in self.variable_names:
                print("Generating %s - " %vn),
                getattr(self,'genr_'+vn)()
                print("...OK")
        # get one
        else:
            getattr(self,'genr_'+varname)()



    def aggreg_SUT_data(self):
        self.hlpr.aggregate_data_of_a_directory("SUT_all")



#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#--------GENERATE METHODS-------------
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# Generate = Load, extract and write to excel


    def get_main_data(self,main_label):  # checks if Main_data is loaded and returns main data
        # label = "USE_BP", "TTM"...

        # LOAD data if missing
        if main_label not in self.main_tables_data:
            self.load_main_data(main_label)

        # return data
        return self.main_tables_data[main_label]


#-------------------------------------------------------
    def genr_QN(self):  ##    QN   -> sum over rows
        # LOAD
        v_name="QN"
        source_data = self.get_main_data("SUP_BP")

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_aggreg_rows(source_data)

        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_QG(self): ## QG        -> sum over colums
        # LOAD
        v_name="QG"
        source_data = self.get_main_data("SUP_BP")

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_aggreg_cols(source_data)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_VA(self):  ## extract 1 row
        # LOAD
        v_name="VA"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row) #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_DN(self):  # sum over rows
        # LOAD
        source_data = self.get_main_data("USEDOM_BP")
        v_name="DN"

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_aggreg_rows(source_data)
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_MN(self):  ## extract rows
        # LOAD
        v_name="MN"
        source_data = self.get_main_data("USEIMP_BP")

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_aggreg_rows(source_data)
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_EDN(self):  ## extract rows
        # LOAD
        source_data = self.get_main_data("USEDOM_BP")
        v_name="EDN"

        # EXTRACT desired data - EDN
        #rows = [3,13,28]

        rows=[]
        for en in self.sectors_engy:
            rows.append(self.sectors.index(en))

        target1 = self.hlpr.IPTS_extract_rows(source_data,rows) # extracting 3 rows - aggregation needed
        target1 = self.hlpr.IPTS_aggreg_rows(target1)
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_XD(self):  ## extract rows
        # LOAD
        v_name="XD"
        source_data = self.get_main_data("USEDOM_BP")


        # EXTRACT desired data - EDN
        rows=[]
        for nen in self.sectors_nonengy:
            rows.append(self.sectors.index(nen))

        target1 = self.hlpr.IPTS_extract_rows(source_data,rows) # extracting 74-3 rows - aggreg needed
        target1 = self.hlpr.IPTS_aggreg_rows(target1)  # aggregating
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])




#-------------------------------------------------------
    def genr_EMN(self):  ## extract rows
        # LOAD
        source_data = self.get_main_data("USEIMP_BP")
        v_name="EMN"

        # EXTRACT desired data - EMN
        rows=[]
        for en in self.sectors_engy:
            rows.append(self.sectors.index(en))

        target1 = self.hlpr.IPTS_extract_rows(source_data,rows)# extracting 3 rows - aggregation needed
        target1 = self.hlpr.IPTS_aggreg_rows(target1)  # aggregating
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_XM(self):  ## extract rows
        # LOAD
        v_name="XM"
        source_data = self.get_main_data("USEIMP_BP")


        # EXTRACT desired data - XM
        rows=[]
        for nen in self.sectors_nonengy:
            rows.append(self.sectors.index(nen))

        target1 = self.hlpr.IPTS_extract_rows(source_data,rows) # extracting 74-3  rows - aggregation needed
        target1 = self.hlpr.IPTS_aggreg_rows(target1)  # aggregating
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_TLS(self):  ##    QN   -> sum over rows
        # LOAD
        v_name="TLS"
        source_data = self.get_main_data("TLS")

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_aggreg_rows(source_data)
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

###-------------------------------------------------------
##    def genr_TTM(self):  ##    QN   -> sum over rows
##        # LOAD
##        v_name="TTM"
##        source_data = self.get_main_data("TTM")
##
##        # EXTRACT desired data
##        target1 = self.hlpr.IPTS_aggreg_rows(source_data)
##        target1 = self.hlpr.IPTS_transpose(target1)
##
##        # WRITE to excel
##        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_L(self):  ## extract 1 row
        # LOAD
        v_name="L"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_WAGES(self):  ## extract 1 row
        # LOAD
        v_name="WAGES"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_OPTX(self):  ## extract 1 row
        # LOAD
        v_name="OPTX"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name])

        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_DEPR(self):  ## extract 1 row
        # LOAD
        v_name="DEPR"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_SURPLUS_GR(self):  ## extract 1 row
        # LOAD
        v_name="SURPLUS_GR"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_SURPLUS_NT(self):  ## extract 1 row
        # LOAD
        v_name="SURPLUS_NT"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_AGB(self):  ## extract 1 row
        # LOAD
        v_name="AGB"
        v_name1="L"
        v_name2="WAGES"
        source_data = self.get_main_data("VA")

        # EXTRACT desired data
        data_L = self.get_L()
        data_WAGES = self.get_WAGES()


        # subtract
        target1 = self.hlpr.IPTS_sub_elem(data_L,data_WAGES)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_K(self):  ## extract 1 row
        # LOAD
        v_name = "K"
        v_name1="SURPLUS_GR"
        v_name2="OPTX"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name1])
        va_row.append(self.va_detail[v_name2])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 2 rows - no aggreg needed
        target1 = self.hlpr.IPTS_aggreg_rows(target1)
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])




#-------------------------------------------------------
    def genr_EN(self):  ## extract rows
        # LOAD
        v_name="EN"
        source_data = self.get_main_data("USE_BP")


        # EXTRACT desired data - EDN
        rows=[]

        for engy in self.sectors_engy:
            rows.append(self.sectors.index(engy))


        target1 = self.hlpr.IPTS_extract_rows(source_data,rows)  #extracting 3 rows -  aggreg needed
        target1 = self.hlpr.IPTS_aggreg_rows(target1)  #aggregating
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_MG(self): ## QG        -> sum over colums
        # LOAD
        v_name="MG"
        source_data = self.get_main_data("USEIMP_BP")
        source_data2 = self.get_main_data("USEIMP_FD_BP")

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_aggreg_cols(source_data)  # intermed MG

        rows=[0,1,2,4,5,6,9,12]
        target2 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand MG
        target2 = self.hlpr.IPTS_aggreg_cols(target2)

        target1 = self.hlpr.IPTS_add_elem(target1,target2)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])




#-------------------------------------------------------
    def genr_TLSQ(self):  ##    QN   -> sum over rows

        # LOAD
        v_name="TLSQ"

        # EXTRACT desired data
        QN = self.get_QN()
        TLS = self.get_TLS()

        target1 = self.hlpr.IPTS_div_elem(TLS,QN)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_OPTXQ(self):  ##    QN   -> sum over rows

        # LOAD
        v_name="OPTXQ"

        # EXTRACT desired data
        OPTX = self.get_OPTX()
        QN = self.get_QN()

        target1 = self.hlpr.IPTS_div_elem(OPTX,QN)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_SKQ(self):  ##    QN   -> sum over rows

        # LOAD
        v_name="SKQ"

        # EXTRACT desired data
        K = self.get_K()
        QN = self.get_QN()

        target1 = self.hlpr.IPTS_div_elem(K,QN)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_SLQ(self):  ##    QN   -> sum over rows

        # LOAD
        v_name="SLQ"

        # EXTRACT desired data
        L = self.get_L()
        QN = self.get_QN()

        target1 = self.hlpr.IPTS_div_elem(L,QN)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_SXDQ(self):  ##    QN   -> sum over rows

        # LOAD
        v_name="SXDQ"

        # EXTRACT desired data
        XD = self.get_XD()
        QN = self.get_QN()

        target1 = self.hlpr.IPTS_div_elem(XD,QN)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_SXMQ(self):  ##    QN   -> sum over rows

        # LOAD
        v_name="SXMQ"

        # EXTRACT desired data
        XM = self.get_XM()
        QN = self.get_QN()

        target1 = self.hlpr.IPTS_div_elem(XM,QN)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

#-------------------------------------------------------
    def genr_SEQ(self):  ##    QN   -> sum over rows

        # LOAD
        v_name="SEQ"

        # EXTRACT desired data
        EN = self.get_EN()
        QN = self.get_QN()

        target1 = self.hlpr.IPTS_div_elem(EN,QN)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_D(self):  ##    QN   -> sum over rows

        # LOAD
        v_name = "D"

        source_data = self.get_main_data("SUP_BP")

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_share_over_cols(source_data)
        ##target1 = self.hlpr.IPTS_transpose(target1)
        target1 = self.hlpr.IPTS_make_rows_to_1column(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_SXD(self):

        # LOAD
        v_name = "SXD"
        source_data = self.get_main_data("USEDOM_BP")

        # EXTRACT desired data - XD, SXD
        # first XD
        rows=[]
        for nen in self.sectors_nonengy:
            rows.append(self.sectors.index(nen))

        predata = self.hlpr.IPTS_extract_rows(source_data,rows) # whole 71x71 XD matrix

        # second get SXD
        target1 = self.hlpr.IPTS_share_over_rows(predata)  # calc shares
        target1 = self.hlpr.IPTS_make_rows_to_1column(target1)  # rearrange data to single columns for each year

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------

    def genr_SXM(self):

        # LOAD
        v_name = "SXM"
        source_data = self.get_main_data("USEIMP_BP")

        # EXTRACT desired data - XM, SXM
        # first XM
        rows=[]
        for nen in self.sectors_nonengy:
            rows.append(self.sectors.index(nen))

        predata = self.hlpr.IPTS_extract_rows(source_data,rows) # whole 71x74 XM matrix

        # second get SXD
        target1 = self.hlpr.IPTS_share_over_rows(predata)  # calc shares
        target1 = self.hlpr.IPTS_make_rows_to_1column(target1)  # rearrange data to single columns for each year

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_SED(self):

        # LOAD
        v_name = "SED"
        source_data = self.get_main_data("USEDOM_BP")

        # EXTRACT desired data - EN, SED
        # first EN
        datarow = self.get_EN()
        datarow = self.hlpr.IPTS_transpose(datarow)  # EN from column to row



        # second get EDN
        source_data2 = self.get_main_data("USEDOM_BP")

        rows=[]
        for en in self.sectors_engy:
            rows.append(self.sectors.index(en))

        data_b = self.hlpr.IPTS_extract_rows(source_data2,rows) # extracting 3 rows - aggregation needed


        ##print("EN yrs: %s rows: %s  cols: %s " %(len(datarow),len(datarow[0]),len(datarow[0][0])))
        ##print("EDN yrs: %s rows: %s  cols: %s " %(len(data_b),len(data_b[0]),len(data_b[0][0])))

        target1 = self.hlpr.IPTS_divide_by_row(data_b,datarow)
        target1 = self.hlpr.IPTS_make_rows_to_1column(target1)  # rearrange data to single columns for each year

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


#-------------------------------------------------------
    def genr_SEM(self):

        # LOAD
        v_name = "SEM"
        source_data = self.get_main_data("USEIMP_BP")

        # EXTRACT desired data - EN, SED
        # first EN
        datarow = self.get_EN()
        datarow = self.hlpr.IPTS_transpose(datarow)  # EN from column to row
        # second get EDN
        data_b = self.get_EMN_detail()

        target1 = self.hlpr.IPTS_divide_by_row(data_b,datarow)
        target1 = self.hlpr.IPTS_make_rows_to_1column(target1)  # rearrange data to single columns for each year

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])



    def genr_TTM_imed(self):
        # LOAD
        v_name = "TTM_imed"

        # EXTRACT
        source_data = self.get_main_data("TTM")
        target1 = self.hlpr.IPTS_make_rows_to_1column(source_data)

        # WRITE 2 xls
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])



#------------------------FINAL DEMAND ELEMENETS----------------------

##-----------------------------------------------------
    def genr_FM(self):
        # LOAD
        v_name="FM"
        source_data = self.get_main_data("USEIMP_FD_BP")

        # EXTRACT desired data

        rows=[0,1,2,4,5,6,9,12]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_FD(self):
        # LOAD
        v_name="FD"
        source_data = self.get_main_data("USEDOM_FD_BP")

        # EXTRACT desired data

        rows=[0,1,2,4,5,6,9,12]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_F(self):
        # LOAD
        v_name="F"
        source_data = self.get_main_data("USE_FD_BP")

        # EXTRACT desired data

        rows=[0,1,2,4,5,6,9,12]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_CP_bp(self):
        # LOAD
        v_name="CP_bp"
        source_data = self.get_main_data("USE_FD_BP")

        # EXTRACT desired data

        rows=[0,1]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_CP_pp(self):
        # LOAD
        v_name="CP_pp"
        source_data = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[0,1]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_CG_bp(self):
        # LOAD
        v_name="CG_bp"
        source_data = self.get_main_data("USE_FD_BP")

        # EXTRACT desired data

        rows=[2]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_CG_pp(self):
        # LOAD
        v_name="CG_pp"
        source_data = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[2]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_EXP_bp(self):
        # LOAD
        v_name="EXP_bp"
        source_data = self.get_main_data("USE_FD_BP")

        # EXTRACT desired data

        rows=[9,12]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_EXP_pp(self):
        # LOAD
        v_name="EXP_pp"
        source_data = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[9,12]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_GFCF_bp(self):
        # LOAD
        v_name="GFCF_bp"
        source_data = self.get_main_data("USE_FD_BP")

        # EXTRACT desired data

        rows=[4,5]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_GFCF_pp(self):
        # LOAD
        v_name="GFCF_pp"
        source_data = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[4,5]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_ST_bp(self):
        # LOAD
        v_name="ST_bp"
        source_data = self.get_main_data("USE_FD_BP")

        # EXTRACT desired data

        rows=[6]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_ST_pp(self):
        # LOAD
        v_name="ST_pp"
        source_data = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[6]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


    def genr_mcp(self):
        # LOAD
        v_name="mcp"

        source_data = self.get_main_data("USE_FD_BP")
        source_data2 = self.get_main_data("USEIMP_FD_BP")

        # EXTRACT CP

        rows=[0,1]
        target0 = self.hlpr.IPTS_extract_cols(source_data,rows)
        target0 = self.hlpr.IPTS_aggreg_cols(target0)

        # EXTRACT MCP

        target2 = self.hlpr.IPTS_extract_cols(source_data2,rows)
        target2 = self.hlpr.IPTS_aggreg_cols(target2)

        target1 = self.hlpr.IPTS_div_elem(target2,target0)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])



    def genr_mcg(self):
        # LOAD
        v_name="mcg"

        source_data = self.get_main_data("USE_FD_BP")
        source_data2 = self.get_main_data("USEIMP_FD_BP")

        # EXTRACT Total

        rows=[2]
        target0 = self.hlpr.IPTS_extract_cols(source_data,rows)
        target0 = self.hlpr.IPTS_aggreg_cols(target0)

        # EXTRACT Imported

        target2 = self.hlpr.IPTS_extract_cols(source_data2,rows)
        target2 = self.hlpr.IPTS_aggreg_cols(target2)

        target1 = self.hlpr.IPTS_div_elem(target2,target0)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


    def genr_mgfcf(self):
        # LOAD
        v_name="mgfcf"

        source_data = self.get_main_data("USE_FD_BP")
        source_data2 = self.get_main_data("USEIMP_FD_BP")

        # EXTRACT CP

        rows=[4,5]
        target0 = self.hlpr.IPTS_extract_cols(source_data,rows)
        target0 = self.hlpr.IPTS_aggreg_cols(target0)

        # EXTRACT MCP

        target2 = self.hlpr.IPTS_extract_cols(source_data2,rows)
        target2 = self.hlpr.IPTS_aggreg_cols(target2)

        target1 = self.hlpr.IPTS_div_elem(target2,target0)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


    def genr_mst(self):
        # LOAD
        v_name="mst"

        source_data = self.get_main_data("USE_FD_BP")
        source_data2 = self.get_main_data("USEIMP_FD_BP")

        # EXTRACT Total

        rows=[6]
        target0 = self.hlpr.IPTS_extract_cols(source_data,rows)
        target0 = self.hlpr.IPTS_aggreg_cols(target0)

        # EXTRACT Imported

        target2 = self.hlpr.IPTS_extract_cols(source_data2,rows)
        target2 = self.hlpr.IPTS_aggreg_cols(target2)

        target1 = self.hlpr.IPTS_div_elem(target2,target0)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

    def genr_mex(self):
        # LOAD
        v_name="mex"

        source_data = self.get_main_data("USE_FD_BP")
        source_data2 = self.get_main_data("USEIMP_FD_BP")

        # EXTRACT Total

        rows=[9,12]
        target0 = self.hlpr.IPTS_extract_cols(source_data,rows)
        target0 = self.hlpr.IPTS_aggreg_cols(target0)

        # EXTRACT Imported

        target2 = self.hlpr.IPTS_extract_cols(source_data2,rows)
        target2 = self.hlpr.IPTS_aggreg_cols(target2)

        target1 = self.hlpr.IPTS_div_elem(target2,target0)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])



##-----------------------------------------------------
    def genr_TLS_CP(self):
        # LOAD
        v_name="TLS_CP"
        source_data = self.get_main_data("TLS_FD")

        # EXTRACT desired data

        rows=[0,1]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_TLS_CG(self):
        # LOAD
        v_name="TLS_CG"
        source_data = self.get_main_data("TLS_FD")

        # EXTRACT desired data

        rows=[2]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_TLS_GFCF(self):
        # LOAD
        v_name="TLS_GFCF"
        source_data = self.get_main_data("TLS_FD")

        # EXTRACT desired data

        rows=[4,5]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_TLS_EXP(self):
        # LOAD
        v_name="TLS_EXP"
        source_data = self.get_main_data("TLS_FD")

        # EXTRACT desired data

        rows=[9,12]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_TLS_ST(self):
        # LOAD
        v_name="TLS_ST"
        source_data = self.get_main_data("TLS_FD")

        # EXTRACT desired data

        rows=[6]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_TTM_CP(self):
        # LOAD
        v_name="TTM_CP"
        source_data = self.get_main_data("TTM_FD")

        # EXTRACT desired data

        rows=[0,1]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_TTM_CG(self):
        # LOAD
        v_name="TTM_CG"
        source_data = self.get_main_data("TTM_FD")

        # EXTRACT desired data

        rows=[2]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_TTM_GFCF(self):
        # LOAD
        v_name="TTM_GFCF"
        source_data = self.get_main_data("TTM_FD")

        # EXTRACT desired data

        rows=[4,5]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_TTM_EXP(self):
        # LOAD
        v_name="TTM_EXP"
        source_data = self.get_main_data("TTM_FD")

        # EXTRACT desired data

        rows=[9,12]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_TTM_ST(self):
        # LOAD
        v_name="TTM_ST"
        source_data = self.get_main_data("TTM_FD")

        # EXTRACT desired data

        rows=[6]
        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand MG
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])

##-----------------------------------------------------
    def genr_r_TTM_CP(self):
        # LOAD
        v_name="r_TTM_CP"
        source_data = self.get_main_data("TTM_FD")
        source_data2 = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[0,1]
        target0 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand CPpp
        target0 = self.hlpr.IPTS_aggreg_cols(target0)


        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TTM
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target2 = self.hlpr.IPTS_div_elem(target1,target0)

        # WRITE to excel
        self.hlpr.write_simple(target2,varname=v_name,info=self.variable_dict[v_name])



##-----------------------------------------------------
    def genr_r_TTM_imed(self):
        # LOAD
        v_name="r_TTM_imed"
        source_data = self.get_main_data("TTM")
        source_data2 = self.get_main_data("USE_PP")

        # EXTRACT desired data

##        rows=[0,1]
##        target0 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand CPpp
##        target0 = self.hlpr.IPTS_aggreg_cols(target0)


##        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TTM
##        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target1 = self.hlpr.IPTS_div_elem(source_data,source_data2)
        target1 = self.hlpr.IPTS_make_rows_to_1column(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_r_TTM_CG(self):
        # LOAD
        v_name="r_TTM_CG"
        source_data = self.get_main_data("TTM_FD")
        source_data2 = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[2]
        target0 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand CGpp
        target0 = self.hlpr.IPTS_aggreg_cols(target0)


        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TTM
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target2 = self.hlpr.IPTS_div_elem(target1,target0)

        # WRITE to excel
        self.hlpr.write_simple(target2,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_r_TTM_GFCF(self):
        # LOAD
        v_name="r_TTM_GFCF"
        source_data = self.get_main_data("TTM_FD")
        source_data2 = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[4,5]
        target0 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand CPpp
        target0 = self.hlpr.IPTS_aggreg_cols(target0)


        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TTM
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target2 = self.hlpr.IPTS_div_elem(target1,target0)

        # WRITE to excel
        self.hlpr.write_simple(target2,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_r_TTM_EXP(self):
        # LOAD
        v_name="r_TTM_EXP"
        source_data = self.get_main_data("TTM_FD")
        source_data2 = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[9,12]
        target0 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand CPpp
        target0 = self.hlpr.IPTS_aggreg_cols(target0)


        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TTM
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target2 = self.hlpr.IPTS_div_elem(target1,target0)

        # WRITE to excel
        self.hlpr.write_simple(target2,varname=v_name,info=self.variable_dict[v_name])



##-----------------------------------------------------
    def genr_r_TTM_ST(self):
        # LOAD
        v_name="r_TTM_ST"
        source_data = self.get_main_data("TTM_FD")
        source_data2 = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[6]
        target0 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand CPpp
        target0 = self.hlpr.IPTS_aggreg_cols(target0)


        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TTM
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target2 = self.hlpr.IPTS_div_elem(target1,target0)

        # WRITE to excel
        self.hlpr.write_simple(target2,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_r_TLS_imed(self):
        # LOAD
        v_name="r_TLS_imed"
        source_data = self.get_main_data("TLS")
        source_data2 = self.get_main_data("USE_PP")

        # EXTRACT desired data

##        rows=[0,1]
##        target0 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand CPpp
##        target0 = self.hlpr.IPTS_aggreg_cols(target0)


##        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TTM
##        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target1 = self.hlpr.IPTS_div_elem(source_data,source_data2)
        target1 = self.hlpr.IPTS_make_rows_to_1column(target1)

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_r_TLS_CP(self):
        # LOAD
        v_name="r_TLS_CP"
        source_data = self.get_main_data("TLS_FD")
        source_data2 = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[0,1]
        target0 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand CPpp
        target0 = self.hlpr.IPTS_aggreg_cols(target0)


        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TLS
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target2 = self.hlpr.IPTS_div_elem(target1,target0)

        # WRITE to excel
        self.hlpr.write_simple(target2,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_r_TLS_CG(self):
        # LOAD
        v_name="r_TLS_CG"
        source_data = self.get_main_data("TLS_FD")
        source_data2 = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[2]
        target0 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand CGpp
        target0 = self.hlpr.IPTS_aggreg_cols(target0)


        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TLS
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target2 = self.hlpr.IPTS_div_elem(target1,target0)

        # WRITE to excel
        self.hlpr.write_simple(target2,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_r_TLS_GFCF(self):
        # LOAD
        v_name="r_TLS_GFCF"
        source_data = self.get_main_data("TLS_FD")
        source_data2 = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[4,5]
        target0 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand CPpp
        target0 = self.hlpr.IPTS_aggreg_cols(target0)


        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TLS
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target2 = self.hlpr.IPTS_div_elem(target1,target0)

        # WRITE to excel
        self.hlpr.write_simple(target2,varname=v_name,info=self.variable_dict[v_name])


##-----------------------------------------------------
    def genr_r_TLS_EXP(self):
        # LOAD
        v_name="r_TLS_EXP"
        source_data = self.get_main_data("TLS_FD")
        source_data2 = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[9,12]
        target0 = self.hlpr.IPTS_extract_cols(source_data2,rows)  # final demand CPpp
        target0 = self.hlpr.IPTS_aggreg_cols(target0)


        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TLS
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target2 = self.hlpr.IPTS_div_elem(target1,target0)

        # WRITE to excel
        self.hlpr.write_simple(target2,varname=v_name,info=self.variable_dict[v_name])



##-----------------------------------------------------
    def genr_r_TLS_ST(self):
        # LOAD
        v_name="r_TLS_ST"
        source_data = self.get_main_data("TLS_FD")
        source_data2 = self.get_main_data("USE_FD_PP")

        # EXTRACT desired data

        rows=[6]
        target0 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand CPpp
        target0 = self.hlpr.IPTS_aggreg_cols(target0)


        target1 = self.hlpr.IPTS_extract_cols(source_data,rows)  # final demand TLS
        target1 = self.hlpr.IPTS_aggreg_cols(target1)

        target2 = self.hlpr.IPTS_div_elem(target1,target0)

        # WRITE to excel
        self.hlpr.write_simple(target2,varname=v_name,info=self.variable_dict[v_name])



#   CPpp = CPbp / (1-r_tls-r_ttm)
#   CPbp = CPpp * (1-r_tls-r_ttm)
#   tls_cp   =  cp_pp01*r_tls_cp01
#   r_tls_cp =  tls_cp01/cp_pp01
#   r_ttm_cp =  ttm_cp01/cp_pp01

#------------------------GET FUNCTIONS-------------------------------

##
##    def test1d(self):
##        mtr=self.hlpr.gen_2d_matrix(10,10)
##
##        for i,row in enumerate(mtr):
##            for j,col in enumerate(row):
##                mtr[i][j]=rnd.randint(0,100)
##
##        print("vorher")
##        print mtr
##
##        print("nachher")
##        print self.hlpr.IPTS_make_rows_to_1column(mtr)

#-------------------------------------------------------
    def get_QN(self):  ##    QN   -> sum over rows
        # LOAD

        # LOAD
        v_name="QN"
        source_data = self.get_main_data("SUP_BP")

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_aggreg_rows(source_data)
        target1 = self.hlpr.IPTS_transpose(target1)

        return target1


    def get_TLS(self):  ##    QN   -> sum over rows
        # LOAD
        v_name="TLS"
        source_data = self.get_main_data("TLS")

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_aggreg_rows(source_data)
        target1 = self.hlpr.IPTS_transpose(target1)

        return target1


    def get_OPTX(self):  ## extract 1 row
        # LOAD
        v_name="OPTX"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name])

        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)

        return target1


    def get_K(self):  ## extract 1 row
        # LOAD
        v_name = "K"
        v_name1="SURPLUS_GR"
        v_name2="OPTX"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name1])
        va_row.append(self.va_detail[v_name2])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 2 rows - aggreg needed
        target1 = self.hlpr.IPTS_aggreg_rows(target1)
        target1 = self.hlpr.IPTS_transpose(target1)

        return target1


    def get_L(self):  ## extract 1 row
        # LOAD
        v_name="L"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row.append(self.va_detail[v_name])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)

        return target1


    def get_XD(self):  ## extract rows
        # LOAD
        v_name="XD"
        source_data = self.get_main_data("USEDOM_BP")


        # EXTRACT desired data - XD
        rows=[]
        for nen in self.sectors_nonengy:
            rows.append(self.sectors.index(nen))

        target1 = self.hlpr.IPTS_extract_rows(source_data,rows)  #extracting71 rows -  aggreg needed
        target1 = self.hlpr.IPTS_aggreg_rows(target1)  #aggreg to 1 row
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        return target1


    def get_XM(self):  ## extract rows
        # LOAD
        v_name="XM"
        source_data = self.get_main_data("USEIMP_BP")


        # EXTRACT desired data - XM
        rows=[]
        for nen in self.sectors_nonengy:
            rows.append(self.sectors.index(nen))

        target1 = self.hlpr.IPTS_extract_rows(source_data,rows)  #extracting71 rows -  aggreg needed
        target1 = self.hlpr.IPTS_aggreg_rows(target1)  #aggreg to 1 row
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        return target1

    def get_EN(self):  ## extract rows
        # LOAD
        v_name="EN"
        source_data = self.get_main_data("USE_BP")


        # EXTRACT desired data - EDN
        rows=[]

        for engy in self.sectors_engy:
            rows.append(self.sectors.index(engy))


        target1 = self.hlpr.IPTS_extract_rows(source_data,rows)  #extracting 3 rows -  aggreg needed
        target1 = self.hlpr.IPTS_aggreg_rows(target1)  #aggreg to 1 row
        target1 = self.hlpr.IPTS_transpose(target1)

        return target1

#--------------------------------------------------------------
#-------------------------------------------------------



#-------------------------------------------------------
    def get_EMN_detail(self):  ## extract rows
        # LOAD
        source_data = self.get_main_data("USEIMP_BP")
        v_name="EMN"

        # EXTRACT desired data - EMN
        rows=[]
        for en in self.sectors_engy:
            rows.append(self.sectors.index(en))

        target1 = self.hlpr.IPTS_extract_rows(source_data,rows)# extracting 3 rows - aggregation needed
        ##target1 = self.hlpr.IPTS_aggreg_rows(target1)  # aggregating
        ##target1 = self.hlpr.IPTS_transpose(target1)

        return target1




#-------------------------------------------------------
    def get_WAGES(self):  ## extract 1 row
        # LOAD
        v_name="WAGES"
        source_data = self.get_main_data("VA")

        va_row=[]
        va_row.append(self.va_detail[v_name])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)  #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)


        return target1


#-------------------------------------------------------
    def get_VA(self):  ## extract 1 row
        # LOAD
        v_name="VA"
        source_data = self.get_main_data("VA")


        # EXTRACT desired data
        # select row (start by 0)
        va_row=[]
        va_row=self.va_detail[v_name]


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row) #extracting 1 row - no aggreg needed
        target1 = self.hlpr.IPTS_transpose(target1)

        return target1


#----------------------------------------------------------------
    def gen_var_labels(self,name):
        #type A :  var01, var02, var03... var97
        #type B :  var01_01, var01_02, var01_05.... var97_97
        #type C : var01_01, var02_01, var05_01 .. var97_97


    ##  --------------- IN ARBEIT --------
        target=[]

        if vtype=="A":
            for i,sec in enumerate(sectors):
                target.append(name+sec)

        if vtype=="B":
            for i,sec in enumerate(sectors):
                for j,sec2 in enumerate(sectors):
                    target.append(name+sec+"_"+sec2)

        if vtype=="C":
            for i,sec in enumerate(sectors):
                for j,sec2 in enumerate(sectors):
                    target.append(name+sec2+"_"+sec)

        return target


#-------------------------------------------------------
#-------------------------------------------------------
#---------------- IPTS check ing----------
    def check_IPTS_SUT(self):
        hlpr.check_rows()
        pass

    # hier gehören noch funktionen die die Konsistenz von der SUT checken!!
    # umsetzung in helpr

##
##    def write_to_xlsx(self,name,data,info,is2dim):
##        #self.hlpr.write_to_xlsx(name,data,info,is2dim)
##        pass


##
##    def teststart(self,kk):
##        rr = getattr(self,'test'+'end')(kk)
##        print rr
##        print("huiuio")
##
##
##    def testend(self,nbr):
##        print(" i got %s " %nbr)
##        return 33
