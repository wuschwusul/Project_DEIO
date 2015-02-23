# -*- coding: utf-8 -*-



class Config:


    base2008={
                "BS_SUP_USE":"D8:BY81",    # area of supply&use matrix
                "BS_VA":"D89:BY95",       # area of value added !!! Position Varies- this is position of USE_BP!!
                "BS_FD":"CA8:CM81",    #area of consumption,investent,export

             }

    base2003={
                "BS_SUP_USE":"D8:BJ66",    # area of supply&use matrix
                "BS_VA":"D74:BJ79",       # area of value added  !!! Position Varies- this is position of USE_BP!!
                "BS_FD":"BL8:BX66",    #area of consumption

             }


    sectors_2008= ['01', '02', '03', '05', '08', '10', '11', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31', '32', '33', '35', '36', '37', '41', '42', '43', '45', '46', '47', '49', '50', '51', '52', '53', '55', '58', '59', '60', '61', '62', '64', '65', '66', '68', '69', '70', '71', '72', '73', '74', '77', '78', '79', '80', '84', '85', '86', '87', '90', '91', '92', '93', '94', '95', '96', '97']
    sectors_2003= ["01","02","05","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31","32","33","34","35","36","37","40","41","45","50","51","52","55","60","61","62","63","64","65","66","67","70","71","72","73","74","75","80","85","90","91","92","93","95"]
    sectors_engy_2008 = ['05','19','35']
    sectors_engy_2003=  ['10','11','23','40']



    va_detail_2008={
            "L": 0,     #Compensation of employees
            "WAGES": 1,  #Wages and Salaries
            "OPTX": 2,      #Other net taxes on production
            "DEPR": 3,      # Consumption of fixed capital
            "SURPLUS_NT": 4, #Operating surplus, net
            "SURPLUS_GR": 5, #Operating surplus, gross
            "VA": 6         #Value added at basic prices

    }

    va_detail_2003={
            "L": 0,     #Compensation of employees
            "OPTX": 1,      #Other net taxes on production
            "DEPR": 2,      # Consumption of fixed capital
            "SURPLUS_NT": 3, #Operating surplus, net
            "SURPLUS_GR": 4, #Operating surplus, gross
            "VA": 5         #Value added at basic prices
                }

    # Main Tables =SUP_BP,USE_FD_PP,USE_BP,USE_FD_BP,USEDOM_BP,USEIMP_BP,USEDOM_FD_BP,USEIMP_FD_BP,TTM,TLS,TTM_FD,TLS_FD,VA

    main_tables = {  #main_label: variables, base_position, sheetname,
        "SUP_BP":[ "BS_SUP_USE",["sup","bp"] ],         #["QG"],
        "USE_FD_PP":[ "BS_FD",["use","pp"]],        #["CP_PP","GFCF_PP","EXP_PP","ST_PP"],
        "USE_BP":["BS_SUP_USE",["use","bp"]],           # ["QN","EN"],
        "USE_FD_BP":[ "BS_FD",["use","bp"]],        #["CP_BP","GFCF_BP","EXP_BP","ST_BP"],
        "USEDOM_BP":["BS_SUP_USE",["usedom","bp"]],    # ["D_sub","EDN","XD",",SXD","SED"]                   # missing  SXDQ         SXDQ = XDxx / QNxx ; SXD = XDxx_yy/XDxx ..
        "USEIMP_BP":[ "BS_SUP_USE",["useimp","bp"]],        #["MN","MG","EMN","XM","SEM","SXM"],             # missing   SXMQ
        "USEDOM_FD_BP":["BS_FD",["usedom","bp"]],  #["FD","CPD","CGD","GFCFD","STD"],
        "USEIMP_FD_BP":["BS_FD",["useimp","bp"]],  #["FM","CPM","CGM","GFCFM","STM"],
        "TTM":["BS_SUP_USE",["ttm",""]],           #["TTM"],
        "TLS":["BS_SUP_USE",["tls",""]],           #["TLS"],
        "TTM_FD":["BS_FD",["ttm",""]],             #["TTM_FD"],
        "TLS_FD":["BS_FD",["tls",""]],             #["TLS_FD"],
        "VA":["BS_VA",["use","bp"]],                #["VA","SKQ","K","SLQ","L","PTX","SURPLUS_GR","DEPR","AGB","WAGES"],
        "USE_PP":["BS_SUP_USE",["use","pp"]],             # not necessary
        }


        #"SUP_PP":[ [""],"BS_SUP_USE"], # not necessary



    IPTS_variables= {                    # all below only in Basic Prices
        "QN"        : "Sector - Gross Output",                  # done
        "DN"        : "Sector - Domestic input  (XD+EDN = QN-MN-TLS)",  # done
        "XD"        : "Sector - Domestic input nonEnergy",# done
        "EDN"       : "Sector - Domestic input Energy ",# done
        "MN"        : "Sector - Imported input (XM+EMN)",# done
        "XM"        : "Sector - Imported input nonEnergy",# done
        "EMN"       : "Sector - Imported input Energy ",# done
        "TLS"       : "Sector - Taxes less subsidies",# done
        "VA"        : "Sector - Total Value Added",# done
        "L"         : "Value Added - Labour Compensation (AGB+WAGE)",# done
        "WAGES"     : "Value Added - Labour Wage",# done
        "AGB"       : "Value Added - Employer's SocialContribution of Labour Compensation (L-WAGES)", # done
        "K"         : "Value Added - Capital Compensation (Surplus_gr + OPTX)", # done
        "SURPLUS_GR": "Value Added - Gross Surplus of Capital ", # done
        "DEPR"      : "Value Added - Capital Depreciation ", # done
        "SURPLUS_NT": "Value Added - Net Surplus of Capital",# done
        "OPTX"      : "Value Added - Other Taxes on Production",# done
        "TLSQ"      : "Share TLS / QN ",# done
        "OPTXQ"     : "Share OPTX / QN ",# done
        "SKQ"       : "Share K / QN ",# done
        "SLQ"       : "Share L / QN ",# done
        "SXDQ"      : "Share XD / QN ",# done
        "SXMQ"      : "Share XM / QN ",# done
        "SEQ"       : "Share EN / QN ",# done
        "EN"        : "Sector - Domestic&Imported input Energy(EMN+EDN)", # done
        "MG"        : "Commodities - Imports total (MN+FM)", # done
        "QG"        : "Production - Gross Output Commodities, basic prices",# done
        "D"         : "Technology Matrix of Supply Table", # done
        "TTM_CP"    : "Transport Margins PrivCons",  # done
        "TTM_CG"    : "Transport Margins GovCons", # done
        "TTM_GFCF"  : "Transport Margins Investments", # done
        "TTM_EXP"   : "Transport Margins Exports", # done
        "TTM_ST"    : "Transport Margins Inventories", # done
        "r_TTM_imed"   : "Transport Margins Sectors in Detail",              # TTM_sub01_01 ... TTM_sub95_95
        "r_TTM_CP"    : "Transport Margins PrivCons", # done
        "r_TTM_CG"    : "Transport Margins GovCons", # done
        "r_TTM_GFCF"  : "Transport Margins Investments", # done
        "r_TTM_EXP"   : "Transport Margins Exports", # done
        "r_TTM_ST"    : "Transport Margins Inventories",# done
        "TLS_CP"    : "Taxes Less Subsidies PrivCons", # done
        "TLS_CG"    : "Taxes Less Subsidies GovCons", # done
        "TLS_GFCF"  : "Taxes Less Subsidies Investments", # done
        "TLS_EXP"   : "Taxes Less Subsidies Exports", # done
        "TLS_ST"    : "Taxes Less Subsidies Inventories", # done
        "r_TLS_imed"   : "Taxes Less Subsidies in Detail",
        "r_TLS_CP"    : "Share of Taxes Less Subsidies PrivCons",# done
        "r_TLS_CG"    : "Share of Taxes Less Subsidies GovCons",# done
        "r_TLS_GFCF"  : "Share of Taxes Less Subsidies Investments",# done
        "r_TLS_EXP"   : "Share of Taxes Less Subsidies Exports",# done
        "r_TLS_ST"    : "Share of Taxes Less Subsidies Inventories",# done
        "SXD"       : "Shares of Inputs in Domestic Inputs (share of XD)",  #sxd01_01 (row,col) ..sxd95_95  # done
        "SXM"       : "Shares of Inputs in Imported Inputs (share of XM)", # done
        "SED"       : "Shares of dom. Inputs in Energy Inputs (share of EN)",# done
        "SEM"       : "Shares of imp. Inputs in Energy Inputs (share of EN)",# done
        "F"         : "Final Demand - Total, basic prices", # done
        "FD"        : "Final Demand - Domestic Total (F-FM), basic prices", # done
        "FM"        : "Final Demand - Imported Total, basic prices", # done
        "CP_bp"     : "Final Demand - Priv.Consumption, basic prices",# done
        "CP_pp"     : "Final Demand - Priv.Consumption, purchaser prices",# done
        "CG_bp"     : "Final Demand - Priv.Consumption, basic prices",# done
        "CG_pp"     : "Final Demand - Priv.Consumption, purchaser prices",# done
        "EXP_bp"    : "Final Demand - Exports, basic prices",# done
        "EXP_bp"    : "Final Demand - Exports, purchaser prices",# done
        "GFCF_bp"   : "Final Demand - Gross Fixed Capital Formation, basic prices",# done
        "GFCF_pp"   : "Final Demand - Gross Fixed Capital Formation, purchaser prices",# done
        "ST_bp"     : "Final Demand - Inventory, basic prices",# done
        "ST_pp"     : "Final Demand - Inventory, purchaser prices",# done
        "mcp"       : "Final Demand - Import Share Priv.Consumption, basic prices",# done
        "mcg"       : "Final Demand - Import Share Gov.Consumption, basic prices",# done
        "mgfcf"     : "Final Demand - Import Share GFCF, basic prices",# done
        "mst"       : "Final Demand - Import Share Inventory, basic prices",# done
        "mex"       : "Final Demand - Import Share Exports, basic prices",# done
        #Fehlt : Preise (PQ,PE,PMN,PMG,PG,PD,PK,PL)
        # ZU FINDEN AUF EUROSTAT : [sts_inppd_m]
        # sowie heimische Produktionsindex: [sts_inpr_q]

    }

#------------------------------------------------



    def get_base(self,nace):
        if nace=="2008":
            return self.base2008

        if nace=="2003":
            return self.base2003


    def get_main_tables_info(self):
        return self.main_tables





    def get_engy_sectors(self,nace):
        if nace=="2008":
            return self.sectors_engy_2008
        if nace=="2003":
            return self.sectors_engy_2003


    def get_sectors(self,nace):
        if nace=="2008":
            return self.sectors_2008

        if nace=="2003":
            return self.sectors_2003



    def get_nonengy_sectors(self,nace):# create non-engy sectors

        if nace=="2008":
            sne=list(self.sectors_2008)
            for es in self.sectors_engy_2008:
                sne.remove(es)
            return sne

        if nace=="2003":
            sne=list(self.sectors_2003)
            for es in self.sectors_engy_2003:
                sne.remove(es)
            return sne

        return False

    def get_IPTS_variable_names(self):
        return self.IPTS_variables

    def get_IPTS_variable_dict(self):
        return self.IPTS_variables   #.copy()

    def get_va_detail(self,nace):
        if nace=="2003":
            return self.va_detail_2003
        if nace=="2008":
            return self.va_detail_2008
