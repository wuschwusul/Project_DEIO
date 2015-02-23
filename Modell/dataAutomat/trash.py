            ##timespan=["08","09","10"]
            ##timespan=["95","96","97","98","99","00","01","02","03","04","05","06","07"]


##sh_yrs=["95","96","97","98","99","00","01","02","03","04","05","06","07"]

##    def openSUT(self):
##
##
##
##        ws1 = self.wb.get_sheet_by_name(name = 'use08pp')
##        print ws1.cell('BH64').value
##
##        for row in ws1.iter_rows('D94:BY94'):
##            for cell in row:
##                print cell.value
##QNdataarea='D94:BY94'

    self.areas.append( self.merge_areas("K","SURPLUS","OPTX","Production Input - Capital Compensation")   )

    def merge_areas(self,name,area_name1,area_name2,title):  # ("K","SURPLUS","OPTX","Production Input - Capital Compensation")

            self.areas.append(name,)


            #continuous 1Dim area- name, sheet price  [area]    title
            self.areas.append(["QN","use","bp",["D96:BY96"],"Sector Gross Output"])
            self.areas.append(["TLS","tls","",["D82:BY82"],"Sector Taxes Less Subsidy"])
            self.areas.append(["VA","use","bp",["D95:BY95"],"Sector Value Added"])
            self.areas.append(["L","use","bp",["D89:BY89"],"Production Input - Labour Compensation"])
            self.areas.append(["WAGE","use","bp",["D90:BY90"],"Sector Wages"])
            self.areas.append(["SURPLUS","use","bp",["D94:BY94"],"Sector Surplus gross"])
            self.areas.append(["SURPLUS_net","use","bp",["D93:BY93"],"Sector Surplus net"])
            self.areas.append(["DEPR","use","bp",["D92:BY92"],"Sector Consumption of fixed Capital"])
            self.areas.append(["OPTX","use","bp",["D91:BY91"],"Other net taxes on production"])
            self.areas.append(["MN","useimp","bp",["D82:BY82"],"Sector Inputs - Imported"])
            self.areas.append(["DN","usedom","bp",["D82:BY82"],"Sector Inputs - Domestic"])
            self.areas.append(["QG","sup","bp",["BZ8:BZ81"],"Commodity Production"])
            self.areas.append(["MG","useimp","bp",["CP8:CP81"],"Commodity Imports"])
            #non-continuous 1Dim area
            self.areas.append(["EN","use","bp",["D11:by11","D21:by21","D36:by36"],"Sector Energy Commodity Inputs"])
            self.areas.append(["EMN","useimp","bp",["D11:by11","D21:by21","D36:by36"],"Sector Energy Commodity Inputs - Imported"])
            self.areas.append(["EDN","usedom","bp",["D11:by11","D21:by21","D36:by36"],"Sector Energy Commodity Inputs - Domestic"])

            #non-continuous 2Dim area


    def init_main_data(self):
        pass
        #main data (md) vectors [yr][row][col]
##        self.main_data["USE_PP"]= 0
##        md_use_cons_pp=0
##        md_use_gfcf_pp=0
##        md_use_exp_pp=0
##
##        md_use_bp= 0
##        md_use_cons_bp=0
##        md_use_gfcf_bp=0
##        md_use_exp_bp=0
##
##        md_usedom_cons=0
##        md_usedom_gfcf=0
##        md_usedom_exp=0
##
##        md_useimp_cons=0
##        md_useimp_gfcf=0
##        md_useimp_exp=0
##
##        md_sup_bp=0
##        md_sup_pp= 0		# sup bp and sup pp should be equal
##
##        md_ttm=0
##        md_ttm_cons=0
##        md_ttm_gfcf=0
##        md_ttm_exp=0
##
##        md_tls=0
##        md_tls_cons=0
##        md_tls_gfcf=0
##        md_tls_exp=0


    def get_variable_assignment(self,label):  # gets var_label (QN,EN...), returns assigend main area (USE_PP)
        for ma in self.main_assignment:
            if label in self.main_assignment[ma]:
                return ma

        if is_main_data_loaded(asgn)==False:    # i.e. if data (USE_BP) is NOT already loaded in main_data

##    def is_main_data_loaded(self,datalabel): # is (SUP_PP,USE_CONS_PP...) in Main_data{} ?
##        if datalabel in main_data:
##            return True
##        return False

##        print("test")
##        h=DeioHelper()
##        h.test()

##class WIOD:
##
##
##    def __init__(self):
##        pass
##
##
##
##class EuroStat:
##
##
##    def __init__(self):
##        pass

##        for data in dataarea:       # in case of several areas eg["D8:BY10","D11:BY81"]

##        "USE_CONS_PP":[ ["CP_PP"],"BS_FD_CONS"],
##        "USE_GFCF_PP":[ ["GFCF_PP"],"BS_FD_GFCF"],
##        "USE_EXP_PP":[ ["EXP_PP"],"BS_FD_EXP"],
##        "USE_CONS_BP":[ ["CP_BP"],"BS_FD_CONS"],
##        "USE_GFCF_BP":[ ["GFCF_BP"],"BS_FD_GFCF"],
##        "USE_EXP_BP":[ ["EXP_BP"],"BS_FD_EXP"],
##        "USEDOM_CONS_BP":[ [""],"BS_FD_CONS"],
##        "USEDOM_GFCF_BP":[ [""],"BS_FD_GFCF"],
##        "USEDOM_EXP_BP":[ [""],"BS_FD_EXP"],
##        "USEIMP_CONS_BP":[ [""],"BS_FD_CONS"],
##        "USEIMP_GFCF_BP":[ [""],"BS_FD_GFCF"],
##        "USEIMP_EXP_BP":[ [""],"BS_FD_EXP"],
##        "TTM_CONS":[ [""],"BS_FD_CONS"],
##        "TTM_GFCF":[ [""],"BS_FD_GFCF"],
##        "TTM_EXP":[ [""],"BS_FD_EXP"],
##        "TLS_CONS":[ [""],"BS_FD_CONS"],
##        "TLS_GFCF":[ [""],"BS_FD_GFCF"],
##        "TLS_EXP":[ [""],"BS_FD_EXP"],

##        for i,row in enumerate(ws1.iter_rows(data_range)):   # row
##
##            for j,cell in enumerate(row):                            #column
##                if firstround==True:
##                    target.append([])
##
##                target[i].append(cell.value)
##
##
##            firstround=False
##        return target


##    def get_array(self,su_cat,prc_cat,dataarea):  #eg. "sup","pp","A2:D3"
##    #returns a area as an array for each year [ [x,x,x,x],[x,x,x,x,x],[x,x,x,x,x] ]
##        print("Start extract data..."),
##        target=[]
##
##        for i,yr in enumerate(self.timespan):
##            target.append([])  # new column for each year
##
##            ws1 = self.wb.get_sheet_by_name(name = su_cat + yr + prc_cat) # eg. sup08pp
##            for data in dataarea:  # in case of several areas eg["D8:BY10","D11:BY81"]
##                for row in ws1.iter_rows(data):
##                    for cell in row:
##                        target[i].append(cell.value)
##
##        print("...OK")
##        return target

##    def get_variables(self):        # returns all assigned variables  ["QN","EN","QG"...
##        target=[]
##        for entry in self.main_tables:     # Entry = [ ["QN","EN"],"BS_SUP_USE"]
##            for varname in self.main_tables[entry][0]:    #varname = QN
##                if varname not in target:
##                    target.append(varname)
##
##        return target
##



##    def get_variables_assignment(self):             #  {"QN": "USE_PP", "QG": "SUP_PP"...
##        va={}
##
##        for v in self.get_variables():              # get_variables returns ["QN","EN","QG"...
##            va[v]=self.get_main_assignment(v)
##        return va


##    def get_main_assignment(self,var_label): # get "QN" -> "USE_PP"
##        for ma in self.main_tables:
##            if var_label in self.main_tables[ma][0]:
##                return ma
##
##        return False






##    def assign_base_areas(self,nace):
##        base_areas={}
##
##        # define MAIN AREAS:
##        if nace=="2008":
##            base_areas["BS_SUP_USE"]=["D8:BY81"]       # area of supply&use matrix
##            base_areas["BS_VA"]=["D87:BY91"]           # area of value added
##            base_areas["BS_FD_CONS"]=["CA8:CC81"]      #area of consumption
##            base_areas["BS_FD_GFCF"]=["CE8:CG81"]      #area of investent
##            base_areas["BS_FD_EXP"]=["CJ8:CM81"]       #area of export
##
##        if nace=="2003":
##            base_areas["BS_SUP_USE"]=["D8:BJ66"]       # area of supply&use matrix
##            base_areas["BS_VA"]=["D74:BJ77"]           # area of value added
##            base_areas["BS_FD_CONS"]=["BL8:BN66"]      #area of consumption
##            base_areas["BS_FD_GFCF"]=["BP8:BR66"]      #area of investent
##            base_areas["BS_FD_EXP"]=["BU8:BX66"]       #area of export

##        if nace==2003:                             #ASSIGN BASE
##            base_areas=cf.base2003
##        if nace==2008:
##            base_areas=cf.base2008
##    def get_nonengy_sectors(self):
##
##        sne=sectors      # create non-engy sectors
##        for es in sectors_engy:
##            sne.remove(es)
##        return sne



##    def init_sectors(self,nace):
##        if nace=="2008":
##            sectors = ['01', '02', '03', '05', '08', '10', '11', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31', '32', '33', '35', '36', '37', '41', '42', '43', '45', '46', '47', '49', '50', '51', '52', '53', '55', '58', '59', '60', '61', '62', '64', '65', '66', '68', '69', '70', '71', '72', '73', '74', '77', '78', '79', '80', '84', '85', '86', '87', '90', '91', '92', '93', '94', '95', '96', '97']
##            sectors_engy=['05','19','35']
##
##        if nace=="2003":
##            sectors = ["01","02","05","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31","32","33","34","35","36","37","40","41","45","50","51","52","55","60","61","62","63","64","65","66","67","70","71","72","73","74","75","80","85","90","91","92","93","95"]
##            sectors_engy=['10','11','23','40']
##
##        sectors_nonengy=get_nonengy_sectors()     # define non-energy sectors




##
##
##    def assign_variables(self):             #  {"QN": "USE_PP", "QG": "SUP_PP"...
##        for v in variables:
##            var_assignments[v]=get_main_assignment(v)



##    def get_main_assignment(self,var_label): # get "QN" -> "USE_PP"
##        for ma in main_assignments:
##            if var_label in main_assignments[ma][0]:
##                return ma
##
##        return False


##    def assign_main(self):    # assigning Main (USE_PP : QN,EN; USE_CONS_PP: ..
##
##        main_assignments["USE_PP"]=["QN","EN"],"BS_SUP_USE"
##        main_assignments["USE_CONS_PP"]=["..",".."],"BS_FD_CONS"
##        main_assignments["USE_GFCF_PP"]=["..",".."],"BS_FD_GFCF"
##        main_assignments["USE_EXP_PP"]=["..",".."],"BS_FD_EXP"
##
##        main_assignments["USE_BP"]=["..",".."],"BS_SUP_USE"
##        main_assignments["USE_CONS_BP"]=["..",".."],"BS_FD_CONS"
##        main_assignments["USE_GFCF_BP"]=["..",".."],"BS_FD_GFCF"
##        main_assignments["USE_EXP_BP"]=["..",".."],"BS_FD_EXP"
##
##        main_assignments["USEDOM_BP"]=["..",".."],"BS_SUP_USE"
##        main_assignments["USEDOM_CONS_BP"]=["..",".."],"BS_FD_CONS"
##        main_assignments["USEDOM_GFCF_BP"]=["..",".."],"BS_FD_GFCF"
##        main_assignments["USEDOM_EXP_BP"]=["..",".."],"BS_FD_EXP"
##
##        main_assignments["USEIMP_BP"]=["..",".."],"BS_SUP_USE"
##        main_assignments["USEIMP_CONS_BP"]=["..",".."],"BS_FD_CONS"
##        main_assignments["USEIMP_GFCF_BP"]=["..",".."],"BS_FD_GFCF"
##        main_assignments["USEIMP_EXP_BP"]=["..",".."],"BS_FD_EXP"
##
##        main_assignments["SUP_PP"]=["QG",".."],"BS_SUP_USE"
##        main_assignments["SUP_BP"]=["..",".."],"BS_SUP_USE"
##
##        main_assignments["TTM"]=["..",".."],"BS_SUP_USE"
##        main_assignments["TTM_CONS"]=["..",".."],"BS_FD_CONS"
##        main_assignments["TTM_GFCF"]=["..",".."],"BS_FD_GFCF"
##        main_assignments["TTM_EXP"]=["..",".."],"BS_FD_EXP"
##
##        main_assignments["TLS"]=["..",".."],"BS_SUP_USE"
##        main_assignments["TLS_CONS"]=["..",".."],"BS_FD_CONS"
##        main_assignments["TLS_GFCF"]=["..",".."],"BS_FD_GFCF"
##        main_assignments["TLS_EXP"]=["..",".."],"BS_FD_EXP"
##                "BS_FD_GFCF":"BP8:BR66",   #area of investent
##                "BS_FD_EXP":"BU8:BX66"     #area of export
##                "BS_FD_GFCF":"CE8:CG81",   #area of investent
##                "BS_FD_EXP":"CJ8:CM81"     #area of export

##    variables=0                             # ["QN","EN"...]

##    var_assignments={}                           #variables <--- MainArea, Area(s)
##                                            # e.g. "QG": [ ["SUP_PP"],["D8:BY81","D8:BY81"] ]... "QN":["USE_PP"]



##        #init variables to Tables
##        self.var_assignments=cf.get_variables_assignment()                         # assign MainAreas to VAriables

##        #get variables
##        self.variables=cf.get_variables()


##
##    def genr(self,label):       # Starts generation of Variable
##        # label = QN,QG...
##
##        # 1st generation (not combinated) = QN,QG,
##        # 2nd generation (combinated) = SEQ, SLQ,SED...

##        print label
##        print("....")
##
##        if label=="all":
##            self.load_all_main_data()
##            return True
##
####        asgn= self.var_assignments[label]        # QN -> USE_BP
####
####        if  asgn not in self.main_data:         # if main_data is NOT already loaded then
####            self.load_main_data(asgn)           # -> get data ("USE_BP" : [2021.12,0,0,0.12 ...]
##
##        #self.get("QN")
##        #helper.write(...)
##



##    def mainlabel_to_sheetname(self,label):      #USE_PP -> ["use","pp"];
##                                                #"USEIMP_CONS_BP" -> ["useimp","bp"]
##                                                # TLS -> ["tls"]
##        label=label.lower()
##        label=label.split("_")
##        if label.__len__()==1:
##            label.append("")
##            return label
##        else:
##            return [ label[0],label[label.__len__()-1] ]


##    def get_variables(self):
##        return self.variables



##        base_pos =self.main_assignments[main_label][1] # e.g. BS_FD_CONS


        if src not in self.main_tables_data:
            self.load_main_data(src)

        # GET main data
         =  self.main_tables_data[src]

##        target3 = list(target1)  # gen new
##        for yr,year in enumerate(target3):
##            for i,row in enumerate(year):
##                for j,cl in enumerate(row):
##                    t



#-------------------------------------------------------
    def genr_QN(self):  ##    QN   -> sum over rows
        # LOAD
        v_name="QN"
        source_data = self.get_main_data("USE_BP")

        # EXTRACT desired data
        target1 = self.get_QN()

        # WRITE to excel
        self.hlpr.write_simple(target1,varname=v_name,info=self.variable_dict[v_name])





    def get_QN(self):  ##    QN   -> sum over rows
        # LOAD
        v_name="QN"
        source_data = self.get_main_data("USE_BP")

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_sum_over_rows(source_data)
        target1 = self.hlpr.IPTS_transpose(target1)

        return target1

    def get_TLS(self):  ##    QN   -> sum over rows
        # LOAD
        v_name="TLS"
        source_data = self.get_main_data("TLS")

        # EXTRACT desired data
        target1 = self.hlpr.IPTS_sum_over_rows(source_data)
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        return target1
#-------------------------------------------------------
    def get_OPTX(self):  ## extract 1 row
        # LOAD
        v_name="OPTX"
        source_data = self.get_main_data("VA")

        va_row=[]
        va_row.append(self.va_detail[v_name])


        target1 = self.hlpr.IPTS_extract_rows(source_data,va_row)
        target1 = self.hlpr.IPTS_transpose(target1)

        # WRITE to excel
        return target1




arget3[yr][i][j]=target1[yr][i][j]-target2[yr][i][j]