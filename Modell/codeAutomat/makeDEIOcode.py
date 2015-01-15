# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      msommer
#
# Created:     12.01.2015
# Copyright:   (c) msommer 2015
# Licence:     <your licence>
#-------------------------------------------------------------------------------

nace08 = ['01', '02', '03', '05', '08', '10', '11', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31', '32', '33', '35', '36', '37', '41', '42', '43', '45', '46', '47', '49', '50', '51', '52', '53', '55', '58', '59', '60', '61', '62', '64', '65', '66', '68', '69', '70', '71', '72', '73', '74', '77', '78', '79', '80', '84', '85', '86', '87', '90', '91', '92', '93', '94', '95', '96', '97']
nace08_ne = ['01', '02', '03',  '08', '10', '11', '13', '14', '15', '16', '17', '18', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31', '32', '33', '36', '37', '41', '42', '43', '45', '46', '47', '49', '50', '51', '52', '53', '55', '58', '59', '60', '61', '62', '64', '65', '66', '68', '69', '70', '71', '72', '73', '74', '77', '78', '79', '80', '84', '85', '86', '87', '90', '91', '92', '93', '94', '95', '96', '97']
nace08_e= ['05', '19', '35']

def main():
################################
##  GENERAL DATA
#################################



    code = get_01_01()
    code += get_01_02()





# ++++ SAVING THE CODE ++++++++++
    f = open("K:\\Projekt_DEIO\\Modell\\Modellcode\\eviewscode.txt", "w")
    try:
        for lines in code:
            f.writelines(lines)
        f.close()
    except IOError:
        print("Error at file writing")

################# SEGMENT CODE GETTER #####################################

def get_01_01():
# +++ SEGMENT 1 +++
    code_1_01 = MatMul_A("QN","D","QG",False, True, True, nace08)  # returns Array of Code Lines
    code_1_01.insert(0,"'### ### ### ### ### ")
    code_1_01.insert(0,"'### SEGMENT  Q = f (D,QG) \n")
    code_1_01.insert(0,"\n\n'### ### ### ### ### \n")
    return code_1_01

def get_01_02():
# +++ SEGMENT 2 +++
    code_1_02 = MatMul_B("QG","sxd","sxdq","QN",False, True, True, True, nace08_ne)
    code_1_02 += MatMul_B("QG","sed","seq","QN",False, True, True, True, nace08_e)

    for i,line in enumerate(code_1_02):
        code_1_02[i]= line + " + FD"+nace08[i]

    code_1_02.insert(0,"'### ### ### ### ### ")
    code_1_02.insert(0,"'### SEGMENT  Q = f (D,QG) \n")
    code_1_02.insert(0,"\n\n'### ### ### ### ### \n")
    return code_1_02



###################### SUB FUNKTIONEN #################################

def MatMul_A(target, msqr, source, t1, t2, t3, secs):
# für  X(n) = Y(nxn) * Z(n)  in 9 possible various variants
# eg
#  QN01 = D01_01 * QG01 + D01_02 * QG02  = False,True,True  # true => change every round, false => do not change
#  QN01 = D01_01 * QG01 + D02_01 * QG01  = True,False,False
#  QN01 = D01_01 * QG01 + D01_01 * QG02  = False,False,True

    text= []


    for i in secs:
        line=""

        line+="\n\n@IDENTITY "+target+i+" = "        # QN01 =

        for j in secs:
            line+= msqr                   #QN01 = D
        ##########
            if t1==True:                  #QN01 = D01_
                line+=j+"_"
            else:
                line+=i+"_"
        ##########
            if t2==True:                  #QN01 = D01_01 *
                line+=j+" * "
            else:
                line+=i+" * "
        ##########
            if t3==True:                  #QN01 = D01_01 * QG01 +
                line+=source+j
            else:
                line+=source+i

            if j != secs[secs.__len__()-1]:
                line+=" + "

        text.append(line)
    return text


def MatMul_B(target, msqrA, msqrB, source, t1, t2, t3, t4, secs):
    # QG01 = sxd01_01 * sxdq01 * QN01 + sxd01_02 * sxdq02 * QN02

    text= []

    for i in secs:
        line=""

        line+="\n\n@IDENTITY "+target+i+" = "        # QG01 =

        for j in secs:
            line+= msqrA                   #Qg01 = sxd
        ##########
            if t1==True:                  #Qg01 = sxd01_
                line+=j+"_"
            else:
                line+=i+"_"
        ##########
            if t2==True:                  #Qg01 = sxd01_01 *
                line+=j+" * "
            else:
                line+=i+" * "
        ##########
            line+=msqrB                   #Qg01 = sxd01_01 * sxdq
            if t3==True:
                line+=j+" * "
            else:                          #Qg01 = sxd01_01 * sxdq01 *
                line+=i+" * "
        ##########
            if t4==True:                  #Qg01 = sxd01_01 * sxdq01 * QN01 +
                line+=source+j
            else:
                line+=source+i

            if j != secs[secs.__len__()-1]:
                line+=" + "

        text.append(line)
    return text

if __name__ == '__main__':
    main()
