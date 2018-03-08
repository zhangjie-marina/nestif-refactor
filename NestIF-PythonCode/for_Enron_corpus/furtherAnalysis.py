# -*- coding: utf-8 -*-
import ForNestedIf

#this script further analyzes the con equal formulas
def is_arithmetic(l):
    lnew = []
    for i in l:
        try:
            lnew.append(float(i))
        except:
            return False
    if len(l) < 2:
        return False
    l = lnew

    delta = l[1] - l[0]
    for index in range(len(l) - 1):
        if not (l[index + 1] - l[index] == delta):
             return False
    return True


def for_iserror(formula,dic):
    condition_list = []
    truevalue_list = []
    falsevalue_list = []
    truetype_list = []

    returnlist = []

    # print dic

    for eachlist in dic:

        depthone_list = dic[eachlist]
        count = 0
        condition_string = ''
        truevalue_string = ''
        falsevalue_string = ''

        for i in depthone_list:
            value = i.tvalue
            subtype = i.tsubtype
            if value is ',':
                count += 1
                continue
            if count == 0:
                condition_string += value
            elif count == 1:
                truevalue_string += value
                truetype_list.append(subtype)
            else:
                falsevalue_string += value

        condition_list.append(condition_string)
        truevalue_list.append(truevalue_string)
        falsevalue_list.append(falsevalue_string)
        if 'ISERROR' in truevalue_list:
            return True
        return False
        # print condition_list
def classify_equal_formulas(formula):
    ischoose = False
    isand = False
    isnewform = False
    p = ForNestedIf.ExcelParser()
    p.parse('='+formula)
    dic = p.get_dic_depth_token()
    returnlist = p.get_dic_for_equal(dic)
    dic_condition_value = returnlist[0]
    true_value_list = returnlist[3]

    try:
        num_value_list = []
        for item in true_value_list:
            num_value_list.append(float(item))
        print num_value_list
        if is_arithmetic(num_value_list):
            ischoose = True
            return 'ischoose'
        else:
            if len(set(true_value_list)) < len(true_value_list):
                isand = True
                return '  '
            else:
                return 'isnewform'

    except:
        if len(set(true_value_list)) < len(true_value_list):
            isand = True
            return 'isand'
        else:
            isnewform = True
            return 'isnewform'


def see_if_iserror():
    p = ForNestedIf.ExcelParser()

    filenamestring = 'CONEQUAL.txt'
    filepath = 'D:\\Users\\v-jizha4\\ExcelExp\\expAnalyResults\\totalsplits\\' + filenamestring
    readfile = open(filepath)

    lines = readfile.readlines()
    returnlist = []
    filenamelist = []
    for eachline in lines:
        eachline = eachline.strip()
        if 'D:\Users' in eachline:
            filenamelist.append(eachline)
            continue

        returnlist.append(eachline)
    analyzelist = returnlist
    count = 0
    for i in analyzelist:

        p.parse('='+i)
        dic = p.get_dic_depth_token()
        if for_iserror(i,dic):
            count += 1
            print count

def generate_classifyequal():
    p = ForNestedIf.ExcelParser()

    filenamestring = 'CONEQUAL.txt'
    filepath = 'D:\\Users\\v-jizha4\\ExcelExp\\expAnalyResults\\totalsplits\\' + filenamestring
    readfile = open(filepath)

    choose_filename = 'D:\\Users\\v-jizha4\\ExcelExp\\expAnalyResults\\totalsplits\\CHOOSE.txt'
    choose_write = open(choose_filename,'w')

    innerandfilename = 'D:\\Users\\v-jizha4\\ExcelExp\\expAnalyResults\\totalsplits\\innerand_new.txt'
    andwrite = open(innerandfilename, 'w')

    newform_filename = 'D:\\Users\\v-jizha4\\ExcelExp\\expAnalyResults\\totalsplits\\newform_new.txt'
    otherwrite = open(newform_filename, 'w')





    lines = readfile.readlines()
    returnlist = []
    filenamelist = []
    for eachline in lines:
        eachline = eachline.strip()
        if 'D:\Users' in eachline:
            filenamelist.append(eachline)
            continue

        returnlist.append(eachline)
    analyzelist = returnlist
    count = 0
    and_count = 0
    or_count = 0


    for i in analyzelist:
        count+=1
        p.parse('='+i)
        dic = p.get_dic_depth_token()
        result = classify_equal_formulas(i)
        if result == 'ischoose':
            print '-----------------------'
            print filenamelist[count-1]
            print i
            and_count+=1
            print "CHOOSE count: ",and_count
            print "total count: ",count

            choose_write.write(filenamelist[count-1]+'\n')
            choose_write.write(i+'\n')
        elif result == 'isand':
            print '-----------------------'
            print filenamelist[count - 1]
            print i
            or_count += 1
            print "Inner and count: ", or_count
            print "total count: ", count
            andwrite.write(filenamelist[count - 1] + '\n')
            andwrite.write(i + '\n')
        else:

            otherwrite.write(filenamelist[count - 1] + '\n')
            otherwrite.write(i + '\n')

    choose_write.close()
    andwrite.close()
    otherwrite.close()



if __name__ == '__main__':
    # inputs = [
    #     'IF(B4376 = 4, 2, IF(B4376 = 3, 2, IF(B4376 = 2, 1, 0)))',
    #     # 'IF(A2="Venezuela","ANDEAN",IF(A2="Ecuador","ANDEAN",IF(A2="Perú","ANDEAN",IF(A2="Colombia","ANDEAN",IF(A2="Uruguay","SOUTHCONE",IF(A2="Paraguay","SOUTHCONE",IF(A2="Chile","SOUTHCONE",IF(A2="Bolivia","SOUTHCONE",IF(A2="Argentina","SOUTHCONE",IF(A2="República Dominicana","CCA",IF(A2="Puerto Rico","CCA",IF(A2="San Vincente y Granadinas","CCA",IF(A2="Panamá","CCA",IF(A2="Nicaragua","CCA",IF(A2="Honduras","CCA",IF(A2="Guatemala","CCA",IF(A2="El Salvador","CCA",IF(A2="Costa Rica","CCA",IF(A2="México","MEXICO",IF(A2="Brasil","BRAZIL","OTRO"))))))))))))))))))))'
    # ]
    #
    #
    #
    # for i in inputs:
    #     print "========================================"
    #     print "Formula:     " + i
    #     print classify_equal_formulas(i)


    generate_classifyequal()
    # see_if_iserror()