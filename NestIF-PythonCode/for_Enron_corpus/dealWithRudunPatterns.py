# -*- coding: utf-8 -*-
import ForNestedIf
# import get_refactor_fullresults_LOOP_removeredun


#get dic[condition] = [truevalue, falsevalue
def get_condition_values_dic(formula, dic):
    p = ForNestedIf.ExcelParser()
    try:
        p.parse('='+formula)

        if p.tokens.items[0].tvalue != 'IF':
            p.parse('='+p.get_onlyIFfunction())
        threeparts = p.get_threeparts_IF()

        if 'condition_list' in dic:
            dic['condition_list'].append(threeparts[0])
        else:
            dic['condition_list'] = [threeparts[0]]


        if threeparts[0] in dic:

            dic[threeparts[0]].append(threeparts[1])
        else:
            dic[threeparts[0]] = [threeparts[1]]
        if threeparts[0] in dic:

            dic[threeparts[0]].append(threeparts[2])
        else:
            dic[threeparts[0]] = [threeparts[2]]


        if 'IF' in formula:
            for each in threeparts:
                if 'IF' in each:
                    p.parse('='+each)
                    new = p.get_onlyIFfunction()

                    get_condition_values_dic(new,dic)
    except:
        dic = {}


def redundancy_loop_NEW(formula):
    formula = formula.replace(' ', '')
    containredun = True
    while(containredun):
        containredun = False
        para_if_list = ForNestedIf.get_para_iflist_second(formula)
        if not para_if_list:
            para_if_list = []
        p = ForNestedIf.ExcelParser()
        try:
            p.parse('=' + formula)
        except:
            return False
        threeparts_list = p.get_threeparts_IF()
        for eachlist in threeparts_list:
            innerpara_if_list = ForNestedIf.get_para_iflist_second(eachlist)
            if innerpara_if_list:
                for eachinner in innerpara_if_list:
                    para_if_list.append(eachinner)
        if not para_if_list or len(para_if_list) == 0:
            return formula

        for each in para_if_list:
            dic = {}

            get_condition_values_dic(formula, dic)
            if dic == '':
                return False

            returnlist = deal_with_redundancy_oneround(each, dic)
            # print 'returnlist: ', returnlist

            if returnlist:

                if each in formula:
                    containredun = True
                    formula = formula.replace(each,returnlist[1])
                    p = ForNestedIf.ExcelParser()


    return formula


def redundancy_loop(formula, stopboolean,canrefactorboolean):
    formula = formula.replace(' ','')
    if stopboolean:

        para_if_list =ForNestedIf.get_para_iflist_second(formula)


        if not para_if_list:
            stopboolean = True
            if canrefactorboolean:
                return formula
            else:
                return False

        for each in para_if_list:
            dic = {}

            get_condition_values_dic(formula, dic)
            if dic == '':
                return False


            returnlist = deal_with_redundancy_oneround(formula, dic)





            if returnlist:
                stopboolean = False
                if each in formula:
                    canrefactorboolean = True

                    formula = formula.replace(each,returnlist[1])



        p = ForNestedIf.ExcelParser()
        p.parse('='+formula)
        threeparts_list = p.get_threeparts_IF()


        for eachthree in threeparts_list:


            dic = {}

            get_condition_values_dic(eachthree, dic)
            if dic == '':
                return False

            returnresultspart = deal_with_redundancy_oneround(eachthree, dic)

            if returnresultspart:
                stopboolean = False
                if eachthree in formula:
                    canrefactorboolean = True

                    formula = formula.replace(eachthree, returnresultspart[1])

        if canrefactorboolean:
            redundancy_loop(formula, stopboolean,canrefactorboolean)
    if canrefactorboolean:
        return formula
    else:
        return False

def deal_with_redundancy_loop(formula):

    dic = {}
    before = False
    lastreturnlist = ''
    get_condition_values_dic(formula,dic)
    if dic == '':
        return False


    returnlist = deal_with_redundancy_oneround(formula,dic)




    while (returnlist):
        before = True
        lastreturnlist = returnlist
        dic = {}
        get_condition_values_dic(returnlist[1], dic)

        returnlist = deal_with_redundancy_oneround(returnlist[1], dic)
        print 'return list every time: ', returnlist



        if not returnlist and before:
            return lastreturnlist

    return returnlist

def deal_with_redundancy_oneround(formula,dic):

    originalfomula = formula

    contain = False
    condition_list = []
    returnlist = []
    p = ForNestedIf.ExcelParser()
    # print '--------'
    if 'condition_list' not in dic:
        return False
    condition_list = dic['condition_list']


    returnlist.append(formula)

    oppo_list = []
    # dic_oper_oppo = {}
    # dic_oper_oppo['='] = ['!=', '<>']
    # dic_oper_oppo['>='] = ['<']
    # dic_oper_oppo['<='] = ['>']
    # dic_oper_oppo['>'] = ['<=']
    # dic_oper_oppo['<'] = ['>=']
    # dic_oper_oppo['!='] = ['=']
    # dic_oper_oppo['<>'] = ['=']




    for key in dic:
        if 'condition_list' in key:
            continue
        if len(dic[key]) > 2:
            ccc = 0

            while ccc<(len(dic[key])/2):
                ccc+=2

                redun = 'IF('+key+','+dic[key][ccc]+','+dic[key][ccc+1]+')'
                if ',)' in redun:
                    redun = redun.replace(',)', ')')

                if redun in dic[key][0]:

                    # print '1'
                    contain = True
                    newstring = dic[key][0].replace(redun,dic[key][ccc])
                    formula=formula.replace(dic[key][0],newstring)
                elif redun in dic[key][1]:
                    # print '2'
                    contain = True
                    newstring = dic[key][1].replace(redun,dic[key][ccc+1])
                    formula = formula.replace(dic[key][1],newstring)


        if 'IF' not in key and (key != ''):
            newkey = get_oppo_logicformula(key)
            oppo_list.append(newkey)

            # if p.tokens.items[1].tvalue != '' and p.tokens.items[0].ttype != 'function':
            #     for i in dic_oper_oppo[p.tokens.items[1].tvalue]:
            #         if p.tokens.items[2].tvalue == '':
            #             oppo_list.append(p.tokens.items[0].tvalue + i + '\"\"')
            #         else:
            #             substring = ''
            #             cccc = 0
            #             while(cccc < len(p.tokens.items)-2):
            #                 substring += p.tokens.items[cccc+2].tvalue
            #                 cccc += 1
            #             oppo_list.append(p.tokens.items[0].tvalue + i + substring)
    if len(set(condition_list).intersection(oppo_list)) != 0:
        # print '3'

        for i in oppo_list:
            if i in condition_list:
                jone = ''

                jone = get_oppo_logicformula(i)
                if jone not in dic:
                    continue
                redun = 'IF(' + jone + ',' + dic[jone][0] + ',' + dic[jone][1] + ')'


                if ',)' in redun:
                    redun = redun.replace(',)',')')


                if redun in dic[i][0]:
                    newstring = dic[i][0].replace(redun,dic[jone][1])
                    formula = formula.replace(dic[i][0], newstring)
                    contain = True
                elif redun in dic[i][1]:
                    newstring = dic[i][1].replace(redun,dic[jone][0])
                    formula = formula.replace(dic[i][1], newstring)
                    contain = True
            # if contain:
            #     returnlist.append(formula)
                # return returnlist
    isredun = False
    TFlist = []
    # deal with condition combination
    each = ''
    andboolean = True

    for eachcondition_list in condition_list:
        conditionresults = deal_with_AND_condition(eachcondition_list)
        if conditionresults:
            for each in conditionresults:
                if conditionresults[each] not in condition_list:
                    andboolean = False

            if andboolean:

                isredun =  True
                each = eachcondition_list

                break


    if not isredun:
        if contain and formula != originalfomula:
            returnlist.append(formula)
            if len(returnlist) == 2:
                return returnlist
        else:
            return False




    for eachconditionresults in conditionresults:

        eachconditionresults = conditionresults[eachconditionresults]
        oppo_eachconditionresults = get_oppo_logicformula(eachconditionresults)



        if eachconditionresults in dic:


            if each in dic[eachconditionresults][0]:
                TFlist.append('T')
            elif each in dic[eachconditionresults][1]:

                TFlist.append('F')
            elif eachconditionresults in dic[each][0]:

                redun = 'IF(' + eachconditionresults + ',' + dic[eachconditionresults][0] + ',' + dic[eachconditionresults][1] + ')'
                newstring = dic[each][0].replace(redun,dic[eachconditionresults][0])
                formula = formula.replace(dic[each][0],newstring)
                if formula != originalfomula:
                    returnlist.append(formula)
                    if len(returnlist) == 2:

                        return returnlist


        elif oppo_eachconditionresults != '':
            if oppo_eachconditionresults in dic:
                if each in dic[oppo_eachconditionresults][0]:
                    TFlist.append('F')
                elif each in dic[oppo_eachconditionresults][1]:
                    TFlist.append('T')
                elif oppo_eachconditionresults in dic[each][0]:
                    redun = 'IF(' + oppo_eachconditionresults + ',' + dic[oppo_eachconditionresults][0] + ',' + dic[oppo_eachconditionresults][1] + ')'
                    newstring = dic[each][0].replace(redun,dic[oppo_eachconditionresults][0])
                    formula = formula.replace(dic[each][0],newstring)
                    if formula != originalfomula:
                        returnlist.append(formula)
                        if len(returnlist) == 2:
                            return returnlist

            else:
                return False


        redun = 'IF(' + each + ',' + dic[each][0] + ',' + dic[each][1] + ')'
        if ',)' in redun:
            redun = redun.replace(',)',')')

        if 'F' in TFlist:
            contain = True
            formula = formula.replace(redun,dic[each][1])
            if formula != originalfomula:
                returnlist.append(formula)
        elif 'T' in TFlist:
            contain = True
            formula = formula.replace(redun, dic[each][0])
            if formula != originalfomula:
                returnlist.append(formula)

        if returnlist and len(returnlist) == 2:

            return returnlist
    return False




def get_oppo_logicformula(i):
    jone = ''
    if '<=' in i:
        jone = i.replace('<=', '>')
    elif '>=' in i:
        jone = i.replace('>=', '<')
    elif '<>' in i:
        jone = i.replace('<>', '=')
    elif '>' in i:
        jone = i.replace('>', '<=')
    elif '<' in i:
        jone = i.replace('<', '>=')
    elif '=' in i:
        jone = i.replace('=', '<>')
    return jone

def deal_with_AND_condition(formula):
    p = ForNestedIf.ExcelParser()
    try:
        p.parse('='+formula)

    except:
        return False
    stringdic = {}
    indent = 0
    count = 0
    stringdic[count] = ''
    isAndPattern = False # contain AND or &
    total_count = 0
    if p.tokens:

        if not p.tokens.items:
            return False
        if p.tokens.items[0].tvalue == 'AND':
            isAndPattern = True
        elif '&' in formula:
            isAndPattern = True
        if not isAndPattern:
            return False

        for t in p.tokens.items:
            isrange_anothersheet = False
            newvalue = ''

            if '!' in t.tvalue and t.tsubtype == 'range':
                newvalue = '\'' + t.tvalue.split('!')[0] + '\'' + '!' + t.tvalue.split('!')[1]
                isrange_anothersheet = True

            if (t.tsubtype == p.TOK_SUBTYPE_STOP):
                indent -= 1

            if (indent == 1 and t.tvalue == ',') or (indent == 0 and t.tvalue == '&'):
                if 'AND' in stringdic[count]:
                    stringdic[count] = stringdic[count].replace('AND','')
                count += 1
                stringdic[count] = ''
                continue


            else:
                if isrange_anothersheet:
                    stringdic[count] += newvalue

                elif t.tsubtype == 'text':
                    stringdic[count] += '\"' + t.tvalue + '\"'
                elif t.tsubtype == 'number' and t.tvalue == '':
                    stringdic[count] += '\"\"'
                else:
                    stringdic[count] += t.tvalue

            if (t.tsubtype == p.TOK_SUBTYPE_START):
                indent += 1;

    return stringdic


def whole_remove_redundancy(filenamestring):
    filepath = 'D:\\Users\\v-jizha4\\ExcelExp\\expAnalyResults\\' + filenamestring
    readfile = open(filepath)

    innerandfilename = 'D:\\Users\\v-jizha4\\ExcelExp\\expAnalyResults\\formal_results\\removed-redun' + filenamestring
    andwrite = open(innerandfilename, 'w')

    remainfilename = 'D:\\Users\\v-jizha4\\ExcelExp\\expAnalyResults\\formal_results\\without-redun-' + filenamestring
    otherwrite = open(remainfilename, 'w')

    filenamelist = []
    count = 0
    rcount = 0

    while (True):
        eachline = readfile.readline()

        if not eachline:
            break

        eachline = eachline.strip()
        if 'D:\Users' in eachline:
            filenamelist.append(eachline)
            continue

        i = eachline

        if i == '':
            continue


        count += 1
        if count < 5601849:
            continue
        # print count


        if count % 10000 == 0:
            print count


        returnlist =  deal_with_redundancy_loop(i)
        if returnlist:
            rcount += 1
            andwrite.write(returnlist[0]+'\n')
            andwrite.write(returnlist[1] + '\n')
            print 'number of redundant formula: ', rcount
        else:
            otherwrite.write(i + '\n')




    readfile.close()
    andwrite.close()
    otherwrite.close()

if __name__ == '__main__':

    # print deal_with_AND_condition('IF(AND(B10>=31,B10<60),"31-60",IF(AND(B10>=61,B10<90),"61-90","90+"))')
    # p = ForNestedIf.ExcelParser()
    # p.parse('IF(IF(S1215=Z1215,S1215,IF(S1215="",Z1215,S1215))="","",IF(S1215=Z1215,S1215,IF(S1215="",Z1215,S1215)))')
    # print p.get_threeparts_IF()
    # print  p.prettyprint()
    print redundancy_loop_NEW('IF(F3="A",\'PriceListData\'!E20,(IF(F3="B",\'PriceListData\'!F20,(IF(F3="C",\'PriceListData\'!G20,(IF(F3="D",\'PriceListData\'!H20)))))))')





    # whole_remove_redundancy('NestedIf


