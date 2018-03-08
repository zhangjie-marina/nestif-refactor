# -*- coding: utf-8 -*-
# from for_original_corpus import ForNestedIf, dealWithRudunPatterns, furtherAnalysis
import getReducedFormulas
import ForNestedIf
import dealWithRudunPatterns
import furtherAnalysis


def return_true_ifis_useless(formula):
    p = ForNestedIf.ExcelParser()
    try:
        p.parse('=' + formula)
    except:
        return False
    all_innerif_list = p.get_all_innerif_list()
    all_innerif_list.append(formula)
    if len(all_innerif_list) == 0:
        return False

    for eachone in all_innerif_list:

        try:
            p.parse('=' + eachone)
        except:
            return False

        threeparts = p.get_threeparts_IF()

        if (threeparts[1] + '=' + threeparts[2]) == threeparts[0]:
            return True
        elif (threeparts[2] + '=' + threeparts[1]) == threeparts[0]:
            return True

    return False


def return_true_ifis_IFS(formula, threepartlist):
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]
    if len(true_part_list) == 1:
        return False
    for each in true_part_list:
        if 'IF' in each:
            return False
    for each in false_part_list[:-1]:
        if 'IF' not in each:
            return False
    return True


def return_true_ifis_MAXMIN(formula):
    p = ForNestedIf.ExcelParser()
    try:
        p.parse('=' + formula)
    except:
        return False
    all_innerif_list = p.get_all_innerif_list()
    all_innerif_list.append(formula)
    if len(all_innerif_list) == 0:
        return False

    for eachone in all_innerif_list:

        if eachone.count('IF') != 1:
            continue

        threeparts = p.get_threeparts_IF()

        for eachsymbol in ['<', '>', '<=', '>=']:

            if (threeparts[1] + eachsymbol + threeparts[2]) == threeparts[0] or (
                                    threeparts[1] + eachsymbol + '(' + threeparts[2] + ')') == threeparts[0] or (
                                    '(' + threeparts[1] + ')' + eachsymbol + threeparts[2]) == threeparts[0] or (
                                            '(' + threeparts[1] + ')' + eachsymbol + '(' + threeparts[2] + ')') == \
                    threeparts[0]:
                return True
            if (threeparts[2] + eachsymbol + threeparts[1]) == threeparts[0] or (
                                    threeparts[2] + eachsymbol + '(' + threeparts[1] + ')') == threeparts[0] or (
                                    '(' + threeparts[2] + ')' + eachsymbol + threeparts[1]) == threeparts[0] or (
                                            '(' + threeparts[2] + ')' + eachsymbol + '(' + threeparts[1] + ')') == \
                    threeparts[0]:
                return True

    return False


def return_true_ifis_otherOR(formula, threepartlist):
    true_part_list = threepartlist[1]
    if not (len(set(true_part_list)) < len(true_part_list) and len(true_part_list) > 1):
        return False
    dic = {}
    sametruevaluelist = []
    samevaluestring = ''
    dealWithRudunPatterns.get_condition_values_dic(formula, dic)
    for eachkey in dic:
        if 'condition_list' in eachkey:
            continue
        truevalue = dic[eachkey][0]
        falsevalue = dic[eachkey][1]
        if true_part_list.count(truevalue) >1:
            samevaluestring = truevalue
            sametruevaluelist.append(eachkey+'::'+truevalue+'::'+falsevalue)
        if falsevalue == samevaluestring:
            sametruevaluelist.append('!'+eachkey + '::' + falsevalue + '::' + truevalue)
            break
    count = 0
    true = ''
    false = ''
    try:
        finalnewformula = 'IF(OR('+sametruevaluelist[0].split('::')[0]+','
    except:
        return False
    while count < len(sametruevaluelist)-1:
        count += 1
        thistring = sametruevaluelist[count]
        key, true, false = thistring.split('::')

        if not (true in sametruevaluelist[count-1].split('::')[1]):
            return False
        finalnewformula+=key+','
    finalnewformula = finalnewformula[:-1]
    finalnewformula += '),'+true+','+false+')'







def return_true_ifis_AND(formula, threepartlist):
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]
    if not (len(set(false_part_list)) == 1 and len(false_part_list) > 1):
        return False
    for each in true_part_list[:-1]:
        if 'IF' not in each:
            return False
    return True


def return_true_ifis_OR(formula, threepartlist):
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]


    if not (len(set(true_part_list)) == 1 and len(true_part_list) > 1):
        return False
    for each in false_part_list[:-1]:
        if 'IF' not in each:
            return False
    return True


def return_list_ifis_equal(formula, threepartlist):
    # order: choose, match, lookup, idlookup,
    dic_choose_match_lookup_idlookup = {}
    dic_choose_match_lookup_idlookup['choose'] = False
    dic_choose_match_lookup_idlookup['match'] = False
    dic_choose_match_lookup_idlookup['lookup'] = False
    dic_choose_match_lookup_idlookup['idlookup'] = False

    p = ForNestedIf.ExcelParser()
    condition_part_list = threepartlist[0]
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]
    before_equal_list = []
    after_equal_list = []
    after_equal_type_list = []
    true_part_type_list = []


    for each in condition_part_list:
        if '=' not in each or '<=' in each or '>=' in each:
            return False
        dic_choose_match_lookup_idlookup['idlookup'] = True
        before = each.split('=')[0]
        after = each.split('=')[1]

        try:
            p.parse('=' + after)
        except:
            # print 'exception: cannot parse after equal part: ' + formula
            return False
        after_type = p.tokens.items[0].tsubtype
        after_equal_type_list.append(after_type)

        after_equal_list.append(after)
        before_equal_list.append(before)

        # before equal has to be the same
    if not len(set(before_equal_list)) == 1:
        return False

    for each in true_part_list:
        try:
            p.parse('=' + each)
            truepart_type = p.tokens.items[0].tsubtype
            true_part_type_list.append(truepart_type)
        except:
            # print 'exception: cannot parse truevalue part: ' + formula
            return False

    # true value type and after equal type need to be the same
    if not len(set(true_part_type_list)) == 1 or (not len(set(after_equal_type_list)) == 1):
        return dic_choose_match_lookup_idlookup  # can only be indirect vlookup
    if set(after_equal_type_list) == set(['range']):
        if set(true_part_type_list) == set(['number']) and furtherAnalysis.is_arithmetic(true_part_list):
            dic_choose_match_lookup_idlookup['match'] = True
            return dic_choose_match_lookup_idlookup
        elif set(true_part_type_list) == set(['range']):
            dic_choose_match_lookup_idlookup['lookup'] = True
            return dic_choose_match_lookup_idlookup
        elif set(true_part_type_list) == set(['text']):
            return dic_choose_match_lookup_idlookup
    if set(after_equal_type_list) == set(['number']) and furtherAnalysis.is_arithmetic(after_equal_list):
        if set(true_part_type_list) == set(['number']) and furtherAnalysis.is_arithmetic(true_part_list):
            dic_choose_match_lookup_idlookup['match'] = True
            return dic_choose_match_lookup_idlookup
        elif set(true_part_type_list) == set(['range']) or (set(true_part_type_list) == set(['text'])):
            dic_choose_match_lookup_idlookup['choose'] = True
            return dic_choose_match_lookup_idlookup
    if set(after_equal_type_list) == set(['text']):
        if set(true_part_type_list) == set(['number']) and furtherAnalysis.is_arithmetic(true_part_list):
            dic_choose_match_lookup_idlookup['match'] = True
            return dic_choose_match_lookup_idlookup

    return dic_choose_match_lookup_idlookup


def get_para_iflist(formula):
    # input: the formula that you want to refactor
    # output: a list of inner if parts. They do not contain each other. Return false if the formula is not nested if formula
    p = ForNestedIf.ExcelParser()
    try:
        p.parse('=' + formula)
    except:
        return False
    ifreturnfalse = True

    returnlist = p.get_para_if()


    for each in returnlist:

        if (each.count('IF') - each.count('IFS')) > 1 and (each.count('IF') - each.count('IFERROR')) > 1 and (
            each.count('IF') - each.count('IFNA')) > 1:
            ifreturnfalse = False
    if not ifreturnfalse:
        return returnlist
    else:
        return False



def refactor_loop_NEW(formula,dic):

    if formula:
        formula = formula.replace(' ', '')
    else:
        return False
    canrefactor = True
    while canrefactor:
        canrefactor = False
        para_if_list = ForNestedIf.get_para_iflist_second(formula)
        if not para_if_list:
            para_if_list = []

        p = ForNestedIf.ExcelParser()
        p.parse('=' + formula)
        threeparts_list = p.get_threeparts_IF()
        for eachlist in threeparts_list:
            innerpara_if_list = ForNestedIf.get_para_iflist_second(eachlist)
            if innerpara_if_list:
                for eachinner in innerpara_if_list:

                    para_if_list.append(eachinner)

        if not para_if_list or len(para_if_list) == 0:
            return formula

        for each in para_if_list:
            returnresults = pattern_match_oneround(each, dic)




            if returnresults and each in formula:
                canrefactor = True
                formula = formula.replace(each, returnresults)
                p.parse('='+formula)
                newlen = p.get_nested_ifs()
                if newlen <2:
                    break
    return formula


def predealwith_formula(formula):
    formula = formula.strip().replace(' ', '')
    if '[#ThisRow],' in formula:
        formula = formula.replace('[#ThisRow],', '')
    p = ForNestedIf.ExcelParser()
    try:
        p.parse('='+formula)
    except:
        return False
    # if '+' not in formula:
    #     return formula
    else:
        indent = 0
        returnstring = ""
        stopindent = 0
        stop = False
        if p.tokens:
            while (p.tokens.moveNext()):
                token = p.tokens.current();

                t = token

                if (t.tsubtype == p.TOK_SUBTYPE_STOP):
                    indent -= 1
                    if indent == (stopindent - 1):
                        returnstring += t.tvalue + ')'
                if (t.tsubtype == p.TOK_SUBTYPE_START):

                    returnstring += t.tvalue + '('
                elif (t.tsubtype == p.TOK_SUBTYPE_STOP):

                    returnstring += t.tvalue + ')'
                else:
                    if t.tsubtype == 'text':
                        returnstring += '\"' + t.tvalue + '\"'
                    elif t.tsubtype == 'number' and t.tvalue == '':
                        returnstring += '\"\"'
                    else:
                        returnstring += t.tvalue

                if (t.tsubtype == p.TOK_SUBTYPE_START):
                    indent += 1;

        return returnstring.strip()


def whole_process_loop(dic, filepath, total_or_unique):
    original_formu_path = filepath+'nestif-'+total_or_unique+'-original.txt'
    refactor_result_path = filepath+'nestif-'+total_or_unique+'-refactored.txt'
    failpath =  filepath+'nestif-'+total_or_unique+'-fail.txt'
    nodepthreducefile =  filepath+'nestif-zeroreduce-'+total_or_unique+'.txt'

    readfile = open(original_formu_path)
    writesuccess = open(refactor_result_path, 'w')
    writefail = open(failpath,'w')
    zeroreduce = open(nodepthreducefile, 'w')

    count = 0
    totalcount = 0


    while (True):
        thisline = readfile.readline()
        if not thisline:
            break
        totalcount +=1

        try:
            excelfilename, formula,  num = thisline.strip().split('::')
        except:
            try:
                excelfilename, formula = thisline.strip().split('::')
            except:
                print '1============================'+thisline
                continue

        if '#REF' in formula or '#DIV/0' in formula or '#VALUE' in formula:
            print '2============================'
            continue

        formula = predealwith_formula(formula)

        if not formula:
            print '3============================'+thisline
            continue
        originalformu = formula

        print formula
        dic = {}
        dic['and'] = False
        dic['or'] = False
        dic['lookup'] = False
        dic['idlookup'] = False
        dic['choose'] = False
        dic['match'] = False
        dic['ifs'] = False
        dic['maxmin'] = False
        dic['useless'] = False
        dic['redun'] = False



        returnlist = dealWithRudunPatterns.redundancy_loop_NEW(formula)


        if returnlist != formula:
            isother = False
            dic['redun'] = True

            formula = returnlist
        try:

            newformula = refactor_loop_NEW(formula, dic)
        except:
            continue

        if newformula and newformula!= originalformu:

            isother = False
            if count % 100 == 0:
                print count

            count += 1


            truefalsestring = str(dic['and']) + '::' + str(dic['or']) + '::' + str(dic['lookup']) + '::' + str(
                dic['idlookup']) + '::' + str(dic['choose']) + '::' + str(
                    dic['match']) + '::' + str(dic['ifs']) + '::' +str(dic['maxmin']) + '::' + str(dic['useless']) + '::' + str(dic[
                        'redun'])

            p = ForNestedIf.ExcelParser()
            p.parse('='+originalformu)
            olddepth = p.get_nested_ifs()
            try:
                p.parse('=' + newformula)
            except:
                writesuccess.write(
                    excelfilename+'::'+originalformu + '::' + newformula + '::' + truefalsestring  + '\n')
                continue
            newdepth = p.get_nested_ifs()
            depthreduce = olddepth - newdepth
            if depthreduce == 0:
                zeroreduce.write(
                    excelfilename+'::'+originalformu + '::' + newformula + '::' + truefalsestring + '::' + str(olddepth) + '::' + str(newdepth) +'\n')

            else:
                writesuccess.write(excelfilename+'::'+originalformu + '::' + newformula+ '::'+truefalsestring+'::'+str(olddepth)+'::'+str(newdepth)+'\n')
        else:
            # print formula
            writefail.write(excelfilename+'::'+originalformu+'\n')

    writesuccess.close()
    writefail.close()
    zeroreduce.close()
    print count


def pattern_match_oneround(formula,dic):

    # input: a if formula
    # output: return the new formula if pattern exists, otherwise return false

    isother = True

    p = ForNestedIf.ExcelParser()
    try:
        p.parse('=' + formula)
    except:
        return False
    threepartlist = [[], [], []]

    p.get_list_threeparts(formula, threepartlist)

    if threepartlist[1] == '' and  threepartlist[2] == '':
        formula = threepartlist[0]
        try:
            p.parse('=' + formula)
        except:
            return False
    newformula = getReducedFormulas.get_simplified_AND_pattern(formula, threepartlist)
    if newformula:
        isother = False

        modify_truefalsedic(dic,'and')
        return newformula

    newformula = getReducedFormulas.get_simplified_OR_pattern(formula, threepartlist)
    if newformula:
        isother = False
        modify_truefalsedic(dic, 'or')
        return newformula
    newformula = return_true_ifis_otherOR(formula, threepartlist)
    if newformula:
        isother = False
        modify_truefalsedic(dic, 'or')
        return newformula

    if return_list_ifis_equal(formula, threepartlist) and formula.count('IF') > 3:



        isother = False
        dic_small = return_list_ifis_equal(formula, threepartlist)


        # dic_choose_match_lookup_idlookup
        if dic_small['choose'] == True:

            newformula = getReducedFormulas.get_simplified_CHOOSE_pattern(formula, threepartlist)


            if newformula:
                modify_truefalsedic(dic, 'choose')
                return newformula

        if dic_small['match'] == True:

            newformula = getReducedFormulas.get_simplified_MATCH_pattern(formula, threepartlist)
            if newformula:
                modify_truefalsedic(dic, 'match')
                return newformula

        if dic_small['lookup'] == True:

            newformula = getReducedFormulas.get_simplified_LOOKUP_pattern(formula, threepartlist)
            if newformula:
                modify_truefalsedic(dic, 'lookup')
                return newformula

        if dic_small['idlookup'] == True:
            newformula = getReducedFormulas.get_simplified_ID_LOOKUP_pattern(formula,
                                                                             threepartlist)

            if newformula:
                modify_truefalsedic(dic, 'idlookup')
                return newformula

    newformula = getReducedFormulas.get_simplified_MAXMIN_pattern(formula)
    if newformula:
        isother = False
        modify_truefalsedic(dic, 'maxmin')
        return newformula

    newformula = getReducedFormulas.get_simplified_USELESS_pattern(formula)

    if newformula:
        isothere = False
        modify_truefalsedic(dic, 'useless')
        return newformula

    newformula = getReducedFormulas.get_simplified_IFS_pattern(formula, threepartlist)


    if newformula:
        isother = False
        modify_truefalsedic(dic, 'ifs')
        return newformula

    return False


def modify_truefalsedic(dic, string):
    dic[string] = True



if __name__ == '__main__':

    dic = {}
    filepath = '/Users/jiezhang/research-projects/nestedifs/homepage/artifact/'
    whole_process_loop(dic,filepath,'total')








