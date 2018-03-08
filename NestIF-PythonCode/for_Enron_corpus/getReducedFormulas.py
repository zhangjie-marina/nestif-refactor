import ForNestedIf
# import furtherAnalysis
import re
from openpyxl import load_workbook

def get_simplified_AND_pattern(formula,threepartlist):


    condition_part_list = threepartlist[0]
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]

    if not (len(set(false_part_list)) == 1 and len(false_part_list) > 1):
        return False
    for each in true_part_list[:-1]:
        if 'IF' not in each:
            return False

    truevalue = true_part_list[-1]
    falsevalue = false_part_list[0]
    condition = ''
    for i in condition_part_list:
        condition+=(i+',')
    final = 'IF(AND('+condition[:-1]+'),'+truevalue+','+falsevalue+')'
    # final = formula.replace(onlyif, final)
    return final

def get_simplified_OR_pattern(formula,threepartlist):
    condition_part_list = threepartlist[0]
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]
    if not (len(set(true_part_list)) == 1 and len(true_part_list) > 1):
        return False
    for each in false_part_list[:-1]:
        if 'IF' not in each:
            return False

    truevalue = true_part_list[0]
    if len(false_part_list) >0:
        falsevalue = false_part_list[-1]
    else:
        falsevalue = ''
    condition = ''
    for i in condition_part_list:
        condition+=(i+',')
    final = 'IF(OR('+condition[:-1]+'),'+truevalue+','+falsevalue+')'
    return final

def get_simplified_USELESS_pattern(formula):

    p = ForNestedIf.ExcelParser()
    try:
        p.parse('='+formula)
    except:
        return False
    all_innerif_list = p.get_all_innerif_list()
    all_innerif_list.append(formula)

    if len(all_innerif_list) == 0:
        return False

    for eachone in all_innerif_list:


        # if eachone.count('IF') != 1:
        #     continue
        p.parse('='+eachone)

        threeparts = p.get_threeparts_IF()




        if (threeparts[2]+'='+threeparts[1]) == threeparts[0]:

            replacestring = 'IF('+threeparts[0]+','+threeparts[1]+','+threeparts[2]+')'

            return formula.replace(replacestring,threeparts[2])
        elif (threeparts[1]+'='+threeparts[2]) == threeparts[0]:

            replacestring = 'IF('+threeparts[0] + ',' + threeparts[1] + ',' + threeparts[2]+')'
            return formula.replace(replacestring, threeparts[1])

    return False

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

def get_simplified_ID_LOOKUP_pattern(formula, threepartlist):

    condition_part_list = threepartlist[0]
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]

    before_equal_list = []
    after_equal_list = []
    for each in condition_part_list:
        before = each.split('=')[0]
        after = each.split('=')[1]

        after_equal_list.append(after)
        before_equal_list.append(before)
    final = 'VLOOKUP(' + before_equal_list[0] + ',' + 'RangeA1' + ':' + 'RangeB'+str(len(true_part_list))+ ',2' + ')'
    return final

def get_simplified_LOOKUP_pattern(formula, excelname, sheetname, threepartlist):
    condition_part_list = threepartlist[0]
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]

    before_equal_list = []
    after_equal_list = []



    for each in condition_part_list:
        before = each.split('=')[0]
        after = each.split('=')[1]

        after_equal_list.append(after)
        before_equal_list.append(before)


    first_alpha_list = []
    first_num_list = []


    for each in after_equal_list:
        if '!' in each:
            each = each.split('!')[1]
        each = each.replace('$','')
        match1 = re.match(r"([A-Z]+)([0-9]+)", each, re.I)


        if match1:
            items = match1.groups()
            first_alpha_list.append(items[0])
            first_num_list.append(items[1])
    second_alpha_list = []
    second_num_list = []

    if true_part_list[0].split('!')[0] != after_equal_list[0].split('!')[0]:
        return False

    for each in true_part_list:
        if '!' in each:
            each = each.split('!')[1]
        each = each.replace('$', '')
        match1 = re.match(r"([A-Z]+)([0-9]+)", each, re.I)

        if match1:
            items = match1.groups()
            second_alpha_list.append(items[0])
            second_num_list.append(items[1])

    first = ''
    second = ''



    if len(set(first_alpha_list)) == 1 and len(set(second_alpha_list)) == 1:
        first_num_list.sort()
        second_num_list.sort()
        if first_num_list == second_num_list and is_arithmetic(first_num_list):
            if '!' in true_part_list[0]:
                first = true_part_list[0].split('!')[0]+'!'+'$'+first_alpha_list[0]+'$'+first_num_list[0]
                second = true_part_list[0].split('!')[0]+'!'+'$'+second_alpha_list[0]+'$'+first_num_list[-1]
            else:
                first = '$' + first_alpha_list[0] + '$' + first_num_list[0]
                second = '$' + second_alpha_list[0] + '$' + first_num_list[-1]
            final = 'VLOOKUP('+before_equal_list[0]+','+first+':'+second+','+str(1+abs(ord(first_alpha_list[0].upper()) - ord(second_alpha_list[0].upper())))+')'
            return final
        else:
            return False
    elif len(set(first_num_list)) == 1 and len(set(second_num_list)) == 1:
        first_alpha_list.sort()
        second_alpha_list.sort()
        if first_alpha_list == second_alpha_list and if_celllist_neighbours(first_alpha_list):
            if '!' in true_part_list[0]:
                first = true_part_list[0].split('!')[0] + '!' + '$' + first_alpha_list[0] + '$' + first_num_list[0]
                second = true_part_list[0].split('!')[0] + '!' + '$' + second_alpha_list[0] + '$' + first_num_list[-1]
            else:
                first = '$' + first_alpha_list[0] + '$' + first_num_list[0]
                second = '$' + second_alpha_list[0] + '$' + first_num_list[-1]
            final = 'HLOOKUP(' + before_equal_list[0] + ',' + first + ':' + second + ',' + str(1+abs(int(second_num_list[0])-int(first_num_list))) + ')'
            return final
        else:
            return False

def if_celllist_neighbours(thislist):
    try:
        if thislist.join('') in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
            return True
        total = ''
        firstlist = []
        for each in thislist:
            firstlist.append(each[0])
            each = each[1:]
            total.join(each)
        if len(set(firstlist)) != 1:
            return False
        if total in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
            return True
    except:
        return False
    return False



def get_simplified_MAXMIN_pattern(formula):
    p = ForNestedIf.ExcelParser()
    try:
        p.parse(formula)
    except:
        return False
    all_innerif_list = p.get_all_innerif_list()
    all_innerif_list.append(formula)
    if len(all_innerif_list) == 0:
        return False

    for eachone in all_innerif_list:
        final = ''
        if eachone.count('IF') != 1:
            continue
        threeparts = p.get_threeparts_IF()
        for eachsymbol in ['<','>','<=','>=']:

            if (threeparts[1] + eachsymbol + threeparts[2]) == threeparts[0] or (
                                    threeparts[1] + eachsymbol + '(' + threeparts[2] + ')') == threeparts[0] or (
                                    '(' + threeparts[1] + ')' + eachsymbol +threeparts[2]) == threeparts[0] or (
                                '(' + threeparts[1] + ')' + eachsymbol+ '(' + threeparts[2] + ')') == threeparts[0]:
                if eachsymbol == '>' or eachsymbol == '>=':

                    final = 'MAX('+threeparts[1]+','+threeparts[2]+')'
                elif eachsymbol == '<' or eachsymbol == '<=':

                    final = 'MIN(' + threeparts[1] + ',' + threeparts[2] + ')'
            if (threeparts[2] + eachsymbol + threeparts[1]) == threeparts[0] or (
                                    threeparts[2] + eachsymbol + '(' + threeparts[1] + ')') == threeparts[0] or (
                                    '(' + threeparts[2] + ')' + eachsymbol +threeparts[1]) == threeparts[0] or (
                                '(' + threeparts[2] + ')' + eachsymbol+ '(' + threeparts[1] + ')') == threeparts[0]:
                if eachsymbol == '>' or eachsymbol == '>=':

                    final = 'MIN('+threeparts[1]+','+threeparts[2]+')'
                elif eachsymbol == '<' or eachsymbol == '<=':
                    final = 'MAX(' + threeparts[1] + ',' + threeparts[2] + ')'

        if final != '':
            return formula.replace(eachone,final)
    return False

def get_simplified_IFS_pattern(formula, threepartlist):
    condition_list = threepartlist[0]
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]


    if len(condition_list) >1:
        if condition_list[0].replace(condition_list[1],'').count('IF')>1:
            return False



    if len(condition_list) <2:
        return False


    count = 0
    finalstring = ''

    # try:
    if 'IF' in condition_list[0]:
        return False
    while(count < len(true_part_list)):

        finalstring += condition_list[count]+','+true_part_list[count]+','

        count+=1
        if 'IF' in true_part_list[count-1]:
            count += 1

    finalstring += 'TRUE,'+false_part_list[-1]
    return 'IFS('+finalstring+')'


def get_simplified_CHOOSE_pattern(formula, threepartlist):
    condition_part_list = threepartlist[0]
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]

    before_equal_list = []
    after_equal_list = []
    p = ForNestedIf.ExcelParser()

    for each in condition_part_list:
        before = each.split('=')[0]
        after = each.split('=')[1]


        try:
            p.parse('=' + after)
        except:

            return False

        after_equal_list.append(after)
        before_equal_list.append(before)

    choose_string = ''

    for each in true_part_list:
        choose_string+=each+','

    initial = '('+ before_equal_list[0]+'-'+after_equal_list[0]+')/('+after_equal_list[1]+'-'+after_equal_list[0]+')'


    choose_string = choose_string[:-1] #remove the last ','
    final = 'IFERROR(CHOOSE('+str(initial)+'+1 ,'+choose_string+'),'+false_part_list[-1]+')'

    return final

def get_simplified_MATCH_pattern(formula, threepartlist):
    condition_part_list = threepartlist[0]
    true_part_list = threepartlist[1]
    false_part_list = threepartlist[2]

    before_equal_list = []
    after_equal_list = []
    p = ForNestedIf.ExcelParser()

    for each in condition_part_list:
        before = each.split('=')[0]
        after = each.split('=')[1]

        try:
            p.parse('=' + after)
        except:
            # print 'exception: cannot parse after equal part: ' + formula
            return False

        after_equal_list.append(after)
        before_equal_list.append(before)

    match_string = ''

    for each in after_equal_list:
        match_string+=each+','

    number =  float(true_part_list[1]) - float(true_part_list[0])

    match_string = match_string[:-1] #remove the last ','
    if (float(true_part_list[1]) - float(true_part_list[0]) == 1.0) and true_part_list[0] == '1':
        final = 'IFERROR(MATCH('+before_equal_list[0]+',{'+match_string+'},0),'+false_part_list[-1]+')'
    else:
        final = 'IFERROR((MATCH('+before_equal_list[0]+',{'+match_string+'},0)-1)*('+true_part_list[1]+'-'+true_part_list[0]+')+'+true_part_list[0]+','+false_part_list[-1]+')'

    return final

def main_for_ALL(numstring):

    p = ForNestedIf.ExcelParser()
    filenamelist = []

    filepath = 'D:\\Users\\v-jizha4\\results\\metrix-results\\Total-Matrix-' + numstring+'.txt'
    readfile = open(filepath)

    count = 0

    while (True):

        path = ''
        formula = ''
        truefalse = ''



        eachline = readfile.readline()

        if not eachline:
            break
        eachline = eachline.strip()
        i = eachline
        if i == '':
            continue


        if '[#This Row],' in i:
            i = i.replace('[#This Row],', '')
        count += 1
        if count <4178:
            continue


        path,formula,truefalse = i.split('::')

        truefalse_list = truefalse.split(',')

        try:
            if '.xlsm' in path:
                excelname = path.split('xlsm')[0]+'xlsm'
                sheetname = path[path.index('.xlsm')+5:path.index('.txt')]
            if '.xlsx' in path:
                excelname = path.split('xlsx')[0] + 'xlsx'
                sheetname = path[path.index('.xlsx') + 5:path.index('.txt')]
        except:
            print path



        i = formula

        try:
            p.parse(i)
        except:
            continue

        if count % 1000 == 0:
            print 'Matrix-'+numstring+': ' +str(count)

        threepartlist = [[], [], []]
        p.get_list_threeparts(i, threepartlist)

        write_filename = 'D:\\Users\\v-jizha4\\results\\after-refactor-results\\Refactored-'+numstring+'.txt'
        writefile = open(write_filename, 'a')

        if truefalse_list[0] == 'True':
            continue

        try:



            if truefalse_list[1] == 'True':
                newformula = get_simplified_AND_pattern(i,threepartlist)
                writefile.write(excelname+'::'+sheetname+'::'+formula+'::'+newformula+'::'+truefalse+'::AND\n')
                continue
            elif truefalse_list[2] == 'True':
                newformula = get_simplified_OR_pattern(i, threepartlist)
                writefile.write(excelname+'::'+sheetname+'::'+formula + '::' + newformula+'::'+truefalse+'::OR\n')
                continue
            elif truefalse_list[5] == 'True':
                newformula = get_simplified_CHOOSE_pattern(i, threepartlist)
                writefile.write(excelname+'::'+sheetname+'::'+formula + '::' + newformula+'::'+truefalse+'::CHOOSE\n')
                continue
            elif truefalse_list[6] == 'True':
                newformula = get_simplified_MATCH_pattern(i, threepartlist)
                writefile.write(excelname+'::'+sheetname+'::'+formula + '::' + newformula +'::'+truefalse+ '::MATCH\n')
                continue
            elif truefalse_list[3] == 'True':
                newformula = get_simplified_LOOKUP_pattern(i, excelname, sheetname, threepartlist)

                writefile.write(excelname+'::'+sheetname+'::'+formula+'::'+newformula+'::'+truefalse+'::LOOKUP\n')
                continue
            elif truefalse_list[4] == 'True':
                newformula = get_simplified_ID_LOOKUP_pattern(i, excelname, sheetname, threepartlist)

                writefile.write(excelname+'::'+sheetname+'::'+formula+'::'+newformula+'::'+truefalse+'::LOOKUP\n')
                continue

            elif truefalse_list[8] == 'True':
                newformula = get_simplified_MAXMIN_pattern(i, threepartlist)
                writefile.write(excelname+'::'+sheetname+'::'+formula + '::' + newformula +'::'+truefalse+ '::MAXMIN\n')
                continue
            elif truefalse_list[9] == 'True':
                newformula = get_simplified_USELESS_pattern(i, threepartlist)
                writefile.write(excelname+'::'+sheetname+'::'+formula + '::' + newformula +'::'+truefalse+ '::USELESS\n')
                continue
            elif truefalse_list[10] == 'True':
                newformula = get_simplified_IFS_pattern(i, threepartlist)
                writefile.write(excelname+'::'+sheetname+'::'+formula + '::' + newformula +'::'+truefalse+ '::IFS\n')
                continue


        except:

            continue




        writefile.close()


if __name__ == '__main__':
    # p = ForNestedIf.ExcelParser()
    # formula = 'IF(DJ21=[2]Reference!$A$2,[2]Reference!$B$2,IF(DJ21=[2]Reference!$A$3,[2]Reference!$B$3,IF(DJ21=[2]Reference!$A$4,[2]Reference!$B$4,IF(DJ21=[2]Reference!$A$5,[2]Reference!$B$5,IF(DJ21=[2]Reference!$A$6,[2]Reference!$B$6,[2]Reference!$B$7)))))'
    # p.parse(formula)
    # threepartlist = [[],[],[]]
    # p.get_list_threeparts(formula,threepartlist)
    #
    # print threepartlist
    # print '---------'
    # print get_simplified_LOOKUP_pattern(formula, '', '', threepartlist)
    main_for_ALL('two')

