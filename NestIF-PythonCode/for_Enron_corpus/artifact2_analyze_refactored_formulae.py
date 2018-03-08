import ForNestedIf

def deal_with_finalcsv(filepath):
    readfile = open(filepath)
    lines = readfile.readlines()
    tfandnum = 0
    tfornum = 0
    tflookupnum = 0
    tfidlookupnum = 0
    tfchoosenum = 0
    tfmatchnum = 0
    tfifsnum = 0
    tfmaxminnum = 0
    tfuselessnum = 0
    tfredunnum = 0

    for eachline in lines[1:]:
        eachline = eachline.strip()
        splits = eachline.split(',')

        tfand = splits[4]
        tfor = splits[5]
        tflookup = splits[6]
        tfidlookup = splits[7]
        tfchoose = splits[8]
        tfmatch = splits[9]
        tfifs = splits[10]
        tfmaxmin= splits[11]
        tfuseless= splits[12]
        ifredun= splits[13]

        if 'True' in tfand:
            tfandnum +=1
        if 'True' in tfor:
            tfornum +=1
        if 'True' in tflookup:
            tflookupnum +=1
        if 'True' in tfidlookup:
            tfidlookupnum +=1
        if 'True' in tfchoose:
            tfchoosenum +=1
        if 'True' in tfmatch:
            tfmatchnum +=1
        if 'True' in tfifs:
            tfifsnum +=1
        if 'True' in tfmaxmin:
            tfmaxminnum +=1
        if 'True' in tfuseless:
            tfuselessnum +=1
        if 'True' in ifredun:
            tfredunnum += 1
    print 'AND pattern: '+str(tfandnum)
    print 'OR pattern: '+str(tfornum)
    print 'LOOKUP pattern: '+str(tflookupnum+tfidlookupnum)
    print 'CHOOSE pattern: '+str(tfchoosenum)
    print 'MATCH pattern: '+str(tfmatchnum)
    print 'IFS pattern: '+str(tfifsnum)
    print 'MAXMIN pattern: '+str(tfmaxminnum)
    print 'USELESS pattern: '+str(tfuselessnum)
    print 'REDUN pattern: '+str(tfredunnum)



def get_final_func_result(sourcefilepath,resultfilepath):
    writefile = open(resultfilepath,'w')
    readfile = open(sourcefilepath)

    count = 0
    lines = readfile.readlines()


    for newline in lines:
        newline = newline.strip()

        if newline:
            count +=1

            try:
                excelname,oldformula,newformula, tfand, tfor, tflookup, tfidlookup, tfchoose, tfmatch, tfifs, tfmaxmin, tfuseless, ifredun,olddepth, newdepth = newline.split('::')
            except:
                continue

            # print num_formula
            p = ForNestedIf.ExcelParser()
            p.parse('='+oldformula)
            old_depth = p.get_nested_ifs()
            p.parse('=' + newformula)
            new_depth = p.get_nested_ifs()
            depthreduce = old_depth - new_depth
            proportion = depthreduce/old_depth
            writefile.write(olddepth+','+newdepth+','+str(depthreduce)+','+str(proportion)+','+tfand+','+ tfor+','+tflookup+','+tfidlookup+','+ tfchoose+','+ tfmatch+','+ tfifs+','+ tfmaxmin+','+ tfuseless+','+ ifredun+'\n')
        else:
            continue
    writefile.close()



def deal_with_zeroreduce(sourcefilepath,resultfilepath):
    writefile = open(resultfilepath,'a')
    readfile = open(sourcefilepath)

    count = 0

    totalcount = 0


    while(True):
        newline = readfile.readline().strip()
        depthreduce = 0

        if newline:
            count +=1
            if count%1000 == 0:
                print count
            path, sheetname,oldformula,newformula, tfand, tfor, tflookup, tfidlookup, tfchoose, tfmatch, tfifs, tfmaxmin, tfuseless, ifredun,olddepth, newdepth, num = newline.split('::')
            num_formula = float(num)

            # print num_formula
            p = ForNestedIf.ExcelParser()
            p.parse('='+oldformula)
            old_depth = p.get_nested_ifs()
            p.parse('=' + newformula)
            new_depth = p.get_nested_ifs()
            depthreduce = old_depth - new_depth
            if depthreduce>0:
                totalcount += num_formula
                proportion = depthreduce/old_depth
                writefile.write(olddepth+','+newdepth+','+str(depthreduce)+','+str(proportion)+','+num+'\n')


        else:
            break
    writefile.close()


def deal_depth_reduce(dic_depth_num,filepath):
    readfile = open(filepath)
    count = 0
    totalcount = 0
    lines = readfile.readlines()

    dic_depth_num[0.25] = 0
    dic_depth_num[0.5] = 0
    dic_depth_num[0.75] = 0
    dic_depth_num[1.0] = 0
    dic_func_func = {}

    for eachline in lines:

        newline = eachline.strip()
        if newline:
            count +=1
            try:
                excelname,oldformula, newformula, tfand, tfor, tflookup, tfidlookup, tfchoose, tfmatch, tfifs, tfmaxmin, tfuseless, ifredun, olddepth, newdepth = newline.split(
                        '::')
                totalcount += 1
            except:
                continue

            depthreduce = int(olddepth)-int(newdepth)
            reducerate = depthreduce*1.0/int(olddepth)
            depthreduce = int(newdepth)

            if reducerate <= 0.25:
                dic_depth_num[0.25] += 1
            elif reducerate <= 0.5:
                dic_depth_num[0.5] += 1
            elif reducerate <= 0.75:
                dic_depth_num[0.75] += 1
            elif reducerate <= 1.0:
                dic_depth_num[1.0] += 1
        else:
            continue


    return dic_depth_num

def deal_original_depth(dic_depth_num,filepath):
    sourcefilepath = filepath

    readfile = open(sourcefilepath)
    count = 0
    totalcount = 0

    while(True):
        newline = readfile.readline().strip()
        depthreduce = 0

        if newline:
            count +=1


            try:
                excelname,oldformula, newformula, tfand, tfor, tflookup, tfidlookup, tfchoose, tfmatch, tfifs, tfmaxmin, tfuseless, ifredun, olddepth, newdepth = newline.split(
                        '::')

                totalcount += 1
            except:
                continue
            if newdepth in dic_depth_num:
                dic_depth_num[newdepth]+=1
            else:
                dic_depth_num[newdepth] = 1

        else:
            break


    return dic_depth_num

def get_eachpattern_result(total_or_unique):
    eachpatternfilepath = filepath+'eachpattern-result.csv'
    get_final_func_result(
       filepath+'nestif-'+total_or_unique+'-refactored.txt',eachpatternfilepath)
    print '----------------------------------'
    print 'Number of formula with different patterns: '+total_or_unique

    deal_with_finalcsv(eachpatternfilepath)

def get_depth_reduce_result(total_or_unique,filepath):
    dic_depth_num = {}
    sourcefilepath = filepath+'nestif-'+total_or_unique+'-refactored.txt'
    dic = deal_depth_reduce(dic_depth_num,sourcefilepath)
    print '----------------------------------'
    print 'Number of formulae with different depth reduce rate: '+total_or_unique

    for i in dic:
        print 'depth reduce rate:'+str(i)+' || formulae number:'+str(dic[i])

def get_final_depth_result(total_or_unique,filepath):
    dic_depth_num = {}
    sourcefilepath = filepath+'nestif-'+total_or_unique+'-refactored.txt'
    dic = deal_original_depth(dic_depth_num,sourcefilepath)
    print '----------------------------------'
    print 'Number of formulae with different new depth: '+total_or_unique

    for i in dic:
        print 'depth reduce rate:'+str(i)+' || formulae number:'+str(dic[i])




if __name__ == '__main__':

    total_or_unique = 'total'
    filepath = '/Users/jiezhang/research-projects/nestedifs/homepage/artifact/'
    get_depth_reduce_result(total_or_unique,filepath)
    get_final_depth_result(total_or_unique,filepath)
    get_eachpattern_result(total_or_unique)






