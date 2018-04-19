#coding:utf-8
import re
import datetime
import os
import IPy
import xlrd
import xlwt
from xlutils.copy import copy

print ("Start...")

startTime = datetime.datetime.now()
print (startTime)

def filterouttheIPs():
    path0 = raw_input("The switch file's path is : \n")

    filelist = os.listdir(path0)
    print (filelist)

    f2 = open('.\ipaddress.txt', 'w')

    for i in range(0, len(filelist)):
        path1 = os.path.join(path0, filelist[i])
        if os.path.isfile(path1):

            f1 = open(path1, 'r')

            while True:
                linecontext = f1.readline()
                if 'NPI' in linecontext:
                    linecontext = re.sub('\n', ' ', linecontext)
                    f2.write(linecontext, )
                    targettext = f1.readline()
                    f2.write(targettext)
                if linecontext == '':
                    break
            f2.write('\n')

            f1.close()
    f2.close()

if __name__ == "__main__":
    filterouttheIPs()
    
    IPsegments1 = []
    
    f3 = open('.\ipaddress.txt', 'r')
    workbooklog = xlwt.Workbook(encoding = 'utf-8')
    worksheetlog = workbooklog.add_sheet('result', cell_overwrite_ok = False)
    H = 0
    excelheader = ['序號', 'IP網段','IP網段-首', 'IP網段-尾', '廠房','廠房名稱',\
                   '廠房所屬區域', '樓層', '掩碼','網關', '所屬廠區', '所屬廠區名稱',\
                   '是否NPI區', '備注', '有效否', 'VLAN ID','IP版本', '網絡段類型', \
                   '是否啟用DHCP']
    L = 0
    for workbookcell in excelheader:
        worksheetlog.write(H, L, workbookcell)
        L += 1
        
    H =1
    while True:
        linecontext2 = f3.readline()
        L = 0
    
        if linecontext2 == None:
            linecontext2 = f3.readline()
        elif linecontext2 == " ":
            linecontext2 = f3.readline()
        elif linecontext2 == "\n":
            linecontext2 = f3.readline()
        elif linecontext2 =='':
            break
        else:
            #print 'linecontext2 :' , linecontext2
            description = re.search("IPE[a-zA-Z0-9\-\.\/]*",linecontext2, re.X).group()
            descriptionlist = description.split('-') 
            inputIP = re.search("10+\.\d+\.\d+\.\d*", linecontext2, re.X).group()
            #print 'inputIP ', inputIP
            inputnetmask = re.search("255+\.\d+\.\d+\.\d*", linecontext2, re.X).group()
            #print 'inputnetmask', inputnetmask
            IPsegment = IPy.IP(inputIP).make_net(inputnetmask).strNormal()
            IPsegments1.append(IPsegment)
            
            listIPsegmentforgateway = []
            IPsegmentforgateway = IPsegment = IPy.IP(inputIP).make_net(inputnetmask)
            for x in IPsegmentforgateway:
                listIPsegmentforgateway.append(x)
            

        print ("IPsegment :%s " % IPsegment, "description :%s" %description)
        m = IPsegment.strNormal().split('/')
        n = m[1]
        if int(n) < 24:
            ip1 = IPy.IP(IPsegment)
            ip1segmentlength = ip1.len() / 256
            ip1list = ip1.strNormal(0).split('.')
            ip1subsegment = [ip1.strNormal(0)]

            for i in range(1, ip1segmentlength):
                k = int(ip1list[2]) + i                                
                j = ip1list[0] +  '.' + ip1list[1] + '.' + str(k) + '.' + ip1list[3]
                ip1subsegment.append(j)
            print (ip1subsegment)
            
            for i in range(0, len(ip1subsegment)):
                worksheetlog.write(H, 0, str(H))
                worksheetlog.write(H, 1, IPsegment.strNormal()) #.net().strNormal()
                IPsegmentheader = IPy.IP(ip1subsegment[i] +'/24')
                IPsegmentheaderlist = []
                for x in IPsegmentheader:
                    IPsegmentheaderlist.append(x)
                worksheetlog.write(H, 2, IPsegmentheaderlist[2].strNormal()) #IPy.IP(ip1subsegment[i] +'/24').net().strNormal()
                worksheetlog.write(H, 3, IPy.IP(ip1subsegment[i] +'/24').broadcast().strNormal())
                worksheetlog.write(H, 4, descriptionlist[1])
                worksheetlog.write(H, 5, descriptionlist[1])
                zone = str(descriptionlist[1])
                zone = re.search('[A-Za-z]*', zone, re.X).group()
                worksheetlog.write(H, 6, zone)
                worksheetlog.write(H, 7, descriptionlist[2])
                worksheetlog.write(H, 8, inputnetmask)
                worksheetlog.write(H, 9, listIPsegmentforgateway[1].strNormal())
                worksheetlog.write(H, 10, descriptionlist[0])
                worksheetlog.write(H, 11, "n/a")
                worksheetlog.write(H, 12, descriptionlist[3])
                worksheetlog.write(H, 13, 'n/a')
                worksheetlog.write(H, 14, 'Y')
                worksheetlog.write(H, 15, descriptionlist[4])
                worksheetlog.write(H, 16, IPsegment.version())
                worksheetlog.write(H, 17, 'n/a')
                worksheetlog.write(H, 18, 'n/a')
                H += 1
        else:
            worksheetlog.write(H, 0, str(H))
            worksheetlog.write(H, 1, IPsegment.strNormal()) #.net().strNormal()
            worksheetlog.write(H, 2, inputIP)
            worksheetlog.write(H, 3, IPsegment.broadcast().strNormal())
            worksheetlog.write(H, 4, descriptionlist[1])
            worksheetlog.write(H, 5, descriptionlist[1])
            zone = str(descriptionlist[1])
            zone = re.search('[A-Za-z]*', zone, re.X).group()
            worksheetlog.write(H, 6, zone)
            worksheetlog.write(H, 7, descriptionlist[2])
            worksheetlog.write(H, 8, inputnetmask)
            worksheetlog.write(H, 9, listIPsegmentforgateway[1].strNormal())
            worksheetlog.write(H, 10, descriptionlist[0])
            worksheetlog.write(H, 11, "n/a")
            worksheetlog.write(H, 12, descriptionlist[3])
            worksheetlog.write(H, 13, 'n/a')
            worksheetlog.write(H, 14, 'Y')
            worksheetlog.write(H, 15, descriptionlist[4])
            worksheetlog.write(H, 16, IPsegment.version())
            worksheetlog.write(H, 17, 'n/a')
            worksheetlog.write(H, 18, 'n/a')
            H += 1


            
    workbooklog.save(".\IPsegmentlist.xls")
    f3.close()
    pause = raw_input("Press any to exit ...")
