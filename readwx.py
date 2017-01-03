#coding=utf-8

import xlrd
import datetime
import xlwt

def open_excel(file = '/Users/Mike/Documents/wx-syh.xls'):
    try:
        print('read file')
        data = xlrd.open_workbook('wx_syh_new.xls')
        print('read success')
        return data
    except Exception,e:
        print str(e)

def wx_get_table_line():
    data = open_excel()
    table = data.sheet_by_index(0)
    nrows = table.nrows
    ncols = table.ncols
    print('This WX is all read Success')
    print(nrows)

def GetWordNums(text):
    num = 0
    for i in text:
        if i not in ' \n!"#$%&()*+,-./:;<=>?@[\\]^_`{|}~':
            num = num +1
    return num

def wx_get_table_line_eachother():
    selfcount=0
    friendcount=0
    data = open_excel()
    table = data.sheet_by_index(0)
    nrows = table.nrows
    for rownum in range(nrows):
        sendmessagename = table.cell(rowx = rownum,colx = 0 ).value
        if sendmessagename == 'brakes':
            selfcount = selfcount + 1
        else:
            friendcount = friendcount + 1
    print("You have sent all message counts") 
    print(selfcount)
    print("Your messages precent in all messages")
    print( selfcount*1.0/nrows*1.0 )
    print("friend have sent all message counts") 
    print(friendcount)
    print("friend's messgaes precent in all messages") 
    print( friendcount*1.0/nrows * 1.0 )

def wx_get_table_line_eachother_gif():
    s=u'[动画表情]'
    selfcount=0
    gifcount=0
    friendcount=0
    data = open_excel()
    table = data.sheet_by_index(0)
    nrows = table.nrows
    for rownum in range(nrows):
        sendmessagename = table.cell(rowx = rownum,colx = 5 ).value
        if sendmessagename.encode('gb18030')  == s.encode('gb18030'):
            gifcount = gifcount + 1
            if table.cell(rowx = rownum,colx = 0 ).value == 'brakes':
                selfcount = selfcount + 1
            else:
                friendcount = friendcount + 1
    print("You have all gif count is ")
    print(gifcount)
    print("You have sent all message counts") 
    print(selfcount)
    print("Your gif messages precent in all messages")
    print( selfcount*1.0/gifcount*1.0 )
    print("friend have sent all message counts") 
    print(friendcount)
    print("friend's gif messgaes precent in all messages") 
    print( friendcount*1.0/gifcount * 1.0 )

def wx_get_words_count_eachother():
    # gb18030 分析了消息两个内容
    selfcount=0-4
    wordcount=0-4
    friendcount=0-4
    data = open_excel()
    table = data.sheet_by_index(0)
    nrows = table.nrows
    for rownum in range(nrows):
        sendmessagename = table.cell(rowx = rownum,colx = 5 ).value
        wordcount = wordcount + GetWordNums(sendmessagename.encode('gb18030'))
        if table.cell(rowx = rownum,colx = 0 ).value == 'brakes':
            selfcount = selfcount + GetWordNums(sendmessagename.encode('gb18030'))
        else:
            friendcount = friendcount + GetWordNums(sendmessagename.encode('gb18030'))
    print("All word count is ")

    print(wordcount/2)
    print("You have sent all message word counts") 
    print(selfcount/2)
    print("Your messages word precent in all messages")
    print( selfcount*1.0/wordcount*1.0 )
    print("friend have sent all message word counts") 
    print(friendcount/2)
    print("friend's message messgaes precent in all messages") 
    print( friendcount*1.0/wordcount* 1.0 )

def wx_get_timespan_everyday():
    selfcount=0
    friendcount=0
    daynumber=0
    data = open_excel()
    table = data.sheet_by_index(0)
    nrows = table.nrows
    nrowslimit = nrows - 1
    #判断隔天时间
    daydatetime=""
    singledaycount=0
    #判断哪天最高
    mostmessagecount=0
    mostmessagecday=""
    #导出处excel用
    workbook=xlwt.Workbook(encoding='ascii')
    worksheet=workbook.add_sheet('My Worksheet')
    #计算平均间隔要存储的时间
    starttime = table.cell(rowx = 2,colx = 1 ).value
    starttime_s = starttime.encode('utf-8')
    savedatetime = datetime.datetime.strptime(starttime_s, "%Y-%m-%d %H:%M:%S")
    endtime = table.cell(rowx = nrowslimit,colx = 1 ).value
    endtime_s = endtime.encode('utf-8')
    enddatetime = datetime.datetime.strptime(endtime_s,"%Y-%m-%d %H:%M:%S")
    #时间间隔累加
    timespan=0.0
    for rownum in range(2,nrows):
    
        time = table.cell(rowx = rownum,colx = 1 ).value
        time_s = time.encode('utf-8')
        curdatetime = datetime.datetime.strptime(time_s, "%Y-%m-%d %H:%M:%S")

        if curdatetime.strftime("%Y-%m-%d")!=daydatetime:
            timespan=0.0
            if savedatetime!=curdatetime:
                time_from_to = (curdatetime-savedatetime).total_seconds()
                savedatetime=curdatetime
                timespan+=time_from_to
            print(singledaycount)
            print("timespan this date")
            if singledaycount!=0:
                print(timespan/singledaycount)
                saveday = curdatetime-datetime.timedelta(days=1)
                savedaystr = saveday.strftime("%Y-%m-%d")
                print(saveday)
                worksheet.write(daynumber,1,label = timespan/singledaycount)
                worksheet.write(daynumber,0,label = savedaystr)
            print(curdatetime.strftime("%Y-%m-%d"))
            #通过判断字符变化来判断心的一天
            if mostmessagecount<singledaycount:
                mostmessagecount=singledaycount
                mostmessagecday=curdatetime.strftime("%Y-%m-%d")
            singledaycount=0
            daydatetime=curdatetime.strftime("%Y-%m-%d")
            daynumber=daynumber+1

            
        singledaycount=singledaycount+1
        if curdatetime == enddatetime:
            print(singledaycount)
            print(curdatetime.strftime("%Y-%m-%d"))
            worksheet.write(daynumber,1,label = timespan/singledaycount)
            worksheet.write(daynumber,0,label = curdatetime.strftime("%Y-%m-%d"))
        #print(time)
        #curdatetime = datetime.datetime.strptime(time_s, "%Y-%m-%d")
    print(daynumber)
    print("That day you send message for most value:")
    print(mostmessagecday)
    print(mostmessagecount)
    workbook.save("wx_syh_everyday_timespan.xls")   


def wx_get_time_everyday_line():
    selfcount=0
    friendcount=0
    daynumber=0
    data = open_excel()
    table = data.sheet_by_index(0)
    nrows = table.nrows
    nrowslimit = nrows-1
    #判断隔天时间
    daydatetime=""
    singledaycount=0
    #判断哪天最高
    mostmessagecount=0
    mostmessagecday=""
    #导出处excel用
    workbook=xlwt.Workbook(encoding='ascii')
    worksheet=workbook.add_sheet('My Worksheet')
    #第一天和最后一天分开处理
    starttime = table.cell(rowx = 2,colx = 1 ).value
    starttime_s = starttime.encode('utf-8')
    savedatetime = datetime.datetime.strptime(starttime_s, "%Y-%m-%d %H:%M:%S")
    endtime = table.cell(rowx = nrowslimit,colx = 1 ).value
    endtime_s = endtime.encode('utf-8')
    enddatetime = datetime.datetime.strptime(endtime_s,"%Y-%m-%d %H:%M:%S")
    for rownum in range(3,nrows):
        #第一天和最后一天要分开处理
        time = table.cell(rowx = rownum,colx = 1 ).value
        time_s = time.encode('utf-8')
        curdatetime = datetime.datetime.strptime(time_s, "%Y-%m-%d %H:%M:%S")
        if curdatetime.strftime("%Y-%m-%d")!=daydatetime:
            #更新新的一天
            print(singledaycount)
            saveday = curdatetime-datetime.timedelta(days=1)
            savedaystr = saveday.strftime("%Y-%m-%d")
            print(saveday)
            worksheet.write(daynumber,1,label = singledaycount)
            worksheet.write(daynumber,0,label = savedaystr)
            #通过判断字符变化来判断心的一天
            if mostmessagecount<singledaycount:
                mostmessagecount=singledaycount
                mostmessagecday=curdatetime.strftime("%Y-%m-%d")
            singledaycount=0
            daydatetime=curdatetime.strftime("%Y-%m-%d")
            daynumber=daynumber+1
        singledaycount=singledaycount+1
        if curdatetime == enddatetime:
            print(singledaycount)
            print(curdatetime.strftime("%Y-%m-%d"))
            worksheet.write(daynumber,1,label = singledaycount)
            worksheet.write(daynumber,0,label = curdatetime.strftime("%Y-%m-%d"))
        #print(time)
        #curdatetime = datetime.datetime.strptime(time_s, "%Y-%m-%d")
    print(daynumber)
    print("That day you send message for most value:")
    print(mostmessagecday)
    print(mostmessagecount)
    workbook.save("wx_syh_everyday_counts.xls")   
    
def wx_write_xls(nrows,ncol,content,filename):
    workbook=xlwt.Workbook(encoding='ascii')
    worksheet=workbook.add_sheet('My Worksheet')
    worksheet.write(nrows,ncol,label = content)
    workbook.save(filename)
    
def main():
    #wx_get_table_line()
    #wx_get_table_line_eachother()
    #wx_get_table_line_eachother_gif()
    #wx_get_words_count_eachother()
    #wx_get_time_everyday_line()
    wx_get_timespan_everyday()
    #wx_write_xls_test()

if __name__=="__main__":
    main()