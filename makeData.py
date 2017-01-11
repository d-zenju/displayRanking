# -*- coding: utf-8 -*-

import sys
import sqlite3
from datetime import datetime
import pandas
import xlsxwriter
import math
import types

# DB(SQLite3) path
dbpath = './yahoo/csv/yahoo.sqlite3'
csvpath = './category.csv'
yahoo_imgpath = './yahoo/image'
rakuten_imgpath = './rakuten/image'

# unixtime --> date
def unixtime2date(utime):
    loc = datetime.fromtimestamp(float(utime))
    loc = loc.strftime('%Y-%m-%d (%a)')
    return loc


# read category
def readCategory(csv_filepath, master_id):
    df = pandas.read_csv(csv_filepath)
    datas = df[df['MasterID'] == master_id]
    category = list(datas['CategoryName'])
    return category[0]


# check date all
def getData(site, master_id, period, timestamp):
    connector = sqlite3.connect(dbpath)
    cursor = connector.cursor()
    sql = 'select distinct * from ' + site + ' where master_id = "' + master_id + '" and period = "' + period + '" and timestamp = "' + timestamp + '" order by cast(rank as integer) asc'
    cursor.execute(sql)

    result = cursor.fetchall()
     
    cursor.close()
    connector.close()
    
    return result


# calc ranking
def calcRanking(site, master_id, period, datas):
    status = [[]]
    result = [[]]
    connector = sqlite3.connect(dbpath)
    cursor = connector.cursor()
    for i, rows in enumerate(datas):
        res = []
        for row in rows:
            if (i == 0):
                utime = int(row[2])-86400
                sql = 'select distinct timestamp, rank, img_id from ' + site + ' where master_id = "' + master_id + '" and period = "' + period + '" and img_id = "' + row[5] + '" and timestamp < "' + unicode(utime) + '" order by timestamp desc'
            else:
                print row[2]
                # type1
                utime = int(row[2])-86400
                sql = 'select distinct timestamp, rank, img_id from ' + site + ' where master_id = "' + master_id + '" and period = "' + period + '" and img_id = "' + row[5] + '" and timestamp < "' + unicode(utime) + '" order by timestamp desc'
                # type2
                #sql = 'select distinct timestamp, rank, img_id from ' + site + ' where master_id = "' + master_id + '" and period = "' + period + '" and img_id = "' + row[5] + '" and timestamp = "' + datas[i-1][0][2] + '" order by timestamp desc'
                print sql
            cursor.execute(sql)
            ress = cursor.fetchall()
            if (len(ress) > 0):
                res.append(ress[0])
            else:
                res.append((row[2], '0', row[5]))
        result.append(res)
    result = [x for x in result if x]
    cursor.close()
    connector.close()
    
    for i, rows in enumerate(result):
        state = []
        for j, row in enumerate(rows):
            if (int(row[1]) == 0):
                #print datas[i][j][5], datas[i][j][10], row[2], row[1]
                #print 'New'
                state.append('New')
            else:
                #print datas[i][j][5], datas[i][j][10], row[2], row[1]
                subrank = int(datas[i][j][10]) - int(row[1])
                subutime = (float(datas[i][j][2]) - float(row[0])) / 86400.0
                
                if (subutime > 14.0):
                    #print 'Re-Rank(' + str(subutime) + 'wks, ' + str(row[1]) + ')'
                    state.append('Re-Rank(' + str(subutime) + 'wks, ' + str(row[1]) + ')')
                elif (subutime > 3.0):
                    #print 'Re-Rank(' + str(int(math.floor(subutime))) + 'days, ' + str(row[1]) + ')'
                    state.append('Re-Rank(' + str(int(math.floor(subutime))) + 'days, ' + str(row[1]) + ')')
                else:
                    if (subrank == 0):
                        #print 'Stay'
                        state.append('Stay')
                    elif (subrank > 0):
                        #print str(subrank) + 'DOWN'
                        state.append(str(subrank) + 'DOWN')
                    else:
                        #print str(abs(subrank)) + 'UP'
                        state.append(str(abs(subrank)) + 'UP')
        status.append(state)
    status = [x for x in status if x]
    return status


# make Excel file
def makeExcel(filepath, site, category, period, datas, status):
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet()
    
    # title
    title_format = workbook.add_format()
    title_format.set_font_size(18) 
    title = ''
    if (site == 'yahoo'):
        title += u'Yahooショッピング '
    if (site == 'rakuten'):
        title += u'楽天市場 '
    if (period == 'daily'):
        title += u'デイリーランキング '
    if (period == 'weekly'):
        title += u'ウィークリーランキング '
    if (period == 'realtime'):
        title += u'リアルタイムランキング '
    title += u'(' + unicode(category, 'utf-8') + u')'
    worksheet.write('A1', title, title_format)
    
    merge_format = workbook.add_format({'align' : 'center', 'valign': 'vcenter'})
    merge_format.set_text_wrap()
    normal_format = workbook.add_format()
    normal_format.set_text_wrap()
    
    for i in range(2, 42):
        worksheet.set_row(i, 35)
    worksheet.set_column(1, len(datas)*2, 16)

    # rank
    for i in range(0, 10):
        worksheet.merge_range(i*4+2, 0, i*4+5, 0, i+1, merge_format)
    
    # datas
    for i, row in enumerate(datas):
        # timestamp
        worksheet.merge_range(1, i*2+1, 1, i*2+2, unicode(unixtime2date(row[0][2]), 'utf-8'), merge_format)
        
        for j, column in enumerate(row):
            # name
            worksheet.merge_range(j*4+2, i*2+1, j*4+2, i*2+2, column[6], merge_format)
            # image
            if (site == 'yahoo'):
                img_path = yahoo_imgpath + '/' + column[5] + '.jpg'
            if (site == 'rakuten'):
                img_path = rakuten_imgpath + '/' + column[5] + '.jpg'
            worksheet.insert_image(j*4+3, i*2+1, img_path, {'x_scale':0.61, 'y_scale':0.66})
            #worksheet.insert_image(j*4+3, i*2+1, img_path)
            worksheet.merge_range(j*4+3, i*2+1, j*4+5, i*2+1, 'image', merge_format)
            # store name
            worksheet.write(j*4+3, i*2+2, column[8], normal_format)
            # up, down, stay
            worksheet.write(j*4+4, i*2+2, status[i][j], normal_format)
            # price
            worksheet.write(j*4+5, i*2+2, int(column[9]), normal_format)
            
            pass
    
    workbook.close()



def main():
    argv = sys.argv
    argc = len(argv)
    if (argc < 5):
        print 'Usage: python %s savefile_path site_name master_id period timestamp0 timestamp1 ...' %argv[0]
        quit()
    
    # read csv
    category = readCategory(csvpath, argv[3])
    
    print 'save file path: ' + argv[1]
    print 'select website: ' + argv[2]
    print 'select master ID: ' + argv[3]
    print 'select category: ' + category
    print 'select period: ' + argv[4]
    print '------------------------------------'
 
    # get datas, calc rank status
    datas = [[]]
    for i, row in enumerate(argv):
        if (i > 4):
            print 'get datas: ' + row
            datas.append(getData(argv[2], argv[3], argv[4], row))
    datas = [x for x in datas if x]
    status = calcRanking(argv[2], argv[3], argv[4], datas)

    # make excel
    makeExcel(argv[1], argv[2], category, argv[4], datas, status)

if __name__ == '__main__':
    main()
