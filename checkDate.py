# -*- coding: utf-8 -*-

import sys
import sqlite3
from datetime import datetime


# DB(SQLite3) path
dbpath = './yahoo/csv/yahoo.sqlite3'


# unixtime --> date
def unixtime2date(utime):
    loc = datetime.fromtimestamp(float(utime))
    loc = loc.strftime('%Y-%m-%d (%a) %H:%M:%S')
    return loc


# check date all
def checkDate(site, master_id):
    connector = sqlite3.connect(dbpath)
    cursor = connector.cursor()
    sql = 'select distinct timestamp from ' + site + ' where master_id = "' + master_id + '" order by timestamp asc'
    cursor.execute(sql)
    
    result = cursor.fetchall()
    
    for row in result:
        print unixtime2date(unicode(row[0])) + ' : ' + unicode(row[0])

    cursor.close()
    connector.close()


def main():
    argv = sys.argv
    argc = len(argv)
    if (argc != 3):
        print 'Usage: python %s site_name master_id' %argv[0]
        quit()
    
    checkDate(argv[1], argv[2]);


if __name__ == '__main__':
    main()
