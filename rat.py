#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import MySQLdb
import xlwt
import sys
import argparse
from time import gmtime, strftime

dtnow = str(strftime("%Y-%m-%d", gmtime()))

# Create connection to database
try:
    # Create connection to database
    db = MySQLdb.connect(host='rat.net.rts',port=3306, passwd='',db='racktables',user='rack',password='',charset="utf8")
except MySQLdb.Error:
    e = sys.exc_info()[1]
    print("Error %d: %s" % (e.args[0],e.args[1]))
    sys.exit(1)

cur = db.cursor()

parser = argparse.ArgumentParser(
                    prog='RAT.py',
                    description='Downloads data from RackTables Database to Excel file',
                    epilog='''Example: rat.py -c CLIENT_CODE or rat.py -a ASSET_TAG.
                    set ASSET_TAG to "ALL" to download all data from the base''')

parser.add_argument('-c', '--client', help='CLIENT_CODE', required=False)
parser.add_argument('-a', '--assettag', help='ASSET_TAG', required=False)


class Device():
    def __init__(self, line1):
        self.line1 = line1
        
    def id(self):
        # id of the device
        return self.line1[0]
    
    def name(self):
        # Name of the device
        return self.line1[1]

    def tag(self):
        # This is Asset Tag field
        return self.line1[2]

    def rack_unit(self):
        try:
            cur.execute(f"select rack_id,unit_no from RackSpace where object_id like {self.id()};")
            rack_all = cur.fetchall()
            rack = rack_all[0][0]
            units = tuple(set([i[1] for i in rack_all]))
            if len(units) > 1:
                units = str(units[0]) + '-' + str(units[-1])
            else:
                units = str(units[0])
            return rack, units
        except:
            rack = 0
            units = 0
            return rack, units
    
    def location(self):
        rack, units = self.rack_unit()
        if rack == 0:
            rack_n = room = office = 0
        else:
            cur.execute(f"select id,name,row_id,row_name,location_id,location_name from Rack where id like {rack};")
            location = cur.fetchall()
            rack_n = location[0][1]
            room = location[0][3]
            office = location[0][5]
        return rack_n, room, office
    
    def serial(self):
        device_id = self.id()
        cur.execute(f"select string_value from AttributeValue where object_id like {device_id} and attr_id like 1;")
        serial = cur.fetchall()
        if serial == ():
            serial = '0'
        else:
            serial = serial[0][0]
        return serial
    
    def hw_type(self):
        device_id = self.id()
        cur.execute(f"select uint_value from AttributeValue where object_id like {device_id} and attr_id like 2;")
        hw_id = cur.fetchall()
        if hw_id == ():
            hw_type = '0'
            return hw_type
        else:
            hw_id = hw_id[0][0]
            cur.execute(f"select dict_value from Dictionary where dict_key like {hw_id};")
            hw_type = cur.fetchall()[0][0].split('|')[0].strip('[').replace("%GPASS%","")
            return hw_type


wb = xlwt.Workbook()
bold = xlwt.easyxf('font: bold 1')
sheet1 = wb.add_sheet('Sheet 1')
sheet1.col(0).width = 256 * 30
sheet1.col(1).width = 256 * 20
sheet1.col(2).width = 256 * 30
sheet1.col(3).width = 256 * 40
sheet1.col(4).width = 256 * 36


sheet1.write(0, 0, 'Имя устройства',bold)
sheet1.write(0, 1, 'Серийный номер',bold)
sheet1.write(0, 2, 'Инвентарный номер', bold)
sheet1.write(0, 3, 'HW type',bold)
sheet1.write(0, 4, 'Площадка',bold)
sheet1.write(0, 5, 'Помещение',bold)
sheet1.write(0, 6, 'Стойка',bold)
sheet1.write(0, 7, 'Юнит',bold)

args = parser.parse_args()

if args.client:
    tag = args.client
elif args.assettag:
    tag = args.assettag
else:
    print(parser.epilog)
    exit()


if tag == 'ALL':
    cur.execute('select id,name,asset_no,label,comment from Object')
elif tag != 'ALL' and args.client:
    cur.execute(f'select id,name,asset_no,label,comment from Object where name like "%{tag}%"')
elif tag != 'ALL' and args.assettag:
    cur.execute(f'select id,name,asset_no,label,comment from Object where asset_no like "%{tag}%"')

sql = cur.fetchall()


for i in range(len(sql)):
    device = Device(sql[i])
    rack, units = device.rack_unit()
    rack_n, room, office = device.location()
    sheet1.write(i+1, 0, device.name())
    sheet1.write(i+1, 1, device.serial())
    sheet1.write(i+1, 2, device.tag())
    sheet1.write(i+1, 3, device.hw_type())
    sheet1.write(i+1, 4, office)
    sheet1.write(i+1, 5, room)
    sheet1.write(i+1, 6, rack_n)
    sheet1.write(i+1, 7, units)

wb.save(f'{tag} {dtnow}.xls') 

