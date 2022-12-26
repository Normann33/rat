#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import MySQLdb
import xlwt
import sys

# Create connection to database
try:
    # Create connection to database
    db = MySQLdb.connect(host='rat.net.rts',port=3306, passwd='',db='racktables',user='rack', charset="utf8")
except MySQLdb.Error:
    e = sys.exc_info()[1]
    print("Error %d: %s" % (e.args[0],e.args[1]))
    sys.exit(1)

cur = db.cursor()
client = sys.argv[1]

def get_device_id(client):
    if client == 'ALL':
        cur.execute('select id,name,asset_no,label,comment from Object')
    else:
        cur.execute('select id,name,label,comment from Object where name like "%{}%"'.format(client))
    sql = cur.fetchall()
    n = len(sql)
    device_id_list = []
    for i in range(n):
        device_id_list.append(sql[i][0])
    return device_id_list
def get_device_name(client):
    if client == 'ALL':
        cur.execute('select id,name,asset_no,label,comment from Object')
    else:
        cur.execute('select id,name,label,comment from Object where name like "%{}%"'.format(client))
    sql = cur.fetchall()
    n = len(sql)
    device_name_list = []
    for i in range(n):
        device_name_list.append(sql[i][1])
    return device_name_list
def get_asset(client):
    if client == 'ALL':
        cur.execute('select id,name,asset_no,label,comment from Object')
    else:
        cur.execute('select id,name,label,comment from Object where name like "%{}%"'.format(client))
    sql = cur.fetchall()
    n = len(sql)
    device_asset_list = []
    for i in range(n):
        device_asset_list.append(sql[i][2])
    return device_asset_list

def get_rack_id(device_id):
    try:
        cur.execute("select rack_id from RackSpace where atom like 'front' and object_id like {};".format(device_id))
        rack_all = cur.fetchall()
        rack = rack_all[0][0]
        return rack
    except:
        pass
def get_unit(device_id):
    try:
        cur.execute("select unit_no from RackSpace where atom like 'front' and object_id like {};".format(device_id))
        rack_all = cur.fetchall()
        n = len(rack_all)
        units = [] 
        for i in range(n):
            units.append(str(rack_all[i][0]))
        return units
    except:
        pass
def get_location(rack_id):
    if rack_id == None:
        rack_n = '0'
        room = '0'
        office = '0'
    else:
        cur.execute("select id,name,row_id,row_name,location_id,location_name from Rack where id like {};".format(rack_id))
        location = cur.fetchall()
        rack_n = location[0][1]
        room = location[0][3]
        office = location[0][5]
    return rack_n,room,office
def get_serial(device_id):
    cur.execute("select string_value from AttributeValue where object_id like {} and attr_id like 1;".format(device_id))
    serial = cur.fetchall()
    if serial == ():
        serial = '0'
    else:
        serial = serial[0][0]
    return serial
def get_hw_type(device_id):
    cur.execute("select uint_value from AttributeValue where object_id like {} and attr_id like 2;".format(device_id))
    hw_id = cur.fetchall()
    if hw_id == ():
        hw_type = '0'
        return hw_type
    else:
        hw_id = hw_id[0][0]
        cur.execute("select dict_value from Dictionary where dict_key like {};".format(hw_id))
        hw_type = cur.fetchall()[0][0].split('|')[0].strip('[').replace("%GPASS%","")
        return hw_type


device_id_list = get_device_id(client)
device_name_list = get_device_name(client)
device_asset_list = get_asset(client)

wb = xlwt.Workbook()
bold = xlwt.easyxf('font: bold 1')
sheet1 = wb.add_sheet('Sheet 1')
sheet1.col(0).width = 256 * 20
sheet1.col(1).width = 256 * 20
sheet1.col(2).width = 256 * 60
sheet1.col(3).width = 256 * 30
sheet1.col(4).width = 256 * 36


sheet1.write(0, 0, 'Имя устройства',bold)
sheet1.write(0, 1, 'Серийный номер',bold)
sheet1.write(0, 2, 'Инвентарный номер', bold)
sheet1.write(0, 3, 'HW type',bold)
sheet1.write(0, 4, 'Площадка',bold)
sheet1.write(0, 5, 'Помещение',bold)
sheet1.write(0, 6, 'Стойка',bold)
sheet1.write(0, 7, 'Юнит',bold)


for i in range(len(device_id_list)):
    try:
        rack_n,room,office = get_location(get_rack_id(device_id_list[i]))
        serial = get_serial(device_id_list[i])
        hw_type = get_hw_type(device_id_list[i])
        units = ','.join(get_unit(device_id_list[i]))
    except:
        print("Error")
    sheet1.write(i+1, 0, device_name_list[i])
    sheet1.write(i+1, 1, serial)
    sheet1.write(i+1, 2, device_asset_list[i])
    sheet1.write(i+1, 3, hw_type)
    sheet1.write(i+1, 4, office)
    sheet1.write(i+1, 5, room)
    sheet1.write(i+1, 6, rack_n)
    sheet1.write(i+1, 7, units)



wb.save('{}.xls'.format(client)) 
