# coding=utf-8
# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import os
from PIL import Image
from PIL.ExifTags import TAGS
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# 在内存中创建一个workbook对象，而且会至少创建一个 worksheet
wb = Workbook()

# 获取当前活跃的worksheet,默认就是第一个worksheet
ws = wb.active

# 设置单元格的值，A1等于6(测试可知openpyxl的行和列编号从1开始计算)，B1等于7
ws.cell(row=1, column=1).value = 6
# ws.cell("B1").value = 7

# 从第2行开始，写入9行10列数据，值为对应的列序号A、B、C、D...
for row in range(2, 11):
    for col in range(1, 11):
        ws.cell(row=row, column=col).value = get_column_letter(col)

# 可以使用append插入一行数据
ws.append(["我", "你", "她"])

# 保存
wb.save(filename="/Users/ligs/Desktop/a.xlsx")

def get_exif_data(fname):
    ret = {}
    try:
        img = Image.open(fname)
        if hasattr(img, '_getexif'):
            exifinfo = img._getexif()
            if exifinfo != None:
                for tag, value in exifinfo.items():
                    decoded = TAGS.get(tag, tag)
                    ret[decoded] = value
            else:
                ret = 'no exif'
    except IOError:
        print 'IOERROR ' + fname
    return ret

def scan_photos(dir):
    for root,dirs,files in os.walk(dir):
        for file in files:
            print(root)
            print(os.path.join(root,file))

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print("Hi, {0}".format(name))  # Press ⌘F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')
    exif=get_exif_data('./photos/IMG_1478.JPG')
    if exif != "no exif":
        print(exif.get('Model'))
        print(exif.get('DateTime'))
        info=exif.get('GPSInfo')
        lat = info[2]
        lon = info[4]
        lat_deg = lat[0][0] * 1.0 / lat[0][1]
        lat_min = lat[1][0] * 1.0 / lat[1][1]
        lat_sec = lat[2][0] * 1.0 / lat[2][1]
        lon_deg = lon[0][0] * 1.0 / lon[0][1]
        lon_min = lon[1][0] * 1.0 / lon[1][1]
        lon_sec = lon[2][0] * 1.0 / lon[2][1]
        lat_decimal = lat_deg + lat_min / 60.0 + lat_sec / 3600.0
        lon_decimal = lon_deg + lon_min / 60.0 + lon_sec / 2600.0
        if info[1] is u'S':
            lat_decimal = -lat_decimal
        if info[3] is u'W':
            lon_decimal = -lon_decimal
        print(lon_decimal)
        print(lat_decimal)



# See PyCharm help at https://www.jetbrains.com/help/pycharm/
