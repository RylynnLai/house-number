#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import re
import datetime
import xlsxwriter
import sys

reload(sys)
sys.setdefaultencoding('utf8')

class GroupMember(object):
    def __init__(self, root):
        """
          --初始化
        """
        self.building_numbers = ["1A", "1B", "3", "5", "6", "7A", "7B", "8", "9", "10", "11", "12", "13A", "13B", "15A", "15B", "东1A", "东1B", "东2", "东3A", "东3B"]
        self.root = root  # 文件树的根
        self.html_content = ""  # 待查询的群成员HTML片段
        self.members = []  # 群成员名字
        self.prefix_house_numbers = [] # 栋号码
        self.suffix_house_numbers = [] # 门号

        f = file(root + "1.html", "r")
        line = f.read()
        f.close()
        self.html_content = re.search(r'<!--BEGIN HD-->([\s\S]*?)<!--END HD-->', line, re.M|re.I).group(0)
        # print(self.html_content)

    @staticmethod
    def find_member_name(self):
        self.members = re.findall(r'(?<=UserName\)">)([\s\S]*?)(?=</p>)', self.html_content)
        
    @staticmethod
    def analyse_member_house_number(self):
        for name in self.members:
            #----------------------
            result = re.search(r'东(.*?)[a,b,A,B]', name, re.M|re.I)
            if result != None:
                if len(result.group(0)) > 0:
                    self.prefix_house_numbers.append(result.group(0))
                    continue
            #----------------------
            result = re.search(r'东(.*?)[一,二,三,四,五,六]', name, re.M|re.I)
            if result != None:
                if len(result.group(0)) > 0:
                    self.prefix_house_numbers.append(result.group(0))
                    continue
            #----------------------
            result = re.search(r'东(.*?)[0-9]', name, re.M|re.I)
            if result != None:
                if len(result.group(0)) > 0:
                    self.prefix_house_numbers.append(result.group(0))
                    continue
            #----------------------
            result = re.search(r'([0-9]{1,3}?)(?=[座,幢,栋,\-,—,_,#,~,.])', name, re.M|re.I)
            if result != None:
                if len(result.group(0)) > 0:
                    self.prefix_house_numbers.append(result.group(0))
                    continue
            #----------------------
            result = re.search(r'([0-9]*?)[a,b,A,B]', name, re.M|re.I)
            if result != None:
                if len(result.group(0)) > 0:
                    self.prefix_house_numbers.append(result.group(0))
                    continue
           
            self.prefix_house_numbers.append("can't find")


        for name in self.members:
            #----------------------
            result = re.search(r'\d{3,4}', name, re.M|re.I)
            if result != None:
                if len(result.group(0)) == 3:
                    self.suffix_house_numbers.append("0" + result.group(0))
                    continue
                if len(result.group(0)) == 4:
                    self.suffix_house_numbers.append(result.group(0))
                    continue

            self.suffix_house_numbers.append("can't find")

        # for index, name in enumerate(self.members):
        #     print(name + '------' + "栋号：" + self.prefix_house_numbers[index] + "-" + self.suffix_house_numbers[index])

    @staticmethod
    def sort_house(self, house):
        index_list = []
        for index, number in enumerate(self.prefix_house_numbers):
            if house.lower() == number.lower():
                index_list.append(index)
        return index_list

    @staticmethod
    def export_excel(self):
        workbook = xlsxwriter.Workbook('house_numbers.xlsx') # 建立文件
        for number in self.building_numbers:
            worksheet = workbook.add_worksheet(number)
            house_list = self.sort_house(self, number)
            unrecognized = list(house_list)
            unrecognizedString = ""

            for index in range(1,34):
                worksheet.write(index-1,0,("%d楼" % index))
                memberStrings = []
                for house_index, house in enumerate(house_list):
                    if self.suffix_house_numbers[house][0:2] == ("{0:02d}".format(index)):
                        memberStrings.append(self.members[house])
                        unrecognized[house_index] = 99999
                worksheet.write(index-1,1,"   |   ".join(memberStrings))
            
            worksheet.write(35,0,"未识别到")
            for index in unrecognized:
                if index != 99999:
                    unrecognizedString = unrecognizedString + self.members[index] + "   |   "
            worksheet.write(35,1,unrecognizedString)
        
        worksheet = workbook.add_worksheet("未识别")
        num = 0
        for index in range(0,len(self.members)):
            if self.prefix_house_numbers[index] == "can't find":
                worksheet.write(num,0,self.members[index])
                num += 1
                continue
            if self.suffix_house_numbers[index] == "can't find":
                worksheet.write(num,0,self.members[index])
                num += 1
                continue
        workbook.close()

if __name__ == "__main__":

    print datetime.datetime.now()
    obj = GroupMember('/Users/longzhou.lai/Documents/Python/wechatgroupmember/')
    obj.find_member_name(obj)
    obj.analyse_member_house_number(obj)
    obj.export_excel(obj)
    print datetime.datetime.now()