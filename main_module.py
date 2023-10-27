#encoding=utf-8
#author: blue16（blue16.cn）
#date: 2023-10-7
#summary: 通过无课表生成值班表
import tabula
import csv
import openpyxl
import os
import random
#读取指定周，指定时间区域内的所有人，返回纯文本信息
def read_section(sheet,week_num,section_num):
    week_list=['C','D','E','F','G','H','I']
    if(section_num==1):
        cell_id_1=week_list[week_num-1]+"2"
        cell_id_2=week_list[week_num-1]+"3"
    if(section_num==2):
        cell_id_1=week_list[week_num-1]+"4"
        cell_id_2=week_list[week_num-1]+"5"
    if(section_num==3):
        cell_id_1=week_list[week_num-1]+"6"
        cell_id_2=week_list[week_num-1]+"7"
    if(section_num==4):
        cell_id_1=week_list[week_num-1]+"8"
        cell_id_2=week_list[week_num-1]+"9"
    #防止none出来搞事情
    if(sheet[cell_id_1].value!=None):
        str1=sheet[cell_id_1].value
    else:
        str1=""
    if(sheet[cell_id_2].value!=None):
        str2=sheet[cell_id_2].value
    else:
        str2=""
    return str1+str2
#print(read_section(wb_sheet,1,2))

def random_member(member_string):#随机取出一个人，已经弃用
    #先分析出有多少人,切片
    if(member_string==""):
        return ""
    strtmp=member_string.split("，")
    #直接返回index的随机数值
    #print(strtmp)
    index=random.randint(0,len(strtmp)-2)
    #print("抽出的int=",index)
    strtmp2=strtmp[index].split(" ")
    return strtmp2[1]

def get_members(member_string):#获取当前string中的所有人，返回一个list
    strtmp=member_string.split("，")
    ret=[]
    for temp1 in strtmp:
        temp2=temp1.split(" ")
        if(len(temp2)==2):
            b=temp2[1] in ret
            if(temp2[1]!="" and b==False):
                ret.append(temp2[1])
    return ret



#print("Res:",random_member(read_section(wb_sheet,1,1)))
def write_record(sheet,week_num,section_num,member_num,name):
    week_list=['B','C','D','E','F']
    cell_id=week_list[week_num-1]+str(3*section_num+member_num-1)
    #print(cell_id,name)
    sheet[cell_id]=name
    #print(sheet[cell_id].value)


def generate():
    list_finished=[]#记录已经安排的人员
    #如果人都点完了怎么办？
    week_list=['C','D','E','F','G','H','I']
    for week in range(1,6):#左闭右开！！！
        for section in range(1,5):
            #获取人员名单（当前section），随机抽取，然后删除
            section_string=read_section(wb_sheet,week,section)
            members=get_members(section_string)
            k=0#表示已写入人数,这里有问题？？？
            while k<3 and len(members)>0:
                member_name=members[random.randint(0,len(members)-1)]
                write_record(wb_example_sheet,week,section,k,member_name)
                k=k+1
                #print(member_name)
                if(member_name==""):
                    k=k+1
                    continue
                #死循环原因：找不到人了反复
                temp = member_name in list_finished
                if(temp==False):
                    list_finished.append(member_name)
                    #print("this")
                    #按序列写入格子
                    k=k+1
                    write_record(wb_example_sheet,week,section,k,member_name)
                print("section",section_string)
#打开无课表，用于读取数据
wb=openpyxl.load_workbook('Input.xlsx')
wb_sheet=wb['Sheet']
#获取幸运数字（狗头）
#random_seed=input("输入一个幸运数字")
#random_seed=99
#random.seed(random_seed)
#打开模板，准备向模板写入数据
wb_example=openpyxl.load_workbook('example.xlsx')
wb_example_sheet=wb_example['Sheet']
generate()
#write_record(wb_example_sheet,1,1,2,"测试写入")
wb_example.save("output.xlsx")