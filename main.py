import os
import zipfile as zip
import rarfile as rar
import shutil
import re
import xlsxwriter


# make the Categorized file
# ------------------------------------------
dir1 = 'Categorized_HW2'
dir2 = 'AP-HW2'
dir3 = 'Students'

def make_file(dst):
    if os.path.exists(dst):
        shutil.rmtree(dst)
    os.makedirs(dst)

make_file(dir1)
make_file(dir2)
make_file(dir3)

# extract the name of questions from Q6.txt and make folders
# ------------------------------------------
with open('Q6.txt') as f:
    lines = f.readlines()
lines = [line.rstrip('\n') for line in open('Q6.txt') if line.startswith('Q')]
count = len(lines)
path_q = []
q = []
types = []
for i in lines:
    t = re.match('Q[0-9]', i)
    questions = t.group(0)
    q.append(questions)
    a = i.split(': ')
    a.pop(0)
    types += a
    path = dir1 + '\\' + questions
    path_q.append(path)
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path)

# creat excel list for wrong folder names
# ------------------------------------------

with zip.ZipFile('AP-HW2.zip', 'r') as d:
    d.extractall('AP-HW2')
folderr = os.path.join(os.getcwd(),'AP-HW2')
workbook = xlsxwriter.Workbook('error.xlsx')
worksheet = workbook.add_worksheet()
filess = [f for f in os.listdir(folderr) if f.endswith('_assignsubmission_file_')]
student_folder = []
name = []
row = 2
worksheet.write(0, 0, 'Names')
worksheet.write(0, 1, 'Student No.')
worksheet.write(0, 2, '-10%')
worksheet.write(0, 3, 'Q folders')
worksheet.write(0, 4, 'Check!')

# loop for every folder in main folder
for f in filess:
    t = f.split('_')[0]
    worksheet.write(row, 0, t)
    name.append(t)
    d = os.path.join(os.getcwd(), dir3)
    make_file(d + '\\' + t)
    x = folderr + '\\' + f
    student_folder.append(x)
    row += 1
row = 2

for i in range (0, len(types)):
    types[i] = re.sub(' ', '', types[i])

d = os.path.join(os.getcwd(), dir3)

# copying file function
#--------------------------------
def copy_file(Q_num, cnt, list, i, na, student_nums):
    dst = os.path.join(os.getcwd(), path_q[Q_num])
    x = dst + '\\' + student_nums[cnt] + '_' + list[i].split('/')[len(list[i].split('/')) - 1]
    shutil.copyfile(na + '//' + list[i], x)

def Unzip(y, cnt, name, q, types, worksheet, student_nums, row, bool):
    if bool:
        with zip.ZipFile(y, 'r') as s:
            s.extractall('Students' + '//' + name[cnt])
    else:
        with rar.RarFile(y, 'r') as s:
            s.extractall('Students' + '//' + name[cnt])
    list = s.namelist()
    na = 'Students' + '//' + name[cnt]
    print('Processing {} ...'.format(student_nums[cnt]))
    for i in range(0, len(list)):
        for j in range(0, len(q)):
            if (len(types[j].split(',')) > 1):
                a = types[j].split(',')
                for z in range(0, len(a)):
                    reg = q[j] + r'.*\.' + a[z] + '$'
                    dd = []
                    dd += q[j]
                    # dd is 'Q' , '2' ...
                    reg2 = dd[0] + r'.{1}' + dd[1] + r'.*\.' + a[z] + '$'
                    reg3 = 'q' + r'.{1}' + dd[1] + '.*\.' + a[z] + '$'
                    reg4 = r'\D' + dd[1] + '/[^/]+' + '.' + a[z] + '$'
                    if re.search(reg, list[i]) != None:
                        copy_file(j, cnt, list, i, na, student_nums)
                    elif re.search(reg2, list[i]) != None:
                        worksheet.write(row, 3, 'wrong2')
                        copy_file(j, cnt, list, i, na, student_nums)
                    elif re.search(reg3, list[i]) != None:
                        worksheet.write(row, 3, 'wrong3')
                        copy_file(j, cnt, list, i, na, student_nums)
                    elif re.search(reg4, list[i]) != None:
                        worksheet.write(row, 3, 'wrong4')
                        copy_file(j, cnt, list, i, na, student_nums)
            elif (len(types[j].split(',')) == 1):
                reg = q[j] + r'.*\.' + types[j] + '$'
                dd = []
                dd += q[j]
                reg2 = dd[0] + r'.{1}' + dd[1] + r'.*\.' + types[j] + '$'
                reg3 = 'q' + r'.{1}' + dd[1] + '.*\.' + types[j] + '$'
                # reg4 = dd[1] + '/[.*^/]' + types[j] + '$'
                reg4 = r'\D' + dd[1] + '/[^/]+' + '.' + types[j] + '$'
                # HW2/Q3/Queue.cpp
                # HW2/Q2/func.cpp
                if re.search(reg, list[i]) != None:
                    copy_file(j, cnt, list, i, na, student_nums)
                elif re.search(reg2, list[i]) != None:
                    worksheet.write(row, 3, 'wrong2')
                    copy_file(j, cnt, list, i, na, student_nums)
                elif re.search(reg3, list[i]) != None:
                    worksheet.write(row, 3, 'wrong3')
                    copy_file(j, cnt, list, i, na, student_nums)
                elif re.search(reg4, list[i]) != None:
                    worksheet.write(row, 3, 'wrong4')
                    copy_file(j, cnt, list, i, na, student_nums)

student_nums = []
# ---------------------------------------
# the path of every folder in main folder
cnt = 0
for f in student_folder:
    # g is the name of zip or rar folder in every student's file
    g = os.listdir(f)[0]
    # matches the right form of folder name
    # -------------------------------------------------
    matching = re.match('AP-HW2-[0-9]{7}', g)
    num = re.search('[0-9]{7}', g)
    if num == None:
        student_nums.append(name[cnt])
    else:
        worksheet.write(row, 1, num.group(0))
        student_nums.append(num.group(0))
    if matching == None:
        worksheet.write(row, 2, 'Wrong filename')

    list = []

    # path to zip or rar files zip.is_zipfile(y)
    y = f + '\\' + g
    what = True

    if zip.is_zipfile(y):
        bool = True
        Unzip(y, cnt, name, q, types, worksheet, student_nums, row, bool)

    # IN CASE UNRAR MODULE DONT WORK --- NOT COMMENT
    # -------------------------------------------
    elif g.endswith('.zip'):
        worksheet.write(row, 4, 'Wrote zip as rar')
        print('Processing {} ...'.format(name[cnt]))
    elif g.endswith('.rar'):
        worksheet.write(row, 4, 'rar file')
        print('Processing {} ...'.format(name[cnt]))
    else:
        print('Processing {} ...'.format(name[cnt]))
        worksheet.write(row, 4, 'simple folder with zipfile')
        while(what):
            oo = os.listdir(y)
            if len(oo) == 1:
                paaath = y + '\\' + oo[0]
                if zip.is_zipfile(paaath):
                    bool = True
                    Unzip(paaath,cnt, name, q, types, worksheet, student_nums, row, bool)
                    what = False
            else:
                worksheet.write(row, 4, 'Malum nist chie!!')

    # IN CASE UNRAR FUNCTION WORKS ---- UNCOMMENT FOR RAR FILES
    # ---------------------------------------------------


    # elif rar.is_rarfile(y):
    #     bool = False
    #     if g.endswith('.zip'):
    #         worksheet.write(row, 4, 'Wrote zip as rar')
    #         Unzip(y, cnt, name, q, types, worksheet, student_nums, row, bool)
    #     else:
    #         worksheet.write(row, 4, 'rar file')
    #         Unzip(y, cnt, name, q, types, worksheet, student_nums, row, bool)

    # # if the submitted folder is a simple folder(not zip or rar)
    # else:
    #     worksheet.write(row, 4, 'simple folder with zipfile')
    #     while(what):
    #         oo = os.listdir(y)
    #         if len(oo) == 1:
    #             paaath = y + '\\' + oo[0]
    #             if zip.is_zipfile(paaath):
    #                 bool = True
    #                 Unzip(paaath,cnt, name, q, types, worksheet, student_nums, row, bool)
    #                 what = False
    #             # elif rar.is_rarfile(paaath):
    #             #     bool = False
    #             #     Unzip(paaath, cnt, name, q, types, worksheet, student_nums, row, bool)
    #             #     what = False

    row += 1
    cnt += 1
workbook.close()

