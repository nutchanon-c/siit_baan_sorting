
from cmath import inf
import math
import re
import gspread
import json
import pymysql
import os
import PySimpleGUI as sg
import random
import sys
from pymysql import OperationalError
from pymysql import ProgrammingError
from setuptools import Command

# from tkinter import font
# import tkinter
# root = tkinter.Tk()
# fonts = list(font.families())
# fonts.sort()
# root.destroy()



dir_path = os.path.dirname(os.path.realpath(__file__))


if getattr(sys, 'frozen', False):
    dir_path = os.path.dirname(sys.executable)
elif __file__:
    dir_path = os.path.dirname(__file__)
currentDir = dir_path.replace('\\', '/')

# read an extract data from config.json
try:
    with open(currentDir + "/config.json", 'r', encoding="utf8") as f:
        data = json.loads(f.read())
        dbHostName = data['dbHostName']
        dbUserName = data['dbUserName']
        dbPassword = data['dbPassword']
        dbName = data['dbName']
        worksheetName = data['worksheetName']
        sheetName = data['sheetName']
        textFileName = data['textFileName']
        serviceAccountFile = data['serviceAccount']
        backgroundColor = data['backgroundColor']
        textColor = data['textColor']
        labelColor = data['labelColor']
except:
    sg.Popup('Config file not found. Please create a config file.')
    sys.exit()

# textFileName = 'test1.txt'
# worksheetName = 'รายชื่อกลุ่มน้องกิจกรรมเอสไอไทบ้าน'
# sheetName = 'Sheet1'

# dbHostName = 'localhost'
# dbUserName = 'root'
# dbPassword = ''
# dbName = 'baan_sort'
    


def updateData():
    # loads the text file
    try:
        with open(currentDir + '/' + textFileName, 'r', encoding="utf-8") as f:
            a = json.loads(f.read())
    except:
        f = open(currentDir + '/' + textFileName, 'w', encoding="utf-8")
        f.close()
        a = []


    # connects database
    db = pymysql.connect(host=dbHostName,
                            user=dbUserName,
                            password=dbPassword,
                            db=dbName,
                            charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    cursor = db.cursor()
    # fetching data
    cursor.execute("SELECT * FROM student")
    result = cursor.fetchall()
    x = [list(i.values()) for i in result] # convert to from {} to list

    x_cut = [r[1:] for r in x] # removing the id which is PK


    try:
        sa = gspread.service_account(filename= currentDir + "/" + serviceAccountFile)
    except:
        sg.Popup('Service account file not found')
        sys.exit()
    sh = sa.open(worksheetName)

    wks = sh.worksheet(sheetName)
    # print('Rows: ', wks.row_count)
    # print('Col: ', wks.col_count)

    g = wks.get_all_values()
    g = g[1:]
    for record in g:
        record[0] = int(record[0])
            


    add_count = 0
    for record in g:
        if record not in x_cut:
            s = ','.join(["'" + str(x) + "'" for x in record])
            command = "INSERT INTO student (groupNo, fullName, nickName, sex, line_id, phone, bloodType, foodAllergies, dietRestrictions, drugAllergies, otherAllergies, congenitalDiseases, size,emergency) VALUES(" + s + ")" 
            cursor.execute(command)
            db.commit()
            add_count += 1
            a.append(command + '\n') 
    with open(currentDir + '/' + textFileName, 'a', encoding="utf-8") as f:
        # f.write(json.dumps(a))
        for lines in a:
            f.write(lines)
        f.close()
    if add_count == 0:
        print('No new records added')
        return 0
    else:
        print(f'Added {add_count} new records')
        return add_count

def getGroupNumbers():
    # connects database
    db = pymysql.connect(host=dbHostName,
                            user=dbUserName,
                            password=dbPassword,
                            db=dbName,
                            charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    cursor = db.cursor()
    # command = "SELECT DISTINCT groupNo FROM student"
    command = "select distinct s.groupNo from student s where s.groupNo not in(select groupNo from sorted_table)"
    cursor.execute(command)
    result = cursor.fetchall()
    return [x['groupNo'] for x in result]

def getGroupMembers(group):
    db = pymysql.connect(host=dbHostName,
                            user=dbUserName,
                            password=dbPassword,
                            db=dbName,
                            charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    cursor = db.cursor()
    command = "SELECT fullName from student WHERE groupNo = " + str(group)
    
    cursor.execute(command)
    result = cursor.fetchall()
    return [x['fullName'] for x in result]


def insertToSorted(groupNo, baanNo):
    db = pymysql.connect(host=dbHostName,
                            user=dbUserName,
                            password=dbPassword,
                            db=dbName,
                            charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    cursor = db.cursor()
    command = "INSERT INTO sorted_table (groupNo, baan) VALUES(" + str(groupNo) + "," + str(baanNo) + ")"
    cursor.execute(command)
    db.commit()

# code from https://stackoverflow.com/questions/4408714/execute-sql-file-with-python-mysqldb
def exec_sql_file(cursor, sql_file):
    print ("\n[INFO] Executing SQL script file: '%s'" % (sql_file))
    statement = ""

    for line in open(sql_file):
        if re.match(r'--', line):  # ignore sql comment lines
            continue
        if not re.search(r';$', line):  # keep appending lines that don't end in ';'
            statement = statement + line
        else:  # when you get a line ending in ';' then exec statement and reset for next statement
            statement = statement + line
            #print "\n\n[DEBUG] Executing SQL statement:\n%s" % (statement)
            try:
                cursor.execute(statement)
            except (OperationalError, ProgrammingError) as e:
                print ("[WARN] MySQLError during execute statement \n\tArgs: '%s'" % (str(e.args)))

            statement = ""

def createDB():
    db = pymysql.connect(host=dbHostName,
                            user=dbUserName,
                            password=dbPassword,
                            charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    cursor = db.cursor()

    command = f"CREATE DATABASE IF NOT EXISTS `{dbName}` DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci;"
    cursor.execute(command)
    db.commit()
    command = f"USE {dbName};"
    cursor.execute(command)
    db.commit()



    exec_sql_file(cursor, currentDir + "/baan_sort.sql")

    db.commit()
    print('DATABASE CREATED')


def checkDatabaseExistence():
    command = f"SELECT SCHEMA_NAME FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = '{dbName}';"
    db = pymysql.connect(host=dbHostName,
                            user=dbUserName,
                            password=dbPassword,
                            charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    cursor = db.cursor()
    cursor.execute(command)
    result = cursor.fetchall()
    if len(result) == 0:
        createDB()
        

def complementaryColor(my_hex):
    """Returns complementary RGB color

    Example:
    >>>complementaryColor('FFFFFF')
    '000000'
    """
    if my_hex[0] == '#':
        my_hex = my_hex[1:]
    rgb = (my_hex[0:2], my_hex[2:4], my_hex[4:6])
    comp = ['%02X' % (255 - int(a, 16)) for a in rgb]
    return ''.join(comp)


def baan_selected(groupNo, baanNo):
    db = pymysql.connect(host=dbHostName,
                            user=dbUserName,
                            password=dbPassword,
                            db=dbName,
                            charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    cursor = db.cursor()

    # add data to worksheet
    sa = gspread.service_account(filename= currentDir + "/" + serviceAccountFile)
    sh = sa.open(worksheetName)

    wks = sh.worksheet(f'บ้าน {baanNo}')
    g = wks.get_all_values()
    g = g[1:]
    # print('test sheet', g)
    # print('sheet length', len(g))
    
    numRows = len(g) + 1 # offset from header

    command = f"SELECT * FROM student WHERE groupNo = {groupNo};"
    cursor.execute(command)
    result = cursor.fetchall()
    dataDict = [x for x in result]        
    dataList = [list(a.values()) for a in dataDict]
    dataListCut = [x[1:] for x in dataList] # remove PK
    # print(dataListCut)
    # update sheet A numRows+1 to M numRows + 1 + len(dataList)
    wks.update(f'A{numRows+1}:N{numRows + 1 + len(dataList)}', dataListCut)
    # wks.update('A6:B6', [[1,2]])
    # print(len(dataList[0]))





def random_baan(groupNo):
    db = pymysql.connect(host=dbHostName,
                            user=dbUserName,
                            password=dbPassword,
                            db=dbName,
                            charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    cursor = db.cursor()
    command = f"SELECT COUNT(*) as size FROM student WHERE groupNo = {groupNo}"
    cursor.execute(command)
    result = cursor.fetchall()
    groupSize = result[0]['size'] # group size
    print('groupSize', groupSize)


    command = f"SELECT COUNT(*) as boys FROM STUDENT WHERE groupNo = {groupNo} AND sex = 'Male';"
    cursor.execute(command)
    result = cursor.fetchall()
    numBoys = result[0]['boys'] # number of boys in group
    print('numBoys', numBoys)


    stat = float('inf')
    choices = []    

    baan_people = {}
    baan_boys = {}

    for i in range(1, 12):
        # get total amount of people
        command = f"SELECT COUNT(*) totalPpl from student WHERE groupNo IN (SELECT groupNo from sorted_table WHERE baan = {i});"
        cursor.execute(command)
        result = cursor.fetchall()
        baan_people[i] = result[0]['totalPpl']

        # get total amount of boys
        command = f"SELECT COUNT(*) as boysAmount from student WHERE groupNo IN (SELECT groupNo from sorted_table WHERE baan = {i}) AND sex = 'Male';"
        cursor.execute(command)
        result  = cursor.fetchall()
        baan_boys[i] = result[0]['boysAmount']

    for i in range(1, 12):
        sdBoys = 0
        sdPeople = 0
        cvBoys = 0
        cvPeople = 0

        # calculate
        baan_people[i] += groupSize
        baan_boys[i] += numBoys


        totalSortedBoys = sum(baan_boys.values())
        totalSortedPeople = sum(baan_people.values())

        avgBoys = totalSortedBoys / 11
        avgPeople = totalSortedPeople / 11

        sdBoys = math.sqrt((sum([x**2 for x in baan_boys.values()]) / 11) - (avgBoys**2))
        sdPeople = math.sqrt((sum([x**2 for x in baan_people.values()]) / 11) - (avgPeople**2))


        if avgBoys == 0:
            cvBoys = 0
        else:
            cvBoys = sdBoys / avgBoys
        cvPeople = sdPeople / avgPeople

        indicator = (2 * cvBoys) + cvPeople
        if indicator < stat:
            stat = indicator
            choices = [i]
        elif indicator == stat:
            choices.append(i)
    
        baan_people[i] -= groupSize
        baan_boys[i] -= numBoys
    # print('choices: ', choices)
    return random.choice(choices)








def main():
    num = 0
    res = -1
    colors = {
        "backgroundColor" :backgroundColor,
        "labelColor" : labelColor,
        "memberListColor" : textColor
    }
 
    try:
        checkDatabaseExistence()
    except:
        sg.Popup('Database Error', 'Please check your database connection', icon=currentDir + '/icon.ico')
        sys.exit()     
    def resetWindow():
        window['-COMBO-'].update(values=getGroupNumbers())
        window['-MEMBERLABEL-'].update("Please select a group first")
        window['-MEMBERLIST-'].update("")
        window['-RANDOMNUMBER-'].update('')
        window['-RANDOMBUTTON-'].update(visible=True)
        window['-RESULTOKBUTTON-'].update(visible=False)
        
    r = False
    selected = False
    selected_group = 0
    try:
        groupNumberList = getGroupNumbers()
    except:
        sg.Popup('Database Error', 'Please check your database connection', icon=currentDir + '/icon.ico')
        sys.exit()        
    layout = [   
            
        [
            
            [
                sg.Button('Update Data'),
            ], 
            [
                sg.Text('', key='-UPDATEDATA-', size=(30,1),visible=True, justification='c', background_color=colors['backgroundColor'], text_color='black'),   
            ]
        ],
        [
            [
                [sg.Combo(values=groupNumberList, key='-COMBO-', default_value='Select Group', change_submits=True, readonly=True, size=(20,1))],
                [sg.Text('Please select a group first', key='-MEMBERLABEL-', font=('Arial', 20), visible=True, text_color=colors['labelColor'], background_color=colors['backgroundColor'])],
                [sg.Text('', key='-MEMBERLIST-', font=('Arial', 25), size=(25,4), justification='c', text_color=colors['memberListColor'], background_color=colors['backgroundColor'])],
            ]
        ],
        [
            sg.Text(text='', key='-RANDOMNUMBER-', font=('Arial', 100), background_color=colors['backgroundColor'], size=(5,1), justification='c', text_color='#fe00ff'),
        ],
        [
            
            sg.Button('Random', key='-RANDOMBUTTON-', visible= True, size=(20, 2), button_color='#9907fe'), 
            sg.Button('Stop', key='-STOPRANDOMBUTTON-', visible= False, size=(20, 2), button_color='#fe00ff'),
            sg.Button('OK', key='-RESULTOKBUTTON-', visible= False, size=(20, 2), button_color='#2583ff'),
        ],
        [
            sg.Stretch(background_color=colors['backgroundColor']),
            
        ],
        [
            sg.Text('Made by Nutchanon Charnwutiwong',expand_x=True, justification='r', background_color=colors['backgroundColor'],text_color="#"+complementaryColor(backgroundColor[1:]))
        ]

        
        
        ]
    window = sg.Window(title='Baan Sorting', layout=layout, resizable=False,size=(800,600), element_justification='c', font=('Arial', 11), icon=currentDir + '/icon.ico',titlebar_icon=currentDir + '/icon.ico', background_color=colors['backgroundColor'])
    # cnt = 0
    # a = 0
    while True:
        # if r:
        #     cnt += 1
        #     if cnt == 10:
                
        #         cnt = 0
        #         a += 1
        #         print(a)
        #     if a == 3:
        #         r = False
                
        event, values = window.read(timeout=50)
        # continuous random number
        if event == 'Update Data':
            rec = updateData()
            if rec == 0:
                window['-UPDATEDATA-'].update('No new records added')
            else:
                window['-UPDATEDATA-'].update(f'{rec} new records added')
            # sg.Popup('Data updated')
            groupNumberList = getGroupNumbers()
            window['-UPDATEDATA-'].update(visible=True)
            window['-COMBO-'].update(values=groupNumberList)
            

        if event == sg.WIN_CLOSED:
            break
        if event == '-COMBO-':
            window['-UPDATEDATA-'].update('')
            val = values['-COMBO-']
            selected_group = int(val)
            selected = True
            print(selected_group)

            members = getGroupMembers(val)
            member_list_string = '\n'.join(members)
            # print(members)
            window['-MEMBERLABEL-'].update(f'Group {val} Members:')
            window['-MEMBERLIST-'].update(member_list_string)
        if event == '-RANDOMBUTTON-' and selected:
            r = True
            res = random_baan(selected_group)
            window['-RANDOMBUTTON-'].update(visible=False)
            window['-STOPRANDOMBUTTON-'].update(visible = True)
            window['-COMBO-'].update(disabled=True)

        if event == '-STOPRANDOMBUTTON-':
            r = False
            num = res
            window['-RANDOMNUMBER-'].update(num)            
            window['-STOPRANDOMBUTTON-'].update(visible = False)
            window['-RANDOMBUTTON-'].update(visible=False)
            window['-RESULTOKBUTTON-'].update(visible=True)    
            selected = False
            
        if event == '-RESULTOKBUTTON-':
            insertToSorted(selected_group, res)
            baan_selected(selected_group, res)
            resetWindow()
            window['-COMBO-'].update(disabled=False)
            selected = False
            r = False
        if r:
            num = random.randint(1,11)
            window['-RANDOMNUMBER-'].update(num)
            

if __name__ == '__main__':
    # fonts = sg.Text.fonts_installed_list()
    # print(fonts)
    main()
