from distutils.spawn import spawn
from inspect import getcallargs
import os
from readline import get_current_history_length
import sys
from turtle import home
import webbrowser
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import QDate, QLocale, Qt, QModelIndex
from PyQt5 import uic
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font
import subprocess
import configparser
import json
from subprocess import Popen, PIPE

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor


os.environ['QT_MAC_WANTS_LAYER'] = '1' #필수
myClasses =[]
myStudents = []
constDay = ['월', '화', '수', '목', '금', '토', '일']
with open('/Users/uhyeon/Downloads/아카이브/savedata.json', 'r',encoding='UTF-8') as f:
    json_data = json.load(f)
config = json_data
'''
for i in config.keys():
    for j in config[i]['students'].keys():
        config[i]['students'][j]['noKakao'] = False
'''
    
studentListModel = QStandardItemModel()
classListModel = QStandardItemModel()
kakaotalkScript = '''
on run {x,y}
tell application "KakaoTalk"
	reopen -- unminimizes the first minimized window or makes a new default window
	activate (first window whose name is "카카오톡") -- makes the app frontmost
end tell

tell application "System Events" to key code 18 using command down -- activate main window
delay 2
tell application "System Events" to key code 3 using command down
delay 2
set the clipboard to x
tell application "System Events" to key code 9 using command down
delay 2
tell application "System Events" to key code 125
tell application "System Events" to key code 36

set the clipboard to y
tell application "System Events" to key code 9 using command down
tell application "System Events" to key code 36
end run

on splitText(theText, theDelimiter)
        set AppleScript's text item delimiters to theDelimiter
        set theTextItems to every text item of theText
        set AppleScript's text item delimiters to ""
        return theTextItems
end splitText
'''
form_class = uic.loadUiType("/Users/uhyeon/Downloads/아카이브/PersonalProject.ui")[0]



class MyWindow(QMainWindow, form_class):

    def __init__(self):
        global myClasses
        global config
        super().__init__()
        
        self.setWindowTitle("메이킷코드랩")
        self.setWindowIcon(QIcon("/Users/uhyeon/Downloads/아카이브/PP_icon.jpg"))


        self.setupUi(self)

        myClasses = list(config.keys())
        for x in myClasses:
            classListModel.appendRow(QStandardItem(x))
        self.classList.setModel(classListModel)
        self.studentList.setModel(studentListModel)
        
        self.classList.clicked.connect(self.classSelectHandler)
        self.studentList.clicked.connect(self.studentSelectHandler)

        self.feedbackField.textChanged.connect(self.feedbackFieldHandler)

        self.classCommentField.textChanged.connect(self.classCommentFieldHandler)
        self.classSpecialField.textChanged.connect(self.classSpecialFieldHandler)

        self.isHome.stateChanged.connect(self.isHomeHandler)

        self.saveFeedback.clicked.connect(self.saveFeedbackHandler)
        self.sendKakaoButton.clicked.connect(self.allStudentSendKakao)
        self.saveExcel.clicked.connect(self.autoxl)
        self.savePPT.clicked.connect(self.autoppt)

        self.noKakaoCheck.clicked.connect(self.noKakaoCheckHandler)
        self.kakaoNameField.textChanged.connect(self.kakaoNameFieldHandler)
        self.removeStudentButton.clicked.connect(self.removeStudentHandler)
        self.addStudentButton.clicked.connect(self.addStudentHandler)

        self.addClassButton.clicked.connect(self.addClassHandler)
        self.removeClassButton.clicked.connect(self.removeClassHandler)


        self.inquiry() #statusBar에 시간 출력하기
    # def changeName(path, cName):
    #     i = 1
    #     for filename in os.listdir(path):
    #         os.rename(path+filename, path+str(cName))
    def addClassHandler(self):
        uclassName, ok = QInputDialog.getText(self, '반 추가', '반 이름 입력:')
        if ok:
            config[uclassName] = {'students':{}, \
                                    'classSpecial':'', \
                                    'classComment':'', \
                                    "folderName": "___",
                                    "classComment": "",
                                    "classSpecial": "",
                                    "excelCol": 4
            }
            myClasses.append(uclassName)
            classListModel.appendRow(QStandardItem(uclassName))
    
    def removeClassHandler(self):
        config.pop(self.getCurrentClass())
        myClasses.remove(self.getCurrentClass())
        classListModel.removeRow(self.classList.currentIndex().row()) 
            

    def removeStudentHandler(self):
        config[self.getCurrentClass()]['students'].pop(self.getCurrentStudent())
        myStudents.remove(self.getCurrentStudent())
        studentListModel.removeRow(self.studentList.currentIndex().row())
    
    def addStudentHandler(self):
        studentName, ok = QInputDialog.getText(self, '학생 추가', '학생 이름 입력:')
        if ok:
            config[self.getCurrentClass()]['students'][studentName] = {
				"isHome": False,
				"feedback": "",
				"kakaoName": "박우현",
				"noKakao": False
                }
            studentListModel.appendRow(QStandardItem(studentName))
            myStudents.append(studentName)

    def noKakaoCheckHandler(self):
        global config
        if self.noKakaoCheck.isChecked():
            config[self.getCurrentClass()]['students'][self.getCurrentStudent()]['noKakao'] = True
        else:
            config[self.getCurrentClass()]['students'][self.getCurrentStudent()]['noKakao'] = False

    def kakaoNameFieldHandler(self):
        global config
        config[self.getCurrentClass()]['students'][self.getCurrentStudent()]['kakaoName'] = self.kakaoNameField.text()
    def isHomeHandler(self):
        if self.isHome.isChecked():
            config[self.getCurrentClass()]['students'][self.getCurrentStudent()]['isHome'] = True
        else:
            config[self.getCurrentClass()]['students'][self.getCurrentStudent()]['isHome'] = False

    def autoppt(self):
        QLocale.setDefault(QLocale(QLocale.Korean, QLocale.SouthKorea))

        
        cur_date = QDate.currentDate()
        str_date = cur_date.toString("MM/dd")
        str_date += '('+constDay[cur_date.dayOfWeek()-1]+')'
        #print(str_date)
        prs = Presentation('/Users/uhyeon/Desktop/보고파일/박우현 연구원_데일리보고.pptx')
        '''
        for i in range(0, 11):
            print("--------[%d] ------ "%(i))
            slide = prs.slides.add_slide(prs.slide_layouts[i])
            for shape in slide.placeholders:
                print('%d %s' % (shape.placeholder_format.idx, shape.name))
        '''

        #16 = Date, 15, 14, 13,..
        slide_layout = prs.slide_layouts[7]
        slide = prs.slides.add_slide(slide_layout)
        tf = slide.placeholders[16].text_frame
        tf.text =str_date

        #출결
        idx=-1
        for i in range(0, len(myClasses)):
            myClassName = myClasses[i]
            if not constDay[cur_date.dayOfWeek()-1] in myClassName:
                continue
            idx+=1
            homePeople = []
            for student in config[myClassName]['students']:
                if config[myClassName]['students'][student]['isHome']:
                    homePeople.append(student + " 결석")
            if len(homePeople) == 0:
                homePeople.append('전원출석')
            homePeopleStr = ','.join(homePeople)

            #comment, special
            commentStr = config[myClassName]['classComment']
            specialStr = config[myClassName]['classSpecial']
            if specialStr == '':
                specialStr = '없습니다.'

            tf = slide.placeholders[15-idx].text_frame
            tf.text = '반이름 : ' + \
                myClassName + '\n' + \
                    '출결 : ' + homePeopleStr + '\n' + \
                    '진도 : ' + commentStr + '\n' + \
                    '특이사항 : '+specialStr

        self.move_slide(prs, prs.slides.index(slide), 1)

        prs.save('/Users/uhyeon/Desktop/보고파일/박우현 연구원_데일리보고.pptx')

    def move_slide(self, presentation, old_index, new_index):
        xml_slides = presentation.slides._sldIdLst  # pylint: disable=W0212
        slides = list(xml_slides)
        xml_slides.remove(slides[old_index])
        xml_slides.insert(new_index, slides[old_index])

    def classSpecialFieldHandler(self):
        config[self.getCurrentClass()]['classSpecial'] = self.classSpecialField.toPlainText()

    def classCommentFieldHandler(self):
        config[self.getCurrentClass()]['classComment'] = self.classCommentField.toPlainText()

    def allStudentSendKakao(self):
        for x in myStudents:
            if config[self.getCurrentClass()]['students'][x]['isHome'] or config[self.getCurrentClass()]['students'][x]['noKakao']:
                continue
            self.sendKakaotalk(config[self.getCurrentClass()]['students'][x]['kakaoName'],\
            config[self.getCurrentClass()]['students'][x]['feedback'])
    
    def sendKakaotalk(self, x, y):
        p = Popen(['osascript', '-'] + [x,y], stdin=PIPE, stdout=PIPE, stderr=PIPE)
        stdout, stderr = p.communicate(kakaotalkScript.encode('utf-8'))

    def saveFeedbackHandler(self):
        with open('savedata.json', 'w',encoding='UTF-8') as f:
            json.dump(config, f, ensure_ascii=False)

    def feedbackFieldHandler(self):
        config[self.getCurrentClass()]['students'][self.getCurrentStudent()]['feedback'] = self.feedbackField.toPlainText()

    def OpenFolder(self, Path):
        file_to_show = Path
        subprocess.call(["open", "-R", file_to_show])

    def CopySth(self, Text): # 클립보드에 
        cb = QApplication.clipboard()
        cb.clear(mode=cb.Clipboard)
        cb.setText(Text, mode=cb.Clipboard)

    def getCurrentClass(self):
        return myClasses[self.classList.currentIndex().row()]
    
    def getCurrentStudent(self):
        return myStudents[self.studentList.currentIndex().row()]

    def studentSelectHandler(self):
        self.feedbackField.setPlainText(config[self.getCurrentClass()]['students'][self.getCurrentStudent()]['feedback'])

        #load kakao
        self.kakaoNameField.setText(config[self.getCurrentClass()]['students'][self.getCurrentStudent()]['kakaoName'])

        #load isHome
        if config[self.getCurrentClass()]['students'][self.getCurrentStudent()]['isHome'] == True:
            self.isHome.setChecked(True)
        else:
            self.isHome.setChecked(False)

        #load noKakao
        if config[self.getCurrentClass()]['students'][self.getCurrentStudent()]['noKakao'] == True:
            self.noKakaoCheck.setChecked(True)
        else:
            self.noKakaoCheck.setChecked(False)


    def classSelectHandler(self):
        global myStudents
        myStudents =list(config[myClasses[self.classList.currentIndex().row()]]['students'].keys())
        studentListModel.clear()
        for x in myStudents:
            studentListModel.appendRow(QStandardItem(x))

        #load comment, special
        self.classSpecialField.setPlainText(config[self.getCurrentClass()]['classSpecial'])
        self.classCommentField.setPlainText(config[self.getCurrentClass()]['classComment'])

    def Write_KakaoTalk(self):
        cur_date = QDate.currentDate()
        str_date = cur_date.toString("MM.dd(ddd) ")
        file_name = str_date+self.lineEdit.text()
        homework_text = "(숙제) "+self.plainTextEdit_4.toPlainText()+"\n\n"
        first_line = '“안녕하세요?\n메이킷코드랩 코딩학원입니다.\n\n'
        end_line = '\n\n메이킷코드랩 홈페이지 http://makitcodelab.com\n송도센터 032-833-0046\n대치센터 02-6243-5000"'
        class_text = self.plainTextEdit_2.toPlainText()
        student_text = self.plainTextEdit_3.toPlainText()

        path = '/Users/myeongho/MyeongHo_/Codes/메이킷코드/KAKAOTALK/'+str_date
        try:
            if not os.path.exists(path):
                os.makedirs(path)
        except OSError:
            print ('Error: Creating directory. ' +  path)


        if (self.lineEdit.text()== ""):
            file_name = str_date + '(임시저장)'
        
        f = open(path+'/'+file_name+'.txt', 'w', encoding= 'UTF8')
        if (self.checkBox.isChecked()==False):
            msg = first_line+class_text+"\n\n"+student_text+end_line
            f.write(msg)
            f.close()
        else:
            msg = first_line + homework_text+class_text+'\n\n'+student_text+end_line
            f.write(msg)
            f.close()
        
        self.lineEdit.setText('')

    def open_kakaoTalk(self):
        self.OpenFolder("/Users/myeongho/MyeongHo_/Codes/메이킷코드/KAKAOTALK")

    def autoxl(self):
        global myClasses, myStudents
        wb = load_workbook(filename='/Users/uhyeon/Desktop/보고파일/박우현연구원.xlsx')
        tmpClassName = self.getCurrentClass().replace('/',',')
        tmpClassName = tmpClassName.replace(':','시')
        #print(tmpClassName)
        ws = wb[tmpClassName]
        row = 2
        col = 2

        while (ws.cell(row=row, column=col).value != None):
            row += 2
        
        #print(row)
        QLocale.setDefault(QLocale(QLocale.Korean, QLocale.SouthKorea))
        cur_date = QDate.currentDate()
        str_date = cur_date.toString(Qt.DefaultLocaleLongDate)

        titleColor = PatternFill(start_color='EDEDED', end_color='EDEDED', fill_type='solid')
        valueColor = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        defaultFont = Font(name='맑은 고딕', size=9, bold=False, italic=False, color='000000')


        ws.cell(row=row, column=col).value = str_date

        #fill
        ws.cell(row=row,column=col).fill = titleColor
        ws.cell(row=row+1,column=col).fill = valueColor

        #font
        ws.cell(row=row,column=col).font = defaultFont
        ws.cell(row=row+1,column=col).font = defaultFont

        #border
        ws.cell(row=row, column=col).border = Border(left=Side(border_style='thin', color='000000'),\
                                                            right=Side(border_style='thin', color='000000'),\
                                                            top=Side(border_style='thin', color='000000'),\
                                                            bottom=Side(border_style='thin', color='000000'))
        ws.cell(row=row+1, column=col).border = Border(left=Side(border_style='thin', color='000000'),\
                                                            right=Side(border_style='thin', color='000000'),\
                                                            top=Side(border_style='thin', color='000000'),\
                                                            bottom=Side(border_style='thin', color='000000'))
        for i in range(len(config[self.getCurrentClass()]['students'].keys())):
            #border
            ws.cell(row=row, column=col+i+1).border = Border(left=Side(border_style='thin', color='000000'),\
                                                            right=Side(border_style='thin', color='000000'),\
                                                            top=Side(border_style='thin', color='000000'),\
                                                            bottom=Side(border_style='thin', color='000000'))
            ws.cell(row=row+1, column=col+i+1).border = Border(left=Side(border_style='thin', color='000000'),\
                                                            right=Side(border_style='thin', color='000000'),\
                                                            top=Side(border_style='thin', color='000000'),\
                                                            bottom=Side(border_style='thin', color='000000'))
           
            #font
            ws.cell(row=row,column=col+i+1).font = defaultFont
            ws.cell(row=row+1,column=col+i+1).font = defaultFont 

            #fill
            ws.cell(row=row,column=col+i+1).fill = titleColor
            ws.cell(row=row+1,column=col+i+1).fill = valueColor

            #value
            ws.cell(row=row,column=col+i+1).value = myStudents[i]
            if(config[self.getCurrentClass()]['students'][myStudents[i]]['isHome'] == True):
                ws.cell(row=row+1,column=col+i+1).value = '결석'
            else:
                ws.cell(row=row+1,column=col+i+1).value = config[self.getCurrentClass()]['students'][myStudents[i]]['feedback']

        wb.save('/Users/uhyeon/Desktop/보고파일/박우현연구원.xlsx')

# jindo Excel
        wb = load_workbook(filename='/Users/uhyeon/Desktop/보고파일/(박우현 연구원)진도표.xlsx')
        defaultFont = Font(name='맑은 고딕', size=11, bold=False, italic=False, color='000000')
        ws = wb.active
        cName = self.getCurrentClass()
        row = 4
        col = config[cName]['excelCol']
        while (ws.cell(row=row, column=col).value != None):
            row += 1
        #print(row)

        ws.cell(row=row, column=col).value = config[cName]['classComment']
        ws.cell(row=row, column=col).font = defaultFont


        wb.save('/Users/uhyeon/Desktop/보고파일/(박우현 연구원)진도표.xlsx')

        
    def inquiry(self):
        cur_date = QDate.currentDate()
        str_date = cur_date.toString("yyyy년 MM월 dd일 dddd")
        self.statusBar().showMessage(str_date)



app = QApplication(sys.argv)
window = MyWindow()
window.show()
app.exec_()