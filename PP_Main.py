import os
import sys
import webbrowser
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import QDate
from PyQt5 import uic
import subprocess
os.environ['QT_MAC_WANTS_LAYER'] = '1' #필수

form_class = uic.loadUiType("/Users/myeongho/MyeongHo_/Codes/메이킷코드/Personal_Project/PersonalProject.ui")[0]

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("메이킷코드랩")
        self.setWindowIcon(QIcon("/Users/myeongho/MyeongHo_/Codes/메이킷코드/Personal_Project/PP_icon.jpg"))

        self.pushButton.clicked.connect(self.open_webbrowser_calender)
        self.pushButton_2.clicked.connect(self.open_webbrowser_notes)
        self.pushButton_3.clicked.connect(self.open_webbrowser_correction)
        self.pushButton_4.clicked.connect(self.open_webbrowser_ACAweb)
        self.pushButton_5.clicked.connect(self.copydate)
        # self.pushButton_6.clicked.connect(self.Copy_KakaoTalk_Feedback)
        self.pushButton_6.clicked.connect(self.autoxl)
        self.pushButton_7.clicked.connect(self.Copy_form)
        self.pushButton_8.clicked.connect(self.Copy_Sat10_form)
        self.pushButton_9.clicked.connect(self.Copy_Sat13_form)
        self.pushButton_10.clicked.connect(self.Copy_Sun13_form)
        self.pushButton_11.clicked.connect(self.Copy_Sun1510_form)
        self.pushButton_12.clicked.connect(self.OpenMakIT)
        self.pushButton_13.clicked.connect(self.Class)
        self.pushButton_14.clicked.connect(self.Daily_Report_KakaoTalk)
        self.pushButton_15.clicked.connect(self.Change_File_Name)
        self.pushButton_16.clicked.connect(self.Create_folder_sat)
        self.pushButton_17.clicked.connect(self.Create_folder_sun)
        self.pushButton_18.clicked.connect(self.Create_readme)
        self.pushButton_19.clicked.connect(self.Open_Readme)
        self.pushButton_20.clicked.connect(self.Write_KakaoTalk)
        self.pushButton_21.clicked.connect(self.open_kakaoTalk)
        self.inquiry() #statusBar에 시간 출력하기
    # def changeName(path, cName):
    #     i = 1
    #     for filename in os.listdir(path):
    #         os.rename(path+filename, path+str(cName))
    def OpenFolder(self, Path):
        file_to_show = Path
        subprocess.call(["open", "-R", file_to_show])

    def CopySth(self, Text): # 클립보드에 복사하기
        cb = QApplication.clipboard()
        cb.clear(mode=cb.Clipboard)
        cb.setText(Text, mode=cb.Clipboard)
    

    def open_webbrowser_calender(self):
        webbrowser.open('https://calendar.google.com/calendar/u/2/r?tab=rc')
    
    def open_webbrowser_notes(self):
        webbrowser.open('https://docs.google.com/spreadsheets/d/1nJQXdCRSKQGuwczvrQ8_aScwIqU9SfYiHfhy0Lk2R88/edit#gid=1648232156')
        
    def open_webbrowser_correction(self):
        webbrowser.open('https://speller.cs.pusan.ac.kr')

    def open_webbrowser_ACAweb(self):
        webbrowser.open('https://t.aca2000.co.kr/Account/Login?ReturnUrl=%2F')

    def copydate(self):
        cur_date = QDate.currentDate()
        # str_date = cur_date.toString(Qt.DefaultLocaleLongDate)
        str_date = cur_date.toString("yyyy년 MM월 dd일 dddd")
        self.CopySth(str_date)
    
    def Copy_KakaoTalk_Feedback(self):
        self.CopySth('“안녕하세요?\n메이킷코드랩 코딩학원입니다.\n\n메이킷코드랩 홈페이지 http://makitcodelab.com\n송도센터 032-833-0046\n대치센터 02-6243-5000"')
        os.system("open /System/Applications/Notes.app")
    
    def Daily_Report_KakaoTalk(self):
        cur_date = QDate.currentDate()
        str_date = cur_date.toString("MM.dd(ddd)")
        self.CopySth('[서명호 연구원 데일리 보고]\n\n-금일 '+str_date+'데일리 보고 및 수업 진도 관리 파일 첨부합니다.')

    def Copy_form(self):
        self.CopySth("학생(수업/시간/연구원)")
    def Copy_Sat10_form(self):
        self.CopySth("김도율, 김선균 학생 (씽커1/토 10:00-11:30/ 서명호 연구원)")
    def Copy_Sat13_form(self):
        self.CopySth("정지후, 권호윤 이인성 학생 (씽커3 /토 13:00-14:30/ 서명호 연구원)")
    def Copy_Sun13_form(self):
        self.CopySth("정세호, 황혜지, 황민지 학생(마이크로비트 메이킹/일 -1:00/ 서명호 연구원)")
    def Copy_Sun1510_form(self):
        self.CopySth("주현우, 박성현 학생(기초 파이썬/일 -3:10/ 서명호 연구원)")

    def OpenMakIT(self):
        self.OpenFolder("/Users/myeongho/MyeongHo_/Codes/메이킷코드/보고파일 3개")
        
    def Class(self):
        self.OpenFolder("/Users/myeongho/Desktop/5.서명호연구원 (1)")

    def Change_File_Name(self):
        cur_date = QDate.currentDate()
        str_date = cur_date.toString("yyyy.MM.dd")
        path = '/Users/myeongho/MyeongHo_/Codes/메이킷코드/보고파일 3개'

        file_path = [path+'/진도표', path+'/데일리보고', path+'/카카오톡피드백']

        for i in range(0, len(file_path)):
            file_name = file_path[i]
            file_list = os.listdir(file_name)
            src = os.path.join(file_name, file_list[0])
            if(i == 0):
                dst = '[서명호 연구원]진도표'+str_date+'.xlsx'
            elif(i == 1):
                dst = '서명호 연구원_데일리보고'+str_date+'.pptx'
            elif(i==2):
                dst = '서명호연구원'+str_date+'.xlsx'

            dst = os.path.join(file_name, dst)
            os.rename(src,dst)

    def Create_folder_sat(self):
        self.OpenFolder('/Users/myeongho/MyeongHo_/Codes/메이킷코드/토요일')

        folder_name = ['/Users/myeongho/MyeongHo_/Codes/메이킷코드/토요일/10시_씽커2_서명호연구원', '/Users/myeongho/MyeongHo_/Codes/메이킷코드/토요일/13시_마이크로비트메이킹_서명호연구원', '/Users/myeongho/MyeongHo_/Codes/메이킷코드/토요일/15시_파이썬4_서명호연구원']
        for path in folder_name:
            try:
                if not os.path.exists(path):
                    os.makedirs(path)
            except OSError:
                print ('Error: Creating directory. ' +  path)
        
    def Create_folder_sun(self):
        self.OpenFolder('/Users/myeongho/MyeongHo_/Codes/메이킷코드/일요일')
        folder_name = ['/Users/myeongho/MyeongHo_/Codes/메이킷코드/일요일/10시_기초파이썬3_서명호연구원', '/Users/myeongho/MyeongHo_/Codes/메이킷코드/일요일/13시_마이크로비트메이킹_서명호연구원', '/Users/myeongho/MyeongHo_/Codes/메이킷코드/일요일/15시_기초파이썬1_서명호연구원']
        
        for path in folder_name:
            try:
                if not os.path.exists(path):
                    os.makedirs(path)
            except OSError:
                print ('Error: Creating directory. ' +  path)

    def Create_readme(self, f):

        f = open ("/Users/myeongho/MyeongHo_/Codes/메이킷코드/README/readme.txt", 'w', encoding = "UTF8")
        
        text = self.plainTextEdit.toPlainText()
        try:
            f.write(text)
            f.close()
        except:
            f.write("ERROR")
            f.close()

    def Open_Readme(self):
        self.OpenFolder("/Users/myeongho/MyeongHo_/Codes/메이킷코드/README")

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
            f.write(first_line+class_text+"\n\n"+student_text+end_line)
            f.close()
        else:
            f.write(first_line + homework_text+class_text+'\n\n'+student_text+end_line)
            f.close()
        
        self.lineEdit.setText('')

    def open_kakaoTalk(self):
        self.OpenFolder("/Users/myeongho/MyeongHo_/Codes/메이킷코드/KAKAOTALK")

    def autoxl(self):
        from openpyxl import load_workbook

        wb = load_workbook("/Users/myeongho/MyeongHo_/Codes/메이킷코드/보고파일 3개/카카오톡피드백/서명호연구원2021.09.21.xlsx")
        path = '/Users/myeongho/MyeongHo_/Codes/메이킷코드/KAKAOTALK'


        cur_date = QDate.currentDate()
        str_date = cur_date.toString("MM.dd(ddd)")


        ws = wb.active

        for y in range(2, 30):
            for x in range(2, 12):
                print(ws.cell(row = x, column = y).value)

    def inquiry(self):
        cur_date = QDate.currentDate()
        str_date = cur_date.toString("yyyy년 MM월 dd일 dddd")
        self.statusBar().showMessage(str_date)

app = QApplication(sys.argv)
window = MyWindow()
window.show()
app.exec_()