# -*- coding: utf-8 -*-
"""
Created on Mon May 11 16:06:41 2020
batuhan ökmen
@author: okmen
"""

def anaMenu():
    
    global menusecim
    
    print("""
  ═════════════════════════════════════════════════════

             https://batuhanokmen.com/
      

  ═════════════════════════════════════════════════════ """)  





import sys


from PyQt5 import QtWidgets,uic,QtGui
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLineEdit, QTextBrowser, QPushButton, QVBoxLayout, QProgressBar,QComboBox
from excel_ceviri_python import Ui_MainWindow
from PyQt5.QtCore import QBasicTimer
from PyQt5.QtGui import *
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import *
from PyQt5.uic import loadUi
import os
import speech_recognition as sr
from gtts import gTTS
import socket
import getpass
import googletrans
from googletrans import Translator
import webbrowser
import xlsxwriter
import xlrd
from gtts import gTTS
import smtplib


        
def ceviriProgramim():
    
    class MainWindow(QMainWindow):
        def __init__(self):
            super(MainWindow, self).__init__()
            self.ui = Ui_MainWindow()
            self.ui.setupUi(self)
            #Butonları tanımlama
            self.ui.pushButton_excell.clicked.connect(self.excell_clicked_slot)
            self.ui.pushButton_ceviri.clicked.connect(self.ceviri_kaydet_button_clicked_slot)
            self.ui.pushButton_word.clicked.connect(self.word_button_clicked_slot)
            self.ui.pushButton_excel_kaydet.clicked.connect(self.excel_kaydet_clicked_slot)
            self.ui.pushButton_notepad.clicked.connect(self.notepad_clicked_slot)
            self.ui.pushButton_mailSend.clicked.connect(self.mailSend_clicked_slot)
            self.ui.pushButton_ceviriKaydet_orijinal.clicked.connect(self.ceviriKaydet_orijinal_button_clicked_slot)
            self.ui.pushButton_ceviriKaydet.clicked.connect(self.cevir_ses_button_clicked_slot)
            self.ui.commandLinkButton.clicked.connect(self.command_link_button_clicked_slot)
            self.ui.comboBox_s.currentTextChanged.connect(self.selectionchange)
            self.ui.comboBox_t.currentTextChanged.connect(self.selectionchangeCeviri)
            self.progressBar_vtt =QProgressBar(self)
            self.progressBar_vtt.setGeometry(10, 650, 1781, 23)


        

            
        #kaydetme kodu burada
        def ceviri_kaydet_button_clicked_slot(self):
            self.ui.textEdit_target.clear()
            print("başlıyom")
            okudum = self.ui.textEdit_source.toPlainText()
            print(okudum)
            
            global SecilenLangK
            global SecilenLangceviriK
            global ceviriSonucu
            translator = Translator()
            print(SecilenLangceviriK)
            sonuctrans = translator.translate(okudum, src=SecilenLangK, dest=SecilenLangceviriK)
            print("çeviri yapıldı ")
            ceviriSonucu = sonuctrans.text
            print(ceviriSonucu)
            self.ui.textEdit_target.setText(sonuctrans.text)
            ceviriSonucu=self.ui.textEdit_target.toPlainText()
            self.completed = 0
            while self.completed < 100:
                self.completed += 0.001
                self.progressBar_vtt.setValue(self.completed)
            print("Kayıt tamamlandı")

           
    
        #word çalışma kodu burada
        def word_button_clicked_slot(self):
            yazilacak_dosya=self.ui.textEdit_target.toPlainText()
            global ceviriSonucu
            with open('excel_translate_program.doc', "a") as dosya:
                dosya.write(yazilacak_dosya)
                dosya.write("\n")
            dosya.close()
            while self.completed < 100:
                self.completed += 0.001
                self.progressBar_vtt.setValue(self.completed)
            print("Kayıt tamamlandı")
        #pdf
        def excel_kaydet_clicked_slot(self):

            global SecilenLangK
            global SecilenLangceviriK
            global ceviriSonucu
            KelimeSayisi=self.ui.textEdit_wordCaunt.toPlainText()
            dosyaAdi=self.ui.textEdit_excel_isim.toPlainText()

            #okumalar
            dosyaAdi2=(dosyaAdi+(".xlsx"))
            book = xlrd.open_workbook(dosyaAdi2)
            first_sheet = book.sheet_by_index(0)

            #yazmalar
            workbook = xlsxwriter.Workbook('excel_translate_program.xlsx')
            print("Oluşturdum")
            worksheet = workbook.add_worksheet()
            sayac= 1
            satır = 0
            sutun = 0
            progres_artis=(100/int(KelimeSayisi))
            progres= 0
            while 1:
                if (sayac <= int(KelimeSayisi)):    
                    #okuma kodu tamam
                    cumle = first_sheet.cell(satır,sutun)
                    print(cumle.value)
                    satır=satır+1
                    #çevirme
                    
                    translator = Translator()
                    print(SecilenLangceviriK)
                    sonuctrans = translator.translate(cumle.value, src=SecilenLangK, dest=SecilenLangceviriK)
                    print("çeviri yapıldı ")
                    ceviriSonucu = sonuctrans.text
                    print(ceviriSonucu)

                    #yazma   
                    konum=("A%s" % sayac)
                    worksheet.write(konum, ceviriSonucu)
                    sayac=sayac+1
                    #progres
                    progres=progres+progres_artis
                    self.progressBar_vtt.setValue(progres)
                      
                else:
                    print("bitti")
                    workbook.close()
                    self.progressBar_vtt.setValue(100)
                    break   

        #notepad
        def notepad_clicked_slot(self):
            yazilacak_dosya=self.ui.textEdit_target.toPlainText()
            with open('excel_translate_program.txt', "a") as dosya:
                dosya.write(yazilacak_dosya)
                dosya.write("\n")
            dosya.close()
            while self.completed < 100:
                self.completed += 0.001
                self.progressBar_vtt.setValue(self.completed)
            print("Kayıt tamamlandı")
        def mailSend_clicked_slot(self):
            mailicin=self.ui.textEdit_target.toPlainText()
            aliciMail = self.ui.textEdit_mail.toPlainText()
            smtpserver = smtplib.SMTP("smtp.gmail.com", 587)
            smtpserver.ehlo()
            smtpserver.starttls()
            smtpserver.ehlo()
            smtpserver.login("batuhanokmensite@gmail.com","Batuhanmac123")
            message = (mailicin).encode('utf-8').strip()
            smtpserver.sendmail("batuhanokmensite@gmail.com",aliciMail, message)
            smtpserver.quit
            print("mailiniz yollandı")
        


        
        


        

        def selectionchangeCeviri(self):
            global SecilenLangceviri
            global SecilenLangceviriK
            secilenIndexceviri = self.ui.comboBox_t.currentIndex()
            secilenIndexceviriisim = self.ui.comboBox_t.currentText()
            print(secilenIndexceviriisim)
            if (secilenIndexceviri== 1 ):
                SecilenLangceviri=('tr-TR')
                SecilenLangceviriK=('tr')
            elif (secilenIndexceviri== 2 ):
                SecilenLangceviri=('en-US')
                SecilenLangceviriK=('en')
            elif (secilenIndexceviri== 3 ):
                SecilenLangceviri=('es-ES')
                SecilenLangceviriK=('es')
            elif (secilenIndexceviri== 4 ):
                SecilenLangceviri=('de-DE')
                SecilenLangceviriK=('de')
            elif (secilenIndexceviri== 5 ):
                SecilenLangceviri=('pl-PL')
                SecilenLangceviriK=('pl')
            elif (secilenIndexceviri== 6):
                SecilenLangceviri=('pt-BR')
                SecilenLangceviriK=('pt')
            elif (secilenIndexceviri== 7):
                SecilenLangceviri=('it-IT')
                SecilenLangceviriK=('it')
            elif (secilenIndexceviri== 8):
                SecilenLangceviri=('hr-HR')
                SecilenLangceviriK=('hr')
            elif (secilenIndexceviri== 9):
                SecilenLangceviri=('nl-NL')
                SecilenLangceviriK=('nl')
            elif (secilenIndexceviri== 10):
                SecilenLangceviri=('da-DK')
                SecilenLangceviriK=('da')
            elif (secilenIndexceviri== 11):
                SecilenLangceviri=('lt-LT')
                SecilenLangceviriK=('lt')
            elif (secilenIndexceviri== 12):
                SecilenLangceviri=('et-EE')
                SecilenLangceviriK=('et')
            elif (secilenIndexceviri==13):
                SecilenLangceviri=('fr-FR')
                SecilenLangceviriK=('fr')
            elif (secilenIndexceviri==14):
                SecilenLangceviri=('ar-AE')
                SecilenLangceviriK=('ar')
            elif (secilenIndexceviri==15):
                SecilenLangceviri=('ja-JP')
                SecilenLangceviriK=('ja')
            elif (secilenIndexceviri==16):
                SecilenLangceviri=('ko-KR')
                SecilenLangceviriK=('ko')
            elif (secilenIndexceviri==17):
                SecilenLangceviri=('ru-RU')
                SecilenLangceviriK=('ru')
            elif (secilenIndexceviri==18):
                SecilenLangceviri=('zh-CN')
                SecilenLangceviriK=('zh')
            elif (secilenIndexceviri==19):
                SecilenLangceviri=('bg-BG')
                SecilenLangceviriK=('bg')
            elif (secilenIndexceviri==20):
                SecilenLangceviri=('cs-CZ')
                SecilenLangceviriK=('cs')
            elif (secilenIndexceviri==21):
                SecilenLangceviri=('da-DK')
                SecilenLangceviriK=('da')
            elif (secilenIndexceviri==22):
                SecilenLangceviri=('el-GR')
                SecilenLangceviriK=('el')
            elif (secilenIndexceviri==23):
                SecilenLangceviri=('fi-FI')
                SecilenLangceviriK=('fi')
            elif (secilenIndexceviri==24):
                SecilenLangceviri=('he-IL')
                SecilenLangceviriK=('he')
            elif (secilenIndexceviri==25):
                SecilenLangceviri=('hi-IN')
                SecilenLangceviriK=('hi')
            elif (secilenIndexceviri==26):
                SecilenLangceviri=('hr-HR')
                SecilenLangceviriK=('hr')
            elif (secilenIndexceviri==27):
                SecilenLangceviri=('id-ID')
                SecilenLangceviriK=('id')
            elif (secilenIndexceviri==28):
                SecilenLangceviri=('ro-RO')
                SecilenLangceviriK=('ro')
            elif (secilenIndexceviri==29):
                SecilenLangceviri=('sr-RS')
                SecilenLangceviriK=('sr')
            elif (secilenIndexceviri==30):
                SecilenLangceviri=('sv-SE')
                SecilenLangceviriK=('sv')
            elif (secilenIndexceviri==31):
                SecilenLangceviri=('uk-UA')
                SecilenLangceviriK=('uk')
            elif (secilenIndexceviri==32):
                SecilenLangceviri=('vi-VN')
                SecilenLangceviriK=('vi')
            elif (secilenIndexceviri==33):
                SecilenLangceviri=('af-ZA')
                SecilenLangceviriK=('af')
            elif (secilenIndexceviri==34):
                SecilenLangceviri=('az-AZ')
                SecilenLangceviriK=('az')
            elif (secilenIndexceviri==35):
                SecilenLangceviri=('fa-IR')
                SecilenLangceviriK=('fa')
            elif (secilenIndexceviri==36):
                SecilenLangceviri=('fil-PH')
                SecilenLangceviriK=('fil')
            elif (secilenIndexceviri==37):
                SecilenLangceviri=('sw-KE')
                SecilenLangceviriK=('sw')
            elif (secilenIndexceviri==38):
                SecilenLangceviri=('ar-SA')
                SecilenLangceviriK=('ar')
            elif (secilenIndexceviri==39):
                SecilenLangceviri=('ar-QA')
                SecilenLangceviriK=('ar')
            elif (secilenIndexceviri==40):
                SecilenLangceviri=('ar-KW')
                SecilenLangceviriK=('ar')

            else:
                print("yanlış oldu")
        


        def selectionchange(self):
            global SecilenLang
            global SecilenLangK
            secilenIndex = self.ui.comboBox_s.currentIndex()
            secilenIndexisim = self.ui.comboBox_s.currentText()
            print(secilenIndexisim)
            if (secilenIndex== 1 ):
                SecilenLang=('tr-TR')
                SecilenLangK=('tr')
            elif (secilenIndex== 2 ):
                SecilenLang=('en-US')
                SecilenLangK=('en')
            elif (secilenIndex== 3 ):
                SecilenLang=('es-ES')
                SecilenLangK=('es')
            elif (secilenIndex== 4 ):
                SecilenLang=('de-DE')
                SecilenLangK=('de')
            elif (secilenIndex== 5 ):
                SecilenLang=('pl-PL')
                SecilenLangK=('pl')
            elif (secilenIndex== 6):
                SecilenLang=('pt-BR')
                SecilenLangK=('pt')
            elif (secilenIndex== 7):
                SecilenLang=('it-IT')
                SecilenLangK=('it')
            elif (secilenIndex== 8):
                SecilenLang=('hr-HR')
                SecilenLangK=('hr')
            elif (secilenIndex== 9):
                SecilenLang=('nl-NL')
                SecilenLangK=('nl')
            elif (secilenIndex== 10):
                SecilenLang=('da-DK')
                SecilenLangK=('da')
            elif (secilenIndex== 11):
                SecilenLang=('lt-LT')
                SecilenLangK=('lt')
            elif (secilenIndex== 12):
                SecilenLang=('et-EE')
                SecilenLangK=('et')
            elif (secilenIndex==13):
                SecilenLang=('fr-FR')
                SecilenLangK=('fr')
            elif (secilenIndex==14):
                SecilenLang=('ar-AE')
                SecilenLangK=('ar')
            elif (secilenIndex==15):
                SecilenLang=('ja-JP')
                SecilenLangK=('ja')
            elif (secilenIndex==16):
                SecilenLang=('ko-KR')
                SecilenLangK=('ko')
            elif (secilenIndex==17):
                SecilenLang=('ru-RU')
                SecilenLangK=('ru')
            elif (secilenIndex==18):
                SecilenLang=('zh-CN')
                SecilenLangK=('zh')
            elif (secilenIndex==19):
                SecilenLang=('bg-BG')
                SecilenLangK=('bg')
            elif (secilenIndex==20):
                SecilenLang=('cs-CZ')
                SecilenLangK=('cs')
            elif (secilenIndex==21):
                SecilenLang=('da-DK')
                SecilenLangK=('da')
            elif (secilenIndex==22):
                SecilenLang=('el-GR')
                SecilenLangK=('el')
            elif (secilenIndex==23):
                SecilenLang=('fi-FI')
                SecilenLangK=('fi')
            elif (secilenIndex==24):
                SecilenLang=('he-IL')
                SecilenLangK=('he')
            elif (secilenIndex==25):
                SecilenLang=('hi-IN')
                SecilenLangK=('hi')
            elif (secilenIndex==26):
                SecilenLang=('hr-HR')
                SecilenLangK=('hr')
            elif (secilenIndex==27):
                SecilenLang=('id-ID')
                SecilenLangK=('id')
            elif (secilenIndex==28):
                SecilenLang=('ro-RO')
                SecilenLangK=('ro')
            elif (secilenIndex==29):
                SecilenLang=('sr-RS')
                SecilenLangK=('sr')
            elif (secilenIndex==30):
                SecilenLang=('sv-SE')
                SecilenLangK=('sv')
            elif (secilenIndex==31):
                SecilenLang=('uk-UA')
                SecilenLangK=('uk')
            elif (secilenIndex==32):
                SecilenLang=('vi-VN')
                SecilenLangK=('vi')
            elif (secilenIndex==33):
                SecilenLang=('af-ZA')
                SecilenLangK=('af')
            elif (secilenIndex==34):
                SecilenLang=('az-AZ')
                SecilenLangK=('az')
            elif (secilenIndex==35):
                SecilenLang=('fa-IR')
                SecilenLangK=('fa')
            elif (secilenIndex==36):
                SecilenLang=('fil-PH')
                SecilenLangK=('fil')
            elif (secilenIndex==37):
                SecilenLang=('sw-KE')
                SecilenLangK=('sw')
            elif (secilenIndex==38):
                SecilenLang=('ar-SA')
                SecilenLangK=('ar')
            elif (secilenIndex==39):
                SecilenLang=('ar-QA')
                SecilenLangK=('ar')
            elif (secilenIndex==40):
                SecilenLang=('ar-KW')
                SecilenLangK=('ar')
            else:
                print("yanlış")
        """
        Buralar biraz karışacak yapılan şey saçma olsa da gerekli durumundadır
        Burada veri kaydediliş işlemlerini yapacağım

        """
        def excell_clicked_slot(self):
            KelimeSayisi=self.ui.textEdit_wordCaunt.toPlainText()
            self.ui.textEdit_source.clear()
            dosyaAdi=self.ui.textEdit_excel_isim.toPlainText()


            dosyaAdi2=(dosyaAdi+(".xlsx"))

            book = xlrd.open_workbook(dosyaAdi2)
            first_sheet = book.sheet_by_index(0)
            sayac= 1
            satır = 0
            sutun = 0
            progres_artis=(100/int(KelimeSayisi))
            progres= 0
            while 1:
                if (sayac <= int(KelimeSayisi)):        
                    #okuma kodu tamam
                    cumle = first_sheet.cell(satır,sutun)
                    print(cumle.value)
                    sayac=sayac+1
                    satır=satır+1
                    self.ui.textEdit_source.append(cumle.value)
                    #progres
                    progres=progres+progres_artis
                    self.progressBar_vtt.setValue(progres)
        
                else:
                    print("bitti")
                    self.progressBar_vtt.setValue(100)
                    break
        


        def cevir_ses_button_clicked_slot(self):
            global SecilenLangK
            global SecilenLangceviriK

            KelimeSayisi=self.ui.textEdit_wordCaunt.toPlainText()
            book = xlrd.open_workbook('excel_translate_program.xlsx')
            first_sheet = book.sheet_by_index(0)

            
            sayac= 1
            satır = 0
            sutun = 0
            progres_artis=(100/int(KelimeSayisi))
            progres= 0

            while 1:
                if (sayac <= int(KelimeSayisi)):        
                    #okuma kodu tamam
                    cumle = first_sheet.cell(satır,sutun)
                    print(cumle.value)
                    myobj = gTTS(text=cumle.value, lang=SecilenLangceviriK, slow=False)
                    myobj.save("%s.mp3" % cumle.value)
                    sayac=sayac+1
                    satır=satır+1
                    #progres
                    progres=progres+progres_artis
                    self.progressBar_vtt.setValue(progres)
        
                else:
                    print("bitti")
                    self.progressBar_vtt.setValue(100)
                    break

            
                
            
                
            
            




        def ceviriKaydet_orijinal_button_clicked_slot(self):
            global SecilenLang
            global SecilenLangK
            KelimeSayisi=self.ui.textEdit_wordCaunt.toPlainText()
            dosyaAdi=self.ui.textEdit_excel_isim.toPlainText()


            dosyaAdi2=(dosyaAdi+(".xlsx"))
            book = xlrd.open_workbook(dosyaAdi2)
            first_sheet = book.sheet_by_index(0)

            sayac= 1
            satır = 0
            sutun = 0
            progres_artis=(100/int(KelimeSayisi))
            progres= 0
            while 1:
                if (sayac <= int(KelimeSayisi)):        
                    #okuma kodu tamam
                    cumle = first_sheet.cell(satır,sutun)
                    print(cumle.value)
                    myobj = gTTS(text=cumle.value, lang=SecilenLangK, slow=False)
                    myobj.save("%s.mp3" % cumle.value)
                    sayac=sayac+1
                    satır=satır+1
                    #progres
                    progres=progres+progres_artis
                    self.progressBar_vtt.setValue(progres)
        
                else:
                    print("bitti")
                    self.progressBar_vtt.setValue(100)
                    break
        
        









        
        
        def command_link_button_clicked_slot(self):
            url = 'https://batuhanokmen.com/2020/05/25/voice-to-text-indir/'
            webbrowser.open_new_tab(url)
        
    
    
         
    
    
    
    
    if __name__ == "__main__":
        app = QApplication(sys.argv)
    
        window = MainWindow()
        window.setWindowTitle("Translate program")
        window.show()
    
        sys.exit(app.exec_())



"""
başlayalım

"""


while True:
    anaMenu()
    secim= 1
    if secim == 1:
        ceviriProgramim()
    else:
        print("Bir hata oluştu")











