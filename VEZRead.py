# pylint: disable=E0611
# pylint: disable=E1101
#pyi-makespec --onefile --icon=icon.ico --noconsole VEZRead.py
from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QLabel,\
    QScrollArea,QSizePolicy, QTableWidgetItem,QSplitter, QFrame, QSizePolicy, QListView, QTableWidget, qApp, QAction,\
     QMessageBox,QFileDialog, QErrorMessage, QDoubleSpinBox, QSpacerItem, QLineEdit, QItemDelegate 
from PyQt5.QtGui import QPixmap, QPalette, QBrush, QImage, QIcon, QTransform, QStandardItemModel,QStandardItem,\
     QDoubleValidator, QValidator, QCloseEvent
from PyQt5.QtCore import QPersistentModelIndex, Qt,  QSize, QModelIndex
from PyQt5 import uic
import sys
import os
import copy 
import win32api
from docx import Document
#from PIL import Image

import imagescan
import roo
import xlsx
import Word

class MyFrame(QWidget):
    def __init__(self,Trig,parent=None):
        super().__init__(parent)
        self.scaleFactor = 1.0
        self.imageLabel = QLabel()
        self.imageLabel.setBackgroundRole(QPalette.Base)
        self.imageLabel.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.imageLabel.setScaledContents(True)
        self.scrollArea = QScrollArea()
        self.scrollArea.setBackgroundRole(QPalette.Dark)
        self.scrollArea.setWidget(self.imageLabel)
        self.scrollArea.setVisible(False)
        self.layout = QVBoxLayout(self)
        self.layout.addWidget(self.scrollArea)
        self.Trig = Trig
        self.zoom = 1
        self.angle = 0
        #self.open()

    def zoomIn(self):
        try:
            if not self.Trig:
                self.zoom *=1.25
                self.scaleImage(self.zoom)
        except Exception:
            None
 
    def zoomOut(self):
        try:
            if not self.Trig:
                self.zoom *=0.8
                self.scaleImage(self.zoom)
        except Exception:
            None

    def Rotate(self, direction):
        try:
            self.angle += 90 * direction
            t = QTransform().rotate(self.angle)
            self.imageLabel.setPixmap(QPixmap.fromImage(self.image).transformed(t))
            self.scrollArea.setVisible(True)
            self.imageLabel.adjustSize()
            self.scaleImage(self.zoom)
            if self.Trig:
                self.scrollArea.setWidgetResizable(True)
        except Exception:
            None
 
    def normalSize(self):
        try:
            if not self.Trig:
                self.imageLabel.adjustSize()
                #self.scaleFactor = min(self.x/self.ps[0],self.y/self.ps[1])
                self.zoom =1
                self.scaleImage(self.zoom)
        except Exception:
            None
            
    def fitToWindow(self):
        try:
            self.Trig = not self.Trig
            self.scrollArea.setWidgetResizable(self.Trig)
            if not self.Trig:
                self.normalSize()
        except Exception:
            None

    def scaleImage(self, factor):
        try:
            self.imageLabel.resize(self.scaleFactor * factor * self.imageLabel.pixmap().size())
            self.adjustScrollBar(self.scrollArea.horizontalScrollBar(), factor)
            self.adjustScrollBar(self.scrollArea.verticalScrollBar(), factor)
        except Exception:
            None
 
    def adjustScrollBar(self, scrollBar, factor):
        try:
            scrollBar.setValue(int(factor * scrollBar.value()
                                + ((factor - 1) * scrollBar.pageStep() / 2)))
        except Exception:
            None
  
    def Update(self, size): 
        try:
            self.x = size.width()
            self.y = size.height()
            self.resize(self.x,self.y)
            #self.scaleFactor = min(self.x/self.ps[0],self.y/self.ps[1])
            #self.scaleImage(self.zoom)
        except Exception:
            None

    def open(self,file_path):
        try:
            self.image = QImage(file_path)
            self.ps = (self.image.size().width(),self.image.size().height())
            t = QTransform().rotate(self.angle)
            self.imageLabel.setPixmap(QPixmap.fromImage(self.image).transformed(t))
            self.scrollArea.setVisible(True)
            self.imageLabel.adjustSize()
            self.scaleImage(self.zoom)
            if self.Trig:
                self.scrollArea.setWidgetResizable(True)
        except Exception:
            None

class DownloadDelegate(QItemDelegate):
    """ Переопределение поведения ячейки таблицы """
    def __init__(self, parent=None):
        super(DownloadDelegate, self).__init__(parent)

    def createEditor(self, parent, option, index):
        lineedit=QLineEdit(parent)
        lineedit.setValidator(MyValidator(0,5,lineedit))
        return lineedit#QItemDelegate.createEditor(self, parent, option, index)

class MyValidator(QValidator):
    """ Позволяет вводить только числа """
    def __init__(self, min, max, parent):
        QValidator.__init__(self, parent)
        self.s = set(['0','1','2','3','4','5','6','7','8','9','.',',',''])

    def validate(self, s, pos): 
        """ проверяет привильная ли строка """       
        i=-1
        t1 = 0
        t2 = 0
        for i in range(len(s)):
            if s[i] == ".":
                t1 += 1
            if s[i] == ",":
                t2 += 1
            if t1+t2>1:
                i=-1
                break
            if s[i] not in self.s:
                i-=1
                break
        
        if i == len(s)-1:
            return (QValidator.Acceptable, s, pos)
        else:
            return (QValidator.Invalid, s, pos)

    def fixup(self, s):
        """ форматирует неправильную строку """
        s1=''
        t = False
        for i in s:
            if i in self.s:
                if  (i=="." or i==","):
                    if not t:
                        s1+=i
                        t = True
                else:
                    s1+=i
        s=s1

class CustomList(QListView):
    """ Добавляем сигнал завершения ввода  """
    def __init__(self,parent=None):
        super().__init__(parent)
        self.Trig = False

    def setcloseEditorSignal(self,metod):
        self.metod = metod
        self.Trig = True

    def closeEditor(self, editor, hint):
        if self.Trig:
            self.metod()
        QListView.closeEditor(self, editor, hint)

class CustomTable(QTableWidget):
    """ Добавляем сигнал завершения ввода  """
    def __init__(self,parent=None):
        super().__init__(parent)
        self.Trig = False

    def setcloseEditorSignal(self,metod):
        self.metod = metod
        self.Trig = True

    def closeEditor(self, editor, hint):
        if self.Trig:
            self.metod()
        QTableWidget.closeEditor(self, editor, hint)
        

        
class MyWindow(QMainWindow):
    def __init__(self,parent=None):
        super(MyWindow,self).__init__(parent)
        #uic.loadUi('gui_template.ui',self)

        try:
            self.path_home = os.path.expanduser("~\\Desktop\\")
        except Exception:
            self.path_home = ""


        for curren_dir in ["interim_image"]:
            if os.path.exists(curren_dir):
                if os.path.isdir(curren_dir):
                    print(curren_dir+" is here")
                else:
                    try:
                        os.mkdir(curren_dir)
                    except OSError:
                        print ("Error generate dir "+curren_dir)
            else:
                try:
                    os.mkdir(curren_dir)
                except OSError:
                    print ("Error generate dir "+curren_dir)

        self.lv_c = 5
        self.t_c = 0.5
        self.imnam = "слои"

        desktop = QApplication.desktop()
        wd = desktop.width()
        hg = desktop.height()
        ww = 1000
        wh = 500
        if ww>wd: ww = int(0.7*wd)
        if wh>hg: wh = int(0.7*hg)
        x = (wd-ww)//2
        y = (hg-wh)//2
        self.setGeometry(x, y, ww, wh)
        
        
        topVBoxLayout = QVBoxLayout(self) 
        topVBoxLayout.setContentsMargins(0,0,0,0) 


        SPFrame = QFrame() 
        SPFrame.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed) 
        SPFrame.setMinimumSize(QSize(150, 100)) 

        self.ImFrame = QFrame() 
        self.ImFrame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding) 
        #self.ImFrame.setMinimumSize(QSize(300, 100))

        TbFrame = QFrame() 
        TbFrame.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed) 
        TbFrame.setMinimumSize(QSize(0, 100))

        RFrame = QFrame() 
        RFrame.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding) 
        RFrame.setMinimumSize(QSize(0, 100))
        

        self.listView = CustomList()#QListView()
        self.listView.clicked.connect(self.OpenPict)
        self.listView.setcloseEditorSignal(self.Rename)
        #self.listView.editingFinished.connect(self.Rename)
        #self.listView.edit.connect(self.Rename)
        

        BoxLayout1 = QHBoxLayout()
        BoxLayout1.addWidget(self.listView)
        SPFrame.setLayout(BoxLayout1) 

        self.Table = CustomTable()#QTableWidget()
        self.Table.setcloseEditorSignal(self.WriteTable)
        self.Table.setColumnCount(3)
        self.Table.setHorizontalHeaderLabels(["p", "h", "d"])
        self.Table.setRowCount(30)
        self.Table.setColumnWidth(0,50) 
        self.Table.setColumnWidth(1,50)
        self.Table.setColumnWidth(2,50)
        self.Table.setItemDelegate(DownloadDelegate(self))
        BoxLayout2 = QHBoxLayout()
        BoxLayout2.addWidget(self.Table)
        TbFrame.setLayout(BoxLayout2)

        self.pe1 = QLabel("Рэ1=")
        self.pe2 = QLabel("Рэ2=")
        self.pe = QLabel("Рэ=")
        lb1 = QLabel("Длинна вертикального")
        lb2 = QLabel("заземлителя lв, м")
        lb3 = QLabel("Глубина заложения")
        lb4 = QLabel("вертикального заземлителя t, м")
        lb5 = QLabel("Эквивалентное сопротивление")
        lb6 = QLabel("Двухслойная модель")
        lb7 = QLabel("Однослойная модель")
        lb8 = QLabel("Шаблон искомого изображения")
        lb9 = QLabel("Название проекта")
        lb10 = QLabel("Название ВЛ")
        self.vl = QDoubleSpinBox()
        self.vl.setValue(self.lv_c)
        self.vl.valueChanged.connect(self.WriteTable)
        self.t = QDoubleSpinBox()
        self.t.setValue(self.t_c)
        self.t.valueChanged.connect(self.WriteTable)
        self.pe1_le = QLineEdit()
        self.pe2_le = QLineEdit()
        self.pe_le = QLineEdit()
        self.ImageName = QLineEdit()
        self.ImageName.setText(self.imnam)
        self.NPrj = QLineEdit()
        self.NVL = QLineEdit()
        self.ImageName.editingFinished.connect(self.ReadImName)
        spacerItem = QSpacerItem(2, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        BoxLayout3 = QVBoxLayout()
        BoxLayout4 = QHBoxLayout()
        BoxLayout5 = QHBoxLayout()
        BoxLayout6 = QHBoxLayout()
        BoxLayout3.addWidget(lb8)
        BoxLayout3.addWidget(self.ImageName)
        BoxLayout3.addWidget(lb9)
        BoxLayout3.addWidget(self.NPrj)
        BoxLayout3.addWidget(lb10)
        BoxLayout3.addWidget(self.NVL)
        BoxLayout3.addWidget(lb1)
        BoxLayout3.addWidget(lb2)
        BoxLayout3.addWidget(self.vl)
        BoxLayout3.addWidget(lb3)
        BoxLayout3.addWidget(lb4)
        BoxLayout3.addWidget(self.t)
        BoxLayout3.addWidget(lb5)
        BoxLayout3.addWidget(lb6)
        BoxLayout4.addWidget(self.pe1)
        BoxLayout4.addWidget(self.pe1_le)
        BoxLayout3.addLayout(BoxLayout4)
        BoxLayout5.addWidget(self.pe2)
        BoxLayout5.addWidget(self.pe2_le)
        BoxLayout3.addLayout(BoxLayout5)
        BoxLayout3.addWidget(lb7)
        BoxLayout6.addWidget(self.pe)
        BoxLayout6.addWidget(self.pe_le)
        BoxLayout3.addLayout(BoxLayout6)
        BoxLayout3.addItem(spacerItem)
        RFrame.setLayout(BoxLayout3)

        
        Splitter1 = QSplitter(Qt.Horizontal) 
        Splitter1.addWidget(SPFrame)
        Splitter1.addWidget(self.ImFrame)
        Splitter1.addWidget(TbFrame)
        Splitter1.addWidget(RFrame)
        Splitter1.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        Splitter1.setStretchFactor(1, 4) 
        Splitter1.setStretchFactor(1, 1)


        topVBoxLayout.addWidget(Splitter1) 

        self.central_widget = QWidget()
        self.central_widget.setLayout(topVBoxLayout)
        self.setCentralWidget(self.central_widget) 
        
        #self.resize()
        #self.central_widget.show()
        self.PaintForm = MyFrame(False,parent=self.ImFrame)

        self.statusBar()
        menubar = self.menuBar()

        exitAction = QAction( '&Выход', self) #QIcon('exit.png'),
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('Выход и программы')
        exitAction.triggered.connect(lambda:self.closeEvent(QCloseEvent()))

        Open0Action = QAction( '&Открыть изображение', self) #QIcon('exit.png'),
        Open0Action.setShortcut('Ctrl+1')
        Open0Action.setStatusTip('Результаты ВЭЗ предоставлены одиночными изображениями')
        Open0Action.triggered.connect(self.Op0)

        Open1Action = QAction( '&Открыть папку изображений', self) #QIcon('exit.png'),
        Open1Action.setShortcut('Ctrl+2')
        Open1Action.setStatusTip('Результаты ВЭЗ предоставлены одиночными изображениями')
        Open1Action.triggered.connect(lambda:self.Op12(0))

        Open2Action = QAction( '&Открыть папку каталогов', self) #QIcon('exit.png'),
        Open2Action.setShortcut('Ctrl+3')
        Open2Action.setStatusTip('Результаты ВЭЗ предоставлены набором файлов')
        Open2Action.triggered.connect(lambda:self.Op12(1))

        Open3Action = QAction( '&Открыть файл Ecxel', self) #QIcon('exit.png'),
        Open3Action.setShortcut('Ctrl+4')
        Open3Action.setStatusTip('Результаты ВЭЗ файле Ecxel')
        Open3Action.triggered.connect(self.Op3)

        Open4Action = QAction( '&Открыть файлы Ecxel', self) #QIcon('exit.png'),
        Open4Action.setShortcut('Ctrl+5')
        Open4Action.setStatusTip('Результаты ВЭЗ в нескольких файлах Ecxel')
        Open4Action.triggered.connect(self.Op4)

        NewPick = QAction("Добавить точку", self)
        NewPick.setShortcut('Ctrl+A')
        NewPick.setStatusTip('Добавить новую точку')
        NewPick.triggered.connect(self.NewPick)

        DelPick = QAction("Удалить точку", self)
        DelPick.setShortcut('Ctrl+D')
        DelPick.setStatusTip('Удалить точку из списка')
        DelPick.triggered.connect(self.DelPick)

        ReRead = QAction("Прочитать изображение", self)
        ReRead.setShortcut('Ctrl+R')
        ReRead.setStatusTip('Принудительно считывает информацию с изображения')
        ReRead.triggered.connect(self.ReReadImage)

        Select = QAction("Выбрать все точки", self)
        Select.setShortcut('Ctrl+W')
        Select.setStatusTip('Выбираем все точки из списка')
        Select.triggered.connect(self.select)

        ClearTable = QAction("Очистить таблицу", self)
        ClearTable.setShortcut('Ctrl+Shift+D')
        ClearTable.setStatusTip('Очищает текущую таблицу')
        ClearTable.triggered.connect(self.ClearTable)

        SetLvCheck = QAction("Установить lв выбранным точкам", self)
        SetLvCheck.setShortcut('Ctrl+E')
        SetLvCheck.setStatusTip('Изменяет lв выбраных точек на текущее')
        SetLvCheck.triggered.connect(self.SetLvCheck)

        SetTCheck = QAction("Установить t выбранным точкам", self)
        SetTCheck.setShortcut('Ctrl+T')
        SetTCheck.setStatusTip('Изменяет t выбраных точек на текущее')
        SetTCheck.triggered.connect(self.SetTCheck)

        Save1VEZExcel = QAction("Сохранить ВЭЗ в Excel", self)
        Save1VEZExcel.setShortcut('Ctrl+6')
        Save1VEZExcel.setStatusTip('Сохраняет выбраные ВЭЗ в Excel файл')
        Save1VEZExcel.triggered.connect(lambda:self.Save1VEZExcel(1))

        Save2VEZExcel = QAction("Сохранить Pэ в Excel", self)
        Save2VEZExcel.setShortcut('Ctrl+7')
        Save2VEZExcel.setStatusTip('Сохраняет выбраные Рэ в Excel файл')
        Save2VEZExcel.triggered.connect(lambda:self.Save1VEZExcel(2))

        Save3VEZExcel = QAction("Сохранить Pэ в Word", self)
        Save3VEZExcel.setShortcut('Ctrl+8')
        Save3VEZExcel.setStatusTip('Сохраняет выбраные Рэ в Word файл')
        Save3VEZExcel.triggered.connect(lambda:self.Save1VEZExcel(3))

        EditWord = QAction("Редактировать шаблон Word", self)
        EditWord.setShortcut('Ctrl+G')
        EditWord.setStatusTip('Запускает окно для редактирования шаблона Word')
        EditWord.triggered.connect(self.EditWord)

        zoomIn = QAction("Увеличить изображение", self)
        zoomIn.setShortcut('Ctrl++')
        zoomIn.setStatusTip('Увеличивает открытое изображение')
        zoomIn.triggered.connect(self.PaintForm.zoomIn)

        zoomOut = QAction("Уменьшить изображение", self)
        zoomOut.setShortcut('Ctrl+-')
        zoomOut.setStatusTip('Уменьшает открытое изображение')
        zoomOut.triggered.connect(self.PaintForm.zoomOut)

        Rotate1 = QAction("Повернуть изображение по ч.с", self)
        Rotate1.setShortcut('Ctrl+Shift++')
        Rotate1.setStatusTip('Поворачивает изображение по часовой стрелке')
        Rotate1.triggered.connect(lambda:self.PaintForm.Rotate(1))

        Rotate2 = QAction("Повернуть изображение против ч.с", self)
        Rotate2.setShortcut('Ctrl+Shift+-')
        Rotate2.setStatusTip('Поворачивает изображение ппротив часовой стрелке')
        Rotate2.triggered.connect(lambda:self.PaintForm.Rotate(-1))

        NormalSize = QAction("Вернуть исходный размер", self)
        NormalSize.setShortcut('Ctrl+F')
        NormalSize.setStatusTip('Вернуть исходный размер изображения')
        NormalSize.triggered.connect(self.PaintForm.normalSize)

        
        fileMenu = menubar.addMenu('&Файл')
        fileMenu.addAction(Open0Action)
        fileMenu.addAction(Open1Action)
        fileMenu.addAction(Open2Action)
        fileMenu.addAction(Open3Action)
        fileMenu.addAction(Open4Action)
        fileMenu.addAction(exitAction)

        editMenu = menubar.addMenu('&Правка')
        editMenu.addAction(NewPick)
        editMenu.addAction(DelPick)
        editMenu.addAction(ReRead)
        editMenu.addAction(Select)
        editMenu.addAction(SetLvCheck)
        editMenu.addAction(SetTCheck)
        editMenu.addAction(ClearTable)

        imageMenu = menubar.addMenu('&Изображение')
        imageMenu.addAction(zoomIn)
        imageMenu.addAction(zoomOut)
        imageMenu.addAction(Rotate1)
        imageMenu.addAction(Rotate2)
        imageMenu.addAction(NormalSize)

        reportMenu = menubar.addMenu('&Отчёт')
        reportMenu.addAction(Save1VEZExcel)
        reportMenu.addAction(Save2VEZExcel)
        reportMenu.addAction(Save3VEZExcel)
        reportMenu.addAction(EditWord)
        
         
        
        self.setWindowTitle('VEZRead')
        self.setWindowIcon(QIcon('images\\op1.png'))

        

        self.d = {}
        self.sp_m = []
        self.file_path = []
        self.file_name = []
        self.arr = []
        self.arr_lv = []
        self.arr_t = []
        self.adres = None
        self.ski = QStandardItemModel()
        self.listView.setModel(self.ski)
        self.listView.pressed.connect(self.presseditem)

        self.sel = True
        self.block = True

        #self.NewPick()
        
        

    def Zap(self):
        self.sel = True
        self.adres = None
        self.d = {}
        self.sp_m = []
        self.ski = QStandardItemModel()
        for i in range(len(self.file_name)):
            self.sp_m.append(QStandardItem(QIcon('images\\op1.png'),self.file_name[i]))
            check = (Qt.Checked if self.sel else Qt.Unchecked)
            self.sp_m[i].setCheckable(True)
            self.sp_m[i].setCheckState(check)
            self.ski.appendRow(self.sp_m[i])
            self.d[QPersistentModelIndex(self.sp_m[i].index())]=i
        self.listView.setModel(self.ski)
        

    def Op0(self):
        try:
            p=QFileDialog.getOpenFileNames(self, 'Открыть файлы', self.path_home, "*.png *.jpg *.bmp") 
            if p==([], ''): raise Exception("string index out of range")
            self.file_path = []
            self.file_name = []
            for i in p[0]:
                fname, pr = os.path.splitext(os.path.split(i)[1])
                self.file_path.append(i)
                self.file_name.append(fname)
            self.Zap()
            self.arr =[None for j in self.file_name]
            self.arr_lv = [self.lv_c for j in self.file_name]
            self.arr_t = [self.t_c for j in self.file_name]
   
        except Exception as ex:
            if str(ex) != "string index out of range":
                ems = QErrorMessage(self)
                ems.setWindowTitle('Возникла ошибка')
                ems.showMessage('При открытии возникла ошибка ('+str(ex)+').')


    #QFileDialog.getSaveFileName
    def Op12(self,n):
        try:
            p=""
            p = QFileDialog.getExistingDirectory(self, 'Открыть папку', self.path_home) # Обрати внимание на последний элемент
            if p=="": raise Exception("string index out of range")
            if n==1:
                self.file_path, self.file_name = self.Files(p,n,namef=self.ImageName.text())
            else:
                self.file_path, self.file_name = self.Files(p,n)
            #print(self.file_path)
            self.Zap()
            self.arr =[None for j in self.file_name]
            self.arr_lv = [self.lv_c for j in self.file_name]
            self.arr_t = [self.t_c for j in self.file_name]
        except Exception as ex:
            if str(ex) != "string index out of range":
                ems = QErrorMessage(self)
                ems.setWindowTitle('Возникла ошибка')
                ems.showMessage('При открытии возникла ошибка ('+str(ex)+').')

    def Op3(self):
        try:
            p=QFileDialog.getOpenFileNames(self, 'Открыть файлы', self.path_home, "*.xlsx *.xls") 
            if p==([], ''): raise Exception("string index out of range")
            self.file_path = []
            self.file_name = []
            self.arr_lv = []
            self.arr_t = []
            self.arr =[]
            for i in p[0]:
                N, nam, lv, t = xlsx.OpenFile(i)
                self.arr+=N
                self.file_name+=nam 
                self.file_path+=[None for j in nam] 
                self.arr_lv+=[j if j!= None else self.lv_c for j in lv]
                self.arr_t+=[j if j!= None else self.t_c for j in t]
            self.Zap()
            
        except Exception as ex:
            if str(ex) != "string index out of range":
                ems = QErrorMessage(self)
                ems.setWindowTitle('Возникла ошибка')
                ems.showMessage('При открытии возникла ошибка ('+str(ex)+').')

    def Op4(self):
        try:
            p=""
            p = QFileDialog.getExistingDirectory(self, 'Открыть папку', self.path_home) # Обрати внимание на последний элемент
            if p=="": raise Exception("string index out of range")
            file_path, n = self.Files(p,2)
            self.file_path = []
            self.file_name = []
            self.arr_lv = []
            self.arr_t = []
            self.arr =[]
            for  i in file_path:
                N, nam, lv, t = xlsx.OpenFile(i)
                self.arr+=N
                self.file_name+=nam 
                self.file_path+=[None for j in nam]
                self.arr_lv+=[j if j!= None else self.lv_c for j in lv]
                self.arr_t+=[j if j!= None else self.t_c for j in t]

            self.Zap()
        except Exception as ex:
            if str(ex) != "string index out of range":
                ems = QErrorMessage(self)
                ems.setWindowTitle('Возникла ошибка')
                ems.showMessage('При открытии возникла ошибка ('+str(ex)+').')

    
    #os.path.exists(file_path)
    #os.path.isfile
    def Files(self,pt,on,namef=None):
        p = []
        n = []
        r1 = set([".png",".jpg",".bmp"])
        r2 = set([".xls",".xlsx"])
        r=[r1,r1,r2,r2]
        for root, dirs, files in os.walk(pt+'/'):
            if on == 0 or on == 2:
                for filename in files:
                    fname, pr = os.path.splitext(filename)
                    if pr in r[on]:
                        n.append(fname)
                        p.append(pt+"/"+filename)
            
            else:
                for Dir in dirs:
                    for root, dirs, files1 in os.walk(pt+'/'+Dir+'/'):
                        for f in files1:
                            fname, pr = os.path.splitext(f)
                            if fname == namef and pr in r[on]:
                                n.append(Dir)
                                p.append(pt+'/'+Dir+'/'+f)
                        break
            break
    
        return p, n

    def OpenPict(self,modelindex):
        modelindex=QPersistentModelIndex(modelindex)
        ind = self.d[modelindex]
        self.PaintForm.Update(self.ImFrame.size())
        p = self.file_path[ind]
        if p != None:
            self.PaintForm.open(p)
            if self.arr[ind] == None:
                try:
                    arr = imagescan.Scan(p)
                except Exception:
                    arr = None
                self.arr[ind] = arr
            else:
                arr = self.arr[ind]
        else:
            self.PaintForm.open("images\\image_error_full.png")
            arr = self.arr[ind]
        
        self.block = False
        self.vl.setValue(self.arr_lv[ind])
        self.t.setValue(self.arr_t[ind])
        self.block = True
        self.WriteArrTable(arr)

        try:
            pe1,pe2,pe = roo.RschRoo(arr,self.vl.value(),self.t.value())
            self.pe1_le.setText(str(round(pe1,2))+" Ом")
            self.pe2_le.setText(str(round(pe2,2))+" Ом")
            self.pe_le.setText(str(round(pe,2))+" Ом")
        except Exception:
            self.pe1_le.setText("")
            self.pe2_le.setText("")
            self.pe_le.setText("")

    def WriteArrTable(self,arr):
        try:
            self.Table.setRowCount(0)
            self.Table.setRowCount(30 if 30>len(arr) else len(arr) )
            for i in range(len(arr)):
                for j in range(3):
                    self.Table.setItem(i, j, QTableWidgetItem(str(arr[i][j])))
        except Exception:
            self.Table.setRowCount(0)
            self.Table.setRowCount(30)


    def resizeEvent(self, event):
        self.PaintForm.Update(self.ImFrame.size())

    def ReadTable(self):
        i=0
        N = []
        while True:
            try:
                t = self.Table.item(i,0).text()==""
            except Exception:
                t = True
            if t: break
            a1 = self.Table.item(i,0).text().replace(",",".")
            a1 = float(a1) if a1 !='' else a1
            try:
                a2 = self.Table.item(i,1).text().replace(",",".")
                a2 = float(a2) if a2 !='' else a2
            except Exception:
                a2 =""
            try:
                a3 = self.Table.item(i,2).text().replace(",",".")
                a3 = float(a3) if a3 !='' else a3
            except Exception:
                a3=""
            N.append([a1,a2,a3])
            i+=1
        return N
        
    def NewPick(self):
        self.sp_m.append(QStandardItem(QIcon('images\\op1.png'),"Точка"))
        i=len(self.sp_m)-1
        check = (Qt.Checked if True else Qt.Unchecked)
        self.sp_m[i].setCheckable(True)
        self.sp_m[i].setCheckState(check)
        self.ski.appendRow(self.sp_m[i])
        self.d[QPersistentModelIndex(self.sp_m[i].index())]=i

        self.file_name.append("Точка")
        self.file_path.append(None)
        self.arr.append([])
        self.arr_lv.append(self.lv_c)
        self.arr_t.append(self.t_c)

    def DelPick(self):
        if self.adres in self.d:
            self.ski.removeRow(self.adres.row())
            self.listView.clearSelection()
            del self.d[self.adres]
            self.adres =None
        #print("test")
            

    def presseditem(self, modelindex):
        self.adres=QPersistentModelIndex(modelindex)
        

    def Rename(self):
        ind = self.d[self.adres]
        try:
            modelindex = QModelIndex(self.adres)
            self.file_name[ind] =self.ski.itemFromIndex(modelindex).text()
        except Exception as ex:
            print(ex)

    def WriteTable(self):
        if self.block:
            arr = self.ReadTable()
            if self.adres != None:
                ind = self.d[self.adres]
                self.arr[ind] = arr
                self.arr_lv[ind] = self.vl.value()
                self.arr_t[ind] = self.t.value()
            try:
                pe1,pe2,pe = roo.RschRoo(arr,self.vl.value(),self.t.value())
                self.pe1_le.setText(str(round(pe1,2))+" Ом")
                self.pe2_le.setText(str(round(pe2,2))+" Ом")
                self.pe_le.setText(str(round(pe,2))+" Ом")
            except Exception:
                self.pe1_le.setText("")
                self.pe2_le.setText("")
                self.pe_le.setText("")
    
    def ReReadImage(self):
        if self.adres != None:
            ind = self.d[self.adres]
            p = self.file_path[ind]
            if p != None:
                try:
                    arr = imagescan.Scan(p)
                    self.arr[ind] = arr
                    self.WriteArrTable(arr)
                    self.WriteTable()
                except Exception:
                    1
    
    def ClearTable(self):
        if self.adres != None:
            ind = self.d[self.adres]
            self.arr[ind] = []
        
        self.Table.setRowCount(0)
        self.Table.setRowCount(30)


    def select(self):
        i = 0
        self.sel = not self.sel
        while self.ski.item(i):
            item = self.ski.item(i)
            item.setCheckState(Qt.Checked if self.sel else Qt.Unchecked)
            i += 1

    def SetLvCheck(self):
        i = 0
        while self.ski.item(i):
            item = self.ski.item(i)
            if item.checkState():
                modelindex = QPersistentModelIndex(self.ski.indexFromItem(item))
                ind = self.d[modelindex]
                self.arr_lv[ind] = self.vl.value()
            i+=1

    def SetTCheck(self):
        i = 0
        while self.ski.item(i):
            item = self.ski.item(i)
            if item.checkState():
                modelindex = QPersistentModelIndex(self.ski.indexFromItem(item))
                ind = self.d[modelindex]
                self.arr_t[ind] = self.t.value()
            i+=1


    def ReadImName(self):
        self.imnam = self.ImageName.text()

    def Save1VEZExcel(self,trig=1):
        i = 0
        N=[]
        nm = []
        lv = []
        t = []
        Pe1 = []
        Pe2 = []
        Pe =[]
        try:
            fname = QFileDialog.getSaveFileName(self, 'Сохранить файл', self.path_home,'*.xlsx;;*.xls' if trig !=3 else '*.docx')[0] # Обрати внимание на последний элемент
            while self.ski.item(i):
                item = self.ski.item(i)
                if item.checkState():
                    modelindex = QPersistentModelIndex(self.ski.indexFromItem(item))
                    ind = self.d[modelindex]

                    p = self.file_path[ind]
                    if p != None:
                        if self.arr[ind] == None:
                            try:
                                arr = imagescan.Scan(p)
                            except Exception:
                                arr = None
                        else:
                            arr = self.arr[ind]
                    else:
                        arr = self.arr[ind]

                    if arr != None and arr !=[]:
                        if trig==1:
                            N.append(arr)
                            nm.append(self.file_name[ind])
                            lv.append(self.arr_lv[ind] if self.arr_lv[ind]!=self.lv_c else None)
                            t.append(self.arr_t[ind] if self.arr_t[ind]!=self.t_c else None)
                        elif trig==2 or trig==3:
                            try:
                                pe1,pe2,pe = roo.RschRoo(arr,self.arr_lv[ind],self.arr_t[ind])
                            except Exception:
                                1
                            else:
                                nm.append(self.file_name[ind])
                                lv.append(self.arr_lv[ind])
                                t.append(self.arr_t[ind])
                                Pe1.append(round(pe1,2))
                                Pe2.append(round(pe2,2))
                                Pe.append(round(pe,2))

                i+=1
            if trig==1:
                xlsx.SaveFile1(fname,nm,N,lv,t)
            elif trig==2:
                xlsx.SaveFile2(fname,nm,Pe1,Pe2,Pe,lv,t)
            elif trig==3:
                Word.Word(fname,nm,Pe,lv,t,self.NPrj.text(),self.NVL.text())
        
        except Exception as ex:
            if str(ex) != "string index out of range":
                ems = QErrorMessage(self)
                ems.setWindowTitle('Возникла ошибка')
                ems.showMessage('Не получилось сгенерировать массив точек. '+str(ex))
        else:
            mes = QMessageBox.information(self, 'Генерация массива точек','Операция прошла успешно.',
                                          buttons=QMessageBox.Ok,
                                          defaultButton=QMessageBox.Ok)

                    


                
            


    def closeEvent(self, event):
        Message = QMessageBox(QMessageBox.Question,  'Выход из программы',
            "Вы дейстивлеьно хотите выйти?", parent=self)
        Message.addButton('Да', QMessageBox.YesRole)
        Message.addButton('Нет', QMessageBox.NoRole)
        #Message.addButton('Сохранить', QMessageBox.ActionRole)
        reply = Message.exec()
        if reply == 0:
            qApp.quit()
        elif reply == 1:
            event.ignore()

    def EditWord(self):
        try:
            win32api.ShellExecute(0, 'open', "Word_template\\Template.docx", '', '', 1)
        except Exception as ex:
            if str(ex) == "(2, 'ShellExecute', 'Не удается найти указанный файл.')":
                d = Document()
                d.save("Word_template\\Template.docx")
            else:
                print(ex)



        



if __name__=='__main__':
    app=QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())