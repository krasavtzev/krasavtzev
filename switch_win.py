# coding: utf-8
'''
The program switches between running applications in a loop
Windows with specified time intervals for each application and
emulation of pressing key combinations before and after changing focus.
Программа в цикле выполняет переключения между запущенными приложениями
Windows с указанными интервалами времени по каждому приложению и
эмуляцией нажатия комбинаций клавиш до и после изменения фокуса.
Author: Krasavtsev V 
'''
import time
import warnings
import sys
import threading
from PyQt5 import QtWidgets, QtCore
warnings.simplefilter('ignore', category=UserWarning)
sys.coinit_flags = 2
import pywinauto


class MyWindow(QtWidgets.QWidget):
    ''' Main window. A list of running applications is generated \
and is displayed to the screen. \
After specifying the time parameters, the switching cycle starts'''
    signal_err = QtCore.pyqtSignal(int, str)

    def __init__(self, parent=None):
        ''' initialization of variables'''
        QtWidgets.QWidget.__init__(self, parent)
        self.w_main_n = 0
        self.dict_ru = {'l_win_title': 'Переключение между приложениями \
Windows V 1.3',
                        'l_n_cicl': 'Количество циклов',
                        'l_create_lst': 'Сформировать список',
                        'l_start': 'Запустить переключение',
                        'l_terminate': 'Прервать',
                        'l_edit': 'Редактировать параметры',
                        'l_edit_save': 'Редактирование и сохр. параметров',
                        'l_edit_yes': 'F2-Сохранить',
                        'l_edit_no': 'ESC-Отменить',
                        'l_save_p': 'Сохранить настройки',
                        'l_exit': 'Завершить',
                        'l_err': 'Exception. Возможно закрыли окно.',
                        'l_win_name': 'Название приложения',
                        'l_keys_bef': 'Keys до..',
                        'l_time': 'Время, сек',
                        'l_keys_aft': 'Keys после..',
                        'l_warn': 'Предупреждение',
                        'l_inf': 'Сообщение',
                        'l_spisok': 'Список сочетаний клавиш(разделитель - \
 ";". Список должен начинаться с ";").\
              \n Примеры: F5 обновить вкладку,Ctrl+Tab - Циклическое движение \
вперед по открытым вкладкам в браузере.'}
        self.dict_en = {'l_win_title': 'Switch between applications V 1.3',
                        'l_n_cicl': 'Cycle count',
                        'l_create_lst': 'Create a list',
                        'l_start': 'Start switching',
                        'l_terminate': 'Terminate',
                        'l_edit': 'Edit parameters',
                        'l_edit_save': 'Editing and saving parameters',
                        'l_edit_yes': 'F2-Save',
                        'l_edit_no': 'ESC-Cancel',
                        'l_save_p': 'Save settings',
                        'l_exit': 'Exit',
                        'l_err': 'Exception. Perhaps the window was closed.',
                        'l_win_name': 'Application Name',
                        'l_keys_bef': 'Keys bef..',
                        'l_time': 'Time, sec',
                        'l_keys_aft': 'Keys after..',
                        'l_warn': 'Warning',
                        'l_inf': 'Information',
                        'l_spisok': 'List of keyboard shortcuts\
 (separator - ";". The list must begin with ";")\
              \n Examples: F5 refresh tab, Ctrl+Tab - Cyclic movement \
forward through open tabs in the browser.'}
        self.dict_all = {'Ru': self.dict_ru, 'En': self.dict_en}
        self.save_p = QtCore.QSettings('switch_win.ini', 1)
        if self.save_p.contains('s_lang'):
            self.s_lang = self.save_p.value('s_lang')
        else:
            self.s_lang = 'En'
        self.dict_cur = self.dict_all[self.s_lang]
        self.s_lst_str_new = ''
        self.apps, self.apps_s_m, self.apps_kol_m = [], [], []
        self.windows1 = pywinauto.Desktop(backend="uia").windows()
        self.win_edit_param = QtWidgets.QDialog(self, QtCore.Qt.Dialog)
        self.flag = True
        self.vbox = QtWidgets.QVBoxLayout()
        self.hbox = QtWidgets.QHBoxLayout()
        self.qgrid_w = QtWidgets.QGridLayout()
        self.box = QtWidgets.QGroupBox()
        self.hbox2 = QtWidgets.QHBoxLayout()
        self.resize(800, 680)
        self.kol_cicl = QtWidgets.QLabel(' '*150 + self.dict_cur['l_n_cicl'])
        self.n = QtWidgets.QSpinBox()
        self.cbo_lng = QtWidgets.QComboBox()
        self.cbo_lng.addItems(list(self.dict_all.keys()))
        self.cbo_lng.setCurrentText(self.s_lang)
        self.cbo_lng.currentIndexChanged.connect(self.on_change_lng)
        self.hbox.addWidget(self.kol_cicl)
        self.hbox.addWidget(self.n)
        self.hbox.addWidget(self.cbo_lng)
        self.vbox.addLayout(self.hbox)
        if self.save_p.contains('n'):
            self.n.setValue(int(self.save_p.value('n')))
        else:
            self.n.setValue(3)
        if self.save_p.contains('s_lst'):
            self.s_lst = self.save_p.value('s_lst').split(';')
        else:
            self.s_lst = ['   ', 'F5', '^{TAB}']
        self.s_lst_flag = 1
        self.n.setMaximumSize(50, 20)
        self.n.setRange(0, 999)
        self.setLayout(self.vbox)

    def on_change(self, kod_signal, lbl_s):
        ''' displaying an error or shutdown message \
programs (based on signals from the stream)'''
        self.windows1[self.w_main_n].set_focus()
        if kod_signal == 0:
            QtWidgets.QMessageBox.information(self, self.dict_cur['l_warn'],
                                              lbl_s,
                                              buttons=QtWidgets.QMessageBox.Ok)
        if kod_signal == 1:
            QtWidgets.QMessageBox.information(self, self.dict_cur['l_inf'],
                                              lbl_s,
                                              buttons=QtWidgets.QMessageBox.Ok)

    def on_change_lng(self):
        ''' calling screen refresh after changing language'''
        self.s_lang = self.cbo_lng.currentText()
        self.dict_cur = self.dict_all[self.s_lang]
        self.form_spis0()

    def form_spis0(self):
        ''' erasing the old list and calling for the formation of a new one'''
        while self.qgrid_w.count() > 0:
            self.qgrid_w.itemAt(0).widget().setParent(None)
        while self.hbox2.count() > 0:
            self.hbox2.itemAt(0).widget().setParent(None)
        self.hbox2.removeWidget(self.box)
        self.qgrid_w = QtWidgets.QGridLayout()
        form_spis(self)
        self.vbox.addLayout(self.qgrid_w)
        self.flag = True
        self.form_hbox2()
        self.setLayout(self.vbox)
        self.adjustSize()

    def form_hbox2(self):
        ''' creating buttons on the screen'''
        self.hbox2 = QtWidgets.QHBoxLayout()
        self.box = QtWidgets.QGroupBox()
        self.box.setMaximumSize(750, 50)
        btn_add1 = QtWidgets.QPushButton(self.dict_cur['l_create_lst'])
        btn_add1.clicked.connect(self.form_spis0)
        btn_add2 = QtWidgets.QPushButton(self.dict_cur['l_start'])
        btn_add2.clicked.connect(self.swith_wins0)
        btn_add3 = QtWidgets.QPushButton(self.dict_cur['l_terminate'])
        btn_add3.clicked.connect(self.swith_stop)
        btn_add4 = QtWidgets.QPushButton(self.dict_cur['l_edit'])
        btn_add4.clicked.connect(self.edit_param)
        btn_add5 = QtWidgets.QPushButton(self.dict_cur['l_save_p'])
        btn_add5.clicked.connect(self.save_parametrs)
        btn_add6 = QtWidgets.QPushButton(self.dict_cur['l_exit'])
#        btn_add6.clicked.connect(QtWidgets.qApp.quit)
        btn_add6.clicked.connect(self.close)
        self.hbox2.addWidget(btn_add1)
        self.hbox2.addWidget(btn_add2)
        self.hbox2.addWidget(btn_add3)
        self.hbox2.addWidget(btn_add4)
        self.hbox2.addWidget(btn_add5)
        self.hbox2.addWidget(btn_add6)
        self.box.setLayout(self.hbox2)
        self.vbox.addWidget(self.box)

    def swith_wins0(self):
        ''' starting a switch in a thread'''
        self.flag = True
        threading.Thread(target=self.swith_wins).start()

    def swith_wins(self):
        ''' main cycle of switching between applications'''
        n_cicl = 0
        while n_cicl < self.n.value() and self.flag is True:
            for i in range(len(self.windows1)):
                if self.apps_kol_m[i].value() > 0:
                    try:
                        self.windows1[i].set_focus()
                    except Exception:
                        s_err = self.dict_cur['l_err'] + ' ' +\
                                self.apps_s_m[i].text()
                        self.signal_err.emit(0, s_err)
                    perem_bef = self.apps_bef[i].currentText().strip()
                    if perem_bef > ' ':
                        pywinauto.keyboard.send_keys(perem_bef)
                    time.sleep(self.apps_kol_m[i].value())
                    try:
                        self.windows1[i].set_focus()
                    except Exception:
                        s_err = self.dict_cur['l_err'] + ' ' +\
                                self.apps_s_m[i].text()
                        self.signal_err.emit(0, s_err)
                    perem_aft = self.apps_aft[i].currentText().strip()
                    if perem_aft > ' ':
                        pywinauto.keyboard.send_keys(perem_aft)
            n_cicl += 1
        if n_cicl == self.n.value() and self.flag is True or self.flag is None:
            s_err = self.dict_cur['l_n_cicl'] + ' : ' + str(n_cicl)
            self.signal_err.emit(1, s_err)

    def swith_stop(self):
        ''' interruption of the switching cycle'''
        self.flag = None

    def edit_param(self):
        ''' editing parameters: keyboard shortcuts'''
        self.win_edit_param.setWindowTitle(self.dict_cur['l_edit_save'])
        self.win_edit_param.resize(600, 400)
        self.win_edit_param.setWindowModality(QtCore.Qt.WindowModal)
        self.win_edit_param.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        vbox = QtWidgets.QVBoxLayout()
        p_text = '''
https://pywinauto.readthedocs.io/en/latest/code/pywinauto.keyboard.html

Available key codes:

{SCROLLLOCK}, {VK_SPACE}, {VK_LSHIFT}, {VK_PAUSE}, {VK_MODECHANGE},
{BACK}, {VK_HOME}, {F23}, {F22}, {F21}, {F20}, {VK_HANGEUL}, {VK_KANJI},
{VK_RIGHT}, {BS}, {HOME}, {VK_F4}, {VK_ACCEPT}, {VK_F18}, {VK_SNAPSHOT},
{VK_PA1}, {VK_NONAME}, {VK_LCONTROL}, {ZOOM}, {VK_ATTN}, {VK_F10}, {VK_F22},
{VK_F23}, {VK_F20}, {VK_F21}, {VK_SCROLL}, {TAB}, {VK_F11}, {VK_END},
{LEFT}, {VK_UP}, {NUMLOCK}, {VK_APPS}, {PGUP}, {VK_F8}, {VK_CONTROL},
{VK_LEFT}, {PRTSC}, {VK_NUMPAD4}, {CAPSLOCK}, {VK_CONVERT}, {VK_PROCESSKEY},
{ENTER}, {VK_SEPARATOR}, {VK_RWIN}, {VK_LMENU}, {VK_NEXT}, {F1}, {F2},
{F3}, {F4}, {F5}, {F6}, {F7}, {F8}, {F9}, {VK_ADD}, {VK_RCONTROL},
{VK_RETURN}, {BREAK}, {VK_NUMPAD9}, {VK_NUMPAD8}, {RWIN}, {VK_KANA},
{PGDN}, {VK_NUMPAD3}, {DEL}, {VK_NUMPAD1}, {VK_NUMPAD0}, {VK_NUMPAD7},
{VK_NUMPAD6}, {VK_NUMPAD5}, {DELETE}, {VK_PRIOR}, {VK_SUBTRACT}, {HELP},
{VK_PRINT}, {VK_BACK}, {CAP}, {VK_RBUTTON}, {VK_RSHIFT}, {VK_LWIN}, {DOWN},
{VK_HELP}, {VK_NONCONVERT}, {BACKSPACE}, {VK_SELECT}, {VK_TAB}, {VK_HANJA},
{VK_NUMPAD2}, {INSERT}, {VK_F9}, {VK_DECIMAL}, {VK_FINAL}, {VK_EXSEL},
{RMENU}, {VK_F3}, {VK_F2}, {VK_F1}, {VK_F7}, {VK_F6}, {VK_F5}, {VK_CRSEL},
{VK_SHIFT}, {VK_EREOF}, {VK_CANCEL}, {VK_DELETE}, {VK_HANGUL}, {VK_MBUTTON},
{VK_NUMLOCK}, {VK_CLEAR}, {END}, {VK_MENU}, {SPACE}, {BKSP}, {VK_INSERT},
{F18}, {F19}, {ESC}, {VK_MULTIPLY}, {F12}, {F13}, {F10}, {F11}, {F16},
{F17}, {F14}, {F15}, {F24}, {RIGHT}, {VK_F24}, {VK_CAPITAL}, {VK_LBUTTON},
{VK_OEM_CLEAR}, {VK_ESCAPE}, {UP}, {VK_DIVIDE}, {INS}, {VK_JUNJA},
{VK_F19}, {VK_EXECUTE}, {VK_PLAY}, {VK_RMENU}, {VK_F13}, {VK_F12}, {LWIN},
{VK_DOWN}, {VK_F17}, {VK_F16}, {VK_F15}, {VK_F14}

~ is a shorter alias for {ENTER}
'+': {VK_SHIFT}
'^': {VK_CONTROL}
'%': {VK_MENU} a.k.a. Alt key
Example how to use modifiers:

send_keys('^a^c') # select all (Ctrl+A) and copy to clipboard (Ctrl+C)
send_keys('+{INS}') # insert from clipboard (Shift+Ins)
send_keys('%{F4}') # close an active window with Alt+F4
Repetition count can be specified for special keys.
{ENTER 2} says to press Enter twice.

Example which shows how to press and hold or release a key on the keyboard:
  send_keys("{VK_SHIFT down}"
          "pywinauto"
          "{VK_SHIFT up}") # to type PYWINAUTO

send_keys("{h down}"
          "{e down}"
          "{h up}"
          "{e up}"
          "llo") # to type hello
Use curly brackers to escape modifiers and type reserved symbols
as single keys:

send_keys('{^}a{^}c{%}') # type string "^a^c%" (Ctrl will not be pressed)
send_keys('{{}ENTER{}}') # type string "{ENTER}" without pressing Enter key
'''
        f_text = QtWidgets.QPlainTextEdit(p_text)
        f_text.setReadOnly(True)
        spisok = QtWidgets.QLabel(self.dict_cur['l_spisok'])
        self.s_lst_flag = 2
        s_lst_str_old = ''
        for i in self.s_lst:
            s_lst_str_old = s_lst_str_old + i + ';'
        s_lst_str_old = s_lst_str_old[:-1]  # убрали заключительную ;
        self.s_lst_str_new = QtWidgets.QLineEdit(s_lst_str_old)
        vbox.addWidget(f_text)
        vbox.addWidget(spisok)
        vbox.addWidget(self.s_lst_str_new)
        self.win_edit_param.setLayout(vbox)
        end_edit = QtWidgets.QDialogButtonBox(QtCore.Qt.Horizontal)
        btn_save = end_edit.addButton(self.dict_cur['l_edit_yes'],
                                      QtWidgets.QDialogButtonBox.AcceptRole)
        btn_save.setShortcut('F2')
        btn_esc = end_edit.addButton(self.dict_cur['l_edit_no'],
                                     QtWidgets.QDialogButtonBox.RejectRole)
        btn_esc.setShortcut('Esc')
        end_edit.accepted.connect(self.save_parametrs)
        end_edit.rejected.connect(self.no_edit)
        vbox.addWidget(end_edit)
        self.win_edit_param.exec()

    def save_parametrs(self):
        '''Saving parameters: number of cycles, language, values\
            keys before and after the window, window display time.'''
        self.save_p.setValue('n', self.n.value())
        if self.s_lst_flag == 1:
            s_lst_str_n = ''
            for i in self.s_lst:
                s_lst_str_n = s_lst_str_n + i + ';'
            s_lst_str_n = s_lst_str_n[:-1]
            self.save_p.setValue('s_lst', s_lst_str_n)
        else:
            self.s_lst = self.s_lst_str_new.text().split(';')
            self.save_p.setValue('s_lst', self.s_lst_str_new.text())
        j = 0
        s_apps, s_apps_kol_m, s_apps_bef, s_apps_aft = '', '', '', ''
        for w in self.windows1:
            if self.apps_kol_m[j].value() > 0:
                s_apps += w.window_text() + ';'
                s_apps_kol_m += str(self.apps_kol_m[j].value()) + ';'
                s_apps_bef += self.apps_bef[j].currentText().strip() + ';'
                s_apps_aft += self.apps_aft[j].currentText().strip() + ';'
            j += 1
        s_apps = s_apps[:-1]   # убираем последний разделитель
        s_apps_kol_m = s_apps_kol_m[:-1]
        s_apps_bef = s_apps_bef[:-1]
        s_apps_aft = s_apps_aft[:-1]
        self.save_p.setValue('s_apps', s_apps)
        self.save_p.setValue('s_apps_bef', s_apps_bef)
        self.save_p.setValue('s_apps_kol_m', s_apps_kol_m)
        self.save_p.setValue('s_apps_aft', s_apps_aft)
        self.save_p.setValue('s_lang', self.cbo_lng.currentText())
        self.save_p.sync()
        if self.s_lst_flag == 2:
            self.win_edit_param.close()
        self.s_lst_flag = 1
        self.form_spis0()

    def no_edit(self):
        '''closing the parameter editing window without making changes'''
        self.win_edit_param.close()


def form_spis(self):
    ''' creating a list'''
    self.flag = True
    self.apps, self.apps_s_m, self.apps_kol_m = [], [], []
    self.apps_bef, self.apps_aft = [], []
    self.windows1 = pywinauto.Desktop(backend="uia").windows()
    s_apps, s_apps_kol_m, s_apps_bef, s_apps_aft = '', '', '', ''
    if self.save_p.contains('s_apps'):
        s_apps = self.save_p.value('s_apps').split(';')
    if self.save_p.contains('s_apps_bef'):
        s_apps_bef = self.save_p.value('s_apps_bef').split(';')
    if self.save_p.contains('s_apps_kol_m'):
        s_apps_kol_m = self.save_p.value('s_apps_kol_m').split(';')
    if self.save_p.contains('s_apps_aft'):
        s_apps_aft = self.save_p.value('s_apps_aft').split(';')
    j = 0
    smesh = 1
    self.w_main = self.dict_cur['l_win_title']
    self.setWindowTitle(self.w_main)
    self.kol_cicl.setText(' '*150 + self.dict_cur['l_n_cicl'])
    self.qgrid_w.setVerticalSpacing(13)
    self.field1 = QtWidgets.QLabel(self.dict_cur['l_win_name'])
    self.field1.setStyleSheet('font-weight: bold')
    self.qgrid_w.addWidget(self.field1, 0, 0)
    self.field2 = QtWidgets.QLabel(self.dict_cur['l_keys_bef'])
    self.field2.setStyleSheet('font-weight: bold')
    self.qgrid_w.addWidget(self.field2, 0, 1)
    self.field3 = QtWidgets.QLabel(self.dict_cur['l_time'])
    self.field3.setStyleSheet('font-weight: bold')
    self.qgrid_w.addWidget(self.field3, 0, 2)
    self.field4 = QtWidgets.QLabel(self.dict_cur['l_keys_aft'])
    self.field4.setStyleSheet('font-weight: bold')
    self.qgrid_w.addWidget(self.field4, 0, 3)
    for w in self.windows1:
        perem = w.window_text()
        if perem == self.w_main:
            self.w_main_n = j
        self.apps.append(pywinauto.Application(backend='uia').
                         connect(handle=w.handle))
        self.apps_s_m.append(QtWidgets.QLabel(perem))
        self.qgrid_w.addWidget(self.apps_s_m[j], j+smesh, 0)
        self.cbo = QtWidgets.QComboBox()
        self.cbo.addItems(self.s_lst)
        self.apps_bef.append(self.cbo)
        self.apps_kol_m.append(QtWidgets.QSpinBox())
        if perem in s_apps:
            self.apps_kol_m[j].setValue(int(s_apps_kol_m[s_apps.index(perem)]))
            self.cbo.setCurrentText(s_apps_bef[s_apps.index(perem)])
        else:
            self.apps_kol_m[j].setValue(0)
        self.qgrid_w.addWidget(self.apps_bef[j], j+smesh, 1)
        self.apps_kol_m[j].setRange(0, 9999)
        self.apps_kol_m[j].setMaximumSize(60, 20)
        self.qgrid_w.addWidget(self.apps_kol_m[j], j+smesh, 2)
        self.cbo = QtWidgets.QComboBox()
        self.cbo.addItems(self.s_lst)
        self.apps_aft.append(self.cbo)
        if perem in s_apps:
            self.cbo.setCurrentText(s_apps_aft[s_apps.index(perem)])
        self.qgrid_w.addWidget(self.apps_aft[j], j+smesh, 3)
        j += 1


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    main_window = MyWindow()
    main_window.show()
    main_window.signal_err.connect(main_window.on_change,
                                   QtCore.Qt.QueuedConnection)
    main_window.form_spis0()
    sys.exit(app.exec_())
