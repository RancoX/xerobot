import sys,json,csv, subprocess
from datetime import datetime, timedelta
from pathlib import Path
from PySide6.QtGui import QAction,QIcon
from PySide6.QtWidgets import (QWidget,QApplication,QMainWindow,QVBoxLayout,QHBoxLayout,
                               QPushButton,QToolButton,QCheckBox,QLabel,QScrollArea,
                               QLineEdit,QFormLayout,QDialog,QListView,QMessageBox,QFileDialog)
from PySide6.QtCore import QSize,Qt,QAbstractListModel
from xero_bank_reconciler import MAX_WAIT_TIME,URL,USERNAME,PASSWORD,create_driver,login,wait_til_get_elem

BASE_DIR = Path(__file__).resolve().parent

# load in last saved path
with open(BASE_DIR / 'data/save_path.json','r') as rf:
    try:
        SAVE_DIR = Path(json.load(rf)['saved_path'])
        if SAVE_DIR == '':
            SAVE_DIR = BASE_DIR.parent.parent
    except:
        SAVE_DIR = BASE_DIR.parent.parent

# load in logins
with open(BASE_DIR / 'data/credential.json','r') as rf:
    cred = json.load(rf)

# load entity list
def read_entity_list(folder: Path):
    with open(folder / 'data/entities.json','r') as rf:
        entity_dict = json.load(rf)
    
    all_entities=['placeholder',*[x for x in entity_dict.values()]]

    return entity_dict,all_entities

entity_dict, all_entities = read_entity_list(BASE_DIR)


class DictionaryListModel(QAbstractListModel):
    def __init__(self, data=None):
        super().__init__()
        self.data_dict = data or {}

    def data(self, index, role):
        if role == Qt.ItemDataRole.DisplayRole:
            return str(list(self.data_dict.keys())[index.row()])

    def rowCount(self, index):
        return len(self.data_dict)


class EditDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.layout = QVBoxLayout()
        row=QFormLayout()

        self.trust_label = QLabel("Enter Trust Full Name:")
        self.trust = QLineEdit()
        row.addRow(self.trust_label,self.trust)

        self.code_label = QLabel("Enter 6-character Code:")
        self.code = QLineEdit()
        row.addRow(self.code_label,self.code)
        self.layout.addLayout(row)

        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(self.accept)
        self.layout.addWidget(self.save_button)

        self.setLayout(self.layout)

    def get_new_key(self):
        return self.trust.text()

    def get_new_value(self):
        return self.code.text()


class ListDialog(QDialog):
    def __init__(self,parent=None):
        super().__init__(parent)

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Dictionary List View")

        self.central_widget = QWidget(self)
        self.setGeometry(600,200,273,350)

        self.layout = QVBoxLayout()

        self.list_view = QListView()
        self.dict_model = DictionaryListModel(entity_dict)
        self.list_view.setModel(self.dict_model)
        
        self.layout.addWidget(self.list_view)

        self.add_button = QPushButton("Add Item")
        self.add_button.clicked.connect(self.add_item)
        self.layout.addWidget(self.add_button)

        self.update_button = QPushButton("Update Selected Item")
        self.update_button.clicked.connect(self.update_item)
        self.layout.addWidget(self.update_button)

        self.delete_button = QPushButton("Delete Selected Item")
        self.delete_button.clicked.connect(self.delete_item)
        self.layout.addWidget(self.delete_button)

        self.save_button = QPushButton("Save All")
        self.save_button.clicked.connect(self.save_current_list)
        self.layout.addWidget(self.save_button)

        self.central_widget.setLayout(self.layout)

    def add_item(self):
        dialog = EditDialog(self)
        if dialog.exec():
            new_key = dialog.get_new_key()
            new_value = dialog.get_new_value()
            self.dict_model.data_dict[new_key] = new_value
            self.dict_model.layoutChanged.emit()

    def update_item(self):
        selected_index = self.list_view.selectionModel().currentIndex()
        if selected_index.isValid():
            selected_key = list(self.dict_model.data_dict.keys())[selected_index.row()]
            dialog = EditDialog(self)
            dialog.trust.setText(selected_key)
            dialog.code.setText(self.dict_model.data_dict[selected_key])
            if dialog.exec():
                new_key = dialog.get_new_key()
                new_value = dialog.get_new_value()
                self.dict_model.data_dict[new_key] = new_value
                self.dict_model.dataChanged.emit(selected_index, selected_index)

    def delete_item(self):
        selected_index = self.list_view.selectionModel().currentIndex()
        if selected_index.isValid():
            selected_key = list(self.dict_model.data_dict.keys())[selected_index.row()]
            del self.dict_model.data_dict[selected_key]
            self.dict_model.layoutChanged.emit()


    def save_current_list(self):
        with open(BASE_DIR / 'data/entities.json','w') as wf:
            json.dump(self.dict_model.data_dict,wf)
        QMessageBox.information(self,'Save success','The current list is saved.')


class MyMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # meta setups
        self.setWindowTitle('Xero bank reconciler 1.6')
        # Set the window icon
        icon_path=BASE_DIR / "image/argyle.jpg"
        self.setWindowIcon(QIcon(str(icon_path)))
        self.setMinimumSize(QSize(600,400))
        
        # define layout
        # initialize layout
        main_layout=QVBoxLayout()
        form_layout = QFormLayout()
        save_path_layout=QHBoxLayout()
        btn_layout=QHBoxLayout()
        display_layout=QVBoxLayout()

        # define widgets in different sections
        # menu bar
        main_menu = self.menuBar()
        about_menu = main_menu.addMenu("File")
        about_action = QAction("About",self)
        about_action.triggered.connect(self.show_about_info)
        about_menu.addAction(about_action)

        # username section
        user_label=QLabel('Email: ')
        self.username=QLineEdit()
        self.username.setPlaceholderText('enter your email')
        form_layout.addRow(user_label,self.username)

        # password section
        pw_label=QLabel('Password: ')
        self.pw=QLineEdit()
        self.pw.setEchoMode(QLineEdit.EchoMode.Password)
        self.pw.setPlaceholderText('enter your password')
        form_layout.addRow(pw_label,self.pw)
        form_layout.setContentsMargins(0,0,0,20)

        # csv save path
        path_label=QLabel('csv saved at: ')
        self.save_path=QLineEdit(str(SAVE_DIR))
        btn3=QPushButton('...')
        btn3.setMaximumSize(25,25)
        btn3.setToolTip('Select folder for csv')
        btn3.clicked.connect(self.open_folder_selector)
        btn4=QToolButton()
        btn4.setIcon(QIcon(str(BASE_DIR / 'image/folder.png')))
        btn4.setMaximumSize(25,25)
        btn4.setToolTip('Open folder')
        btn4.clicked.connect(self.open_folder)
        save_path_layout.addWidget(self.save_path)
        save_path_layout.addWidget(btn3)
        save_path_layout.addWidget(btn4)
        form_layout.addRow(path_label,save_path_layout)

        # element to find
        element_label=QLabel('Element to find: ')
        self.eleName=QLineEdit('bankWidget-balanceTable__summary--label--KiKPd')
        form_layout.addRow(element_label,self.eleName)

        # press enter to execute
        self.username.returnPressed.connect(self.execute_func)
        self.pw.returnPressed.connect(self.execute_func)

        # checkbox and login button section
        checkbox=QCheckBox('Autofill with the last credentials')
        checkbox.stateChanged.connect(self.autofill)
        checkbox.setChecked(False)
        checkbox.setChecked(True)

        btn=QPushButton('Execute')
        btn2=QPushButton('Check entity')
        btn.clicked.connect(self.execute_func)
        btn2.clicked.connect(self.check_entity_list)

        btn_layout.addWidget(checkbox)
        btn_layout.addWidget(btn)
        btn_layout.addWidget(btn2)

        # display section
        self.scroll=QScrollArea()
        self.display=QLabel('Waiting...')
        self.display.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.display.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.scroll.setWidget(self.display)
        self.scroll.setWidgetResizable(True)
        display_layout.addWidget(self.scroll)

        # statusBar
        self.statusbar=self.statusBar()
        self.update_status_bar()

        # set container widget
        container=QWidget()

        # nest layouts
        main_layout.addLayout(form_layout)
        main_layout.addLayout(btn_layout)
        main_layout.addLayout(display_layout)

        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def autofill(self,stateNum):
        # if checkbox is currently checked then autofill
        if stateNum == 2:
            self.username.setText(USERNAME)
            self.pw.setText(PASSWORD)

    def check_user_pw(self):
        # check if login info changed
        current_user=self.username.text()
        current_pw=self.pw.text()
        return current_user != USERNAME or current_pw != PASSWORD
    
    def update_user_pw(self):
        if self.check_user_pw():
            # create new login info
            new_login={'username':self.username.text(), 'password':self.pw.text()}
            self.display.setText('Login info changed, overriding the previous credentials...')
            # override credential.json
            with open(BASE_DIR / 'data/credential.json','w') as wf:
                json.dump(new_login,wf)
            self.update_display_text('Login credential updated.')

    def update_display_text(self,new_text:str):
        self.display.setText(self.display.text() + '\n' + new_text)
        self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum())
        self.display.repaint()

    def update_status_bar(self,msg='Ready'):
        self.statusbar.showMessage(msg)
        self.statusbar.repaint()

    def show_about_info(self):
        self.display.clear()
        about_info = '''This software was developed by Argyle Investment Data team in Aug 2023 by Ranco Xu.\nIt scrapes data from Xero dashboard once user logged in.\n
    Update log 1.2 @2023-08-30:
        1) Now user can press enter after typing to start the program
        2) Added 'Check entity' page to View/Add/Update/Delete current trust names and codes
        3) Fixed the bug that after typing user login credentials the web browser still uses Liam's info
        
    Update log 1.3 @2023-09-05:
        1) Fixed a bug that caused empty outputs due to single digit day format in date string

    Update log 1.4 @2023-09-05:
        1) Added functionality which allows user to define csv save path

    Update log 1.5 @2023-10-16:
    1) Fixed the problem that the newly added entity is not picked up by the bot
    2) Added a button to open destination folder

    Update log 1.6 @2024-01-30:
    1) Fixed the negative back reconciliation in bracket format which caused error
        '''
        self.update_display_text(about_info)

    def check_entity_list(self,s):
        self.entity_window=ListDialog(self)
        self.entity_window.finished.connect(self.entity_window.save_current_list)
        self.entity_window.exec()

    def open_folder_selector(self,s):
        dialog = QFileDialog(self)
        dialog.setDirectory(self.save_path.text())
        dialog.setFileMode(QFileDialog.FileMode.Directory)
        dialog.setViewMode(QFileDialog.ViewMode.List)
        if dialog.exec():
            foldernames = dialog.selectedFiles()
            if foldernames:
                SAVE_DIR = foldernames[0]
                self.update_save_path(SAVE_DIR)

    def open_folder(self,s):
        folder_path = Path(self.save_path.text())
        subprocess.Popen(f'explorer {folder_path}', shell=True)

    def update_save_path(self,saved_path,save=False):
        self.save_path.setText(str(saved_path))
        with open(BASE_DIR / 'data/save_path.json','w') as wf:
            data={'saved_path':str(saved_path)}
            json.dump(data,wf)

    def execute_func(self):
        # initialize container
        container=[['farm','date','item','value','link']]
        total_farm=0
        farm_incre_flag=True
        total_records=0
        self.update_display_text(', '.join(list(map(str,container[-1]))))
        print(', '.join(list(map(str,container[-1]))))

        # initialize driver
        driver=create_driver(BASE_DIR)

        # get today's date
        reference_day=datetime.today()-timedelta(days=1)
        today=reference_day.strftime('%Y-%m-%d')
        today_time=reference_day.strftime('%Y-%m-%d %H:%M:%S')
        today_date=reference_day.strftime('(%b %#d)')
        self.display.setText(f"program running...\ncurrent time is {today_time}, parsing records only on {today_date}\n")
        self.display.repaint()

        _, all_entities = read_entity_list(BASE_DIR)
        for i,entity in enumerate(all_entities):
            if i==0:
                # open website
                driver.get(URL)
                if self.check_user_pw():
                    login(driver,self.username.text(),self.pw.text())
                    self.update_user_pw()
                else:
                    login(driver,USERNAME,PASSWORD)
                farm=wait_til_get_elem(driver,'xui-pageheading--title',MAX_WAIT_TIME,'class')[0].text.replace('\n','')
                # after successfuly login, check if new login credential was used, if so save it
                self.check_user_pw()
            else:
                DASHBOARD_URL=f'https://go.xero.com/app/{entity}/dashboard'
                driver.get(DASHBOARD_URL)

                farm=wait_til_get_elem(driver,'xui-pageheading--title',MAX_WAIT_TIME,'class')[0].text.replace('\n','')
                statement_bal=wait_til_get_elem(driver,self.eleName.text(),MAX_WAIT_TIME,'class')
                # statement balance are all elements with class name: bankWidget-balanceTable__summary--label--N4zpS
                for j,found in enumerate(statement_bal):
                    item=found.text.strip()
                    
                    if item.lower().startswith('statement balance') and item.endswith(today_date):
                        # increment toatl_farm
                        if farm_incre_flag:
                            total_farm+=1
                            farm_incre_flag=False
                        # increment total_records
                        total_records+=1
                        val=float(statement_bal[j+1].text.replace(',','').replace('(','-').replace(')',''))
                        # save to container
                        container.append([farm,today,item,val,DASHBOARD_URL])
                        self.update_display_text(', '.join(list(map(str,container[-1]))))
                        print(', '.join(list(map(str,container[-1]))))
                        self.update_status_bar(f"Progress: {(i+1)/len(all_entities):.2%}")

                # reset farm_incre_flag
                farm_incre_flag=True

        self.update_display_text(f'\n{"="*80}\nAll data retrieved, a total of {total_records} records found from {total_farm} farms.\nNow saving to csv...')
        print((f'\n{"="*80}\nAll data retrieved, a total of {total_records} records found from {total_farm} farms.\nNow saving to csv...'))

        # write to csv
        fpath=Path(self.save_path.text()) / f'statement_balance.csv'
        with open(fpath,mode='w',newline='',encoding='utf-8') as wf:
            writer=csv.writer(wf)
            writer.writerows(container)

        self.update_display_text(f'Data saved to {fpath}')
        print(f'Data saved to {fpath}')

if __name__=='__main__':
    app=QApplication(sys.argv)

    window=MyMainWindow()
    window.show()

    app.exec()