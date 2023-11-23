from PyQt5.QtWidgets import QDialog, QApplication, QHBoxLayout, QCheckBox, QLabel, QPlainTextEdit, QSizePolicy, QFrame, \
    QGridLayout
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot, QTimer, QSize, Qt
from PyQt5.uic import loadUi
import serial
import serial.tools.list_ports as port
import time
import pandas as pd
import os
import datetime
from time import gmtime, strftime
import jinja2
import tpl

'''
WELCOME TO BOARD TESTING CODE :) 

Author: PARNIKA GUPTA
Date: August 1st 2020, September 27th 2020
Project: GAIA Testing-Automation GUI 
Version: 5.1
Technology: Qt (Windows, MAC, Linux)
Features: Multiple project (AAI | Insight) with seperate look-up tables
                1. Latest libraries upgraded
                2. GAIA png logo added
                3. Copyright included - COPYRIGHT © 2020 Gaia Smart Cities Solutions Pvt. Ltd
                4. Re-designed- In horizontal View 
                5. 23 get commands (Validation check for all 13 get response commands | Updatable)
                6. 25 set commands (Validation check for all | Updatable)
Total Functions: 22 functions installed 
                1.  'Info'                                        Instruction 
                2.  'Disconnected'/'Connected'                    Connection error or not
                3.  'Send' & 4. 'Receive'                         Single GET command
                5.  'Check'                                       Quick single GET command
                6.  'Auto Send' & 7. 'Auto Receive'               Multiple GET commands
                8.  'Repeat Sent-Receive'                         Repetitive multiple GET commands
                9.  'SET Device Id'                               Only set device id of multiple devices 
                10. 'Select All'                                  Toggle selection on SET command
                11. 'SET'                                         Only SET commands
                12. 'SET Check'                                   SET commands with validation
                13. 'Repeat SET Check'                            Repetitive multiple SET commands with validation
                14. 'SET-GET Response'                            Runs whole GUI (SET and then GET) 
                15. 'Analyse GET Data' & 16. 'Analyse SET Data'   Data analysis of GET or SET Records 
                17. 'Empty GET Data' & 18. 'Empty SET Data'       Renew GET or SET Records 
                19. 'Show' & 20. 'Delete'                         Show or Delete all msgs of GSM module
                21. 'Clear Screen'                                Clear Command Response screen
                22. 'Quit'                                        Close App 
                
Additional Features:
1. Finds the serial port automatically.
2. Tells sms is unsuccessful if "ERROR" is found and "+CMGS: " is not found

Upgrade in 5.0:
1. Device ID has receiving feature too
In App: CLI- python Setup.py build-> Compress UI Folder-> Upload in NSIS App-> Setup Ready

'''

# for clearing the terminal screen from a Python script in case executed on terminal
def clear():
    # _ = os.system('cls') if os.name == 'nt' else os.system('clear')
    os.system('cls' if os.name == 'nt' else 'clear')


# class for the GUI functions with send and receive features
class Widget(QDialog):
    @pyqtSlot()
    def __init__(self):
        super(Widget, self).__init__()
        # load the GUI designed from QT Designer
        loadUi('gui._designui', self)

        # app title and icon
        self.setWindowTitle("Testing Automation GUI")
        self.iconName = "company_logo.png"
        self.setWindowIcon(QIcon(self.iconName))

        # excel sheets updating initialisation
        self.Database = 'DATABASE.xlsx'
        self.Record_File = 'GET Response.xlsx'
        self.Record_File_Common_Sheet = 'RECORD'
        self.set_Record_File = 'SET Response.xlsx'

        # info definition
        self.output_te.clear()
        self.output_te.append("")
        self.output_te.append("<b>WELCOME TO BOARD TESTING :) </b>")
        self.info()

        # excel sheet Gaia project database
        self.file_default_values = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['DEFAULT VALUES'])
                                        .dropna(how="any")['DEFAULT VALUES'])
        GAIA_Project = self.file_default_values[10]
        print("GAIA Project Used: " + GAIA_Project)
        try:
            if GAIA_Project!="_":
                self.Database = 'DATABASE_' + GAIA_Project + '.xlsx'
                self.Record_File = 'GET Response_' + GAIA_Project + '.xlsx'
                self.Record_File_Common_Sheet = 'RECORD_' + GAIA_Project
                self.set_Record_File = 'SET Response_' + GAIA_Project + '.xlsx'
            else:
                print("GAIA Project Used: Common")

            print("Database: " + self.Database)
            print("Record_File: " + self.Record_File)
            print("Record_File_Common_Sheet: " + self.Record_File_Common_Sheet)
            print("set_Record_File: " + self.set_Record_File)
            print("Database altered :)")
            print("\n")

            self.output_te.append("<b>Database: </b>" + self.Database)
            self.output_te.append("<b>Record_File: </b>" + self.Record_File)
            self.output_te.append("<b>Record_File_Common_Sheet: </b>" + self.Record_File_Common_Sheet)
            self.output_te.append("<b>set_Record_File: </b>" + self.set_Record_File)

            self.output_te.append("<b>Database altered :)</b>")
            self.output_te.append("\n")
        except Exception as e:
            print('Error in send', e)
            self.output_te.append('<b>\nError in select_database: ' + str(e))

        self.common_dir_parameter = ['NUMBER', 'TIME', 'COMMANDS', 'COMMAND RESPONSE', 'COMMAND PARAMETER']

        self.writer_Record_File = pd.ExcelWriter(self.Record_File, engine='openpyxl')
        self.writer_set_Record_File = pd.ExcelWriter(self.set_Record_File, engine='openpyxl')
        self.ind = 0
        self.list_sms_para = []
        self.list_sms_rsp = []

        # default number of the device
        self.number.setPlainText(list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['NUMBER'])
                                      .dropna(how="any")['NUMBER'])[0])

        # Main Database columns
        self.file_s_no = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['S No.'])
                              .dropna(how="any")['S No.'])
        self.file_commands = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['COMMAND'])
                                  .dropna(how="any")['COMMAND'])
        self.file_timing = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['TIME'])
                                .dropna(how="any")['TIME'])
        self.file_response_no = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['RESPONSE NUMBER'])
                                     .dropna(how="any")['RESPONSE NUMBER'])
        self.file_com_parameter = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['COMMAND PARAMETER'])
                                       .dropna(how="any")['COMMAND PARAMETER'])
        self.file_com_response = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['COMMAND RESPONSE'])
                                      .dropna(how="any")['COMMAND RESPONSE'])
        self.set_checkboxx = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['EXTRA CHECKBOX'])
                                  .dropna(how="any")['EXTRA CHECKBOX'])
        self.set_checkbox = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['CHECKBOX'])
                                 .dropna(how="any")['CHECKBOX'])
        self.set_textboxx = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['EXTRA TEXTBOX'])
                                 .dropna(how="any")['EXTRA TEXTBOX'])
        self.set_textbox = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['TEXTBOX'])
                                .dropna(how="any")['TEXTBOX'])
        self.file_set_dir_parameter = list(
            pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['SET DIR PARAMETER'])
                .dropna(how="any")['SET DIR PARAMETER'])
        self.file_set_dir_extra_parameter = list(
            pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['SET DIR EXTRA PARAMETER'])
                .dropna(how="any")['SET DIR EXTRA PARAMETER'])
        self.file_var_parameter = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['VARIABLE PARAMETER'])
                                       .dropna(how="any")['VARIABLE PARAMETER'])
        self.file_default_text = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['DEFAULT TEXT'])
                                      .dropna(how="any")['DEFAULT TEXT'])
        self.file_set_rx_list = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['SET RX LIST'])
                                     .dropna(how="any")['SET RX LIST'])
        self.file_set_rx_index = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['SET RX INDEX'])
                                      .dropna(how="any")['SET RX INDEX'])
        self.file_extra_com = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['EXTRA COMMAND'])
                                   .dropna(how="any")['EXTRA COMMAND'])
        self.file_default_extra_text = list(
            pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['DEFAULT EXTRA TEXT'])
                .dropna(how="any")['DEFAULT EXTRA TEXT'])
        self.file_set_rx_extra_list = list(
            pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['SET RX EXTRA LIST'])
                .dropna(how="any")['SET RX EXTRA LIST'])
        self.file_set_rx_extra_index = list(
            pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['SET RX EXTRA INDEX'])
                .dropna(how="any")['SET RX EXTRA INDEX'])
        self.file_set_default_mcn = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['SET DEFAULT MCN'])
                                         .dropna(how="any")['SET DEFAULT MCN'])
        self.file_default_values = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['DEFAULT VALUES'])
                                        .dropna(how="any")['DEFAULT VALUES'])

        self.set_response_parameters = self.file_set_dir_parameter[0:1] + self.file_set_dir_extra_parameter[0:1] + \
                                       self.file_set_dir_extra_parameter[2:3] + self.file_set_dir_parameter[1:]

        # auto send values initialisation
        self.send_list = []
        self.check_time.setPlainText(self.file_default_values[0])

        self.Index_from.setPlainText(self.file_default_values[1])
        self.Index_to.setPlainText(self.file_default_values[2])

        self.Wait_time_to_rx.setPlainText(self.file_default_values[3])
        self.No_of_repetitions.setPlainText(self.file_default_values[4])
        self.Delay_in_repetition.setPlainText(self.file_default_values[5])

        self.set_No_of_repetitions.setPlainText(self.file_default_values[6])
        self.set_Delay_in_repetition.setPlainText(self.file_default_values[7])
        self.Set_send_time.setPlainText(self.file_default_values[8])
        self.Set_receive_time.setPlainText(self.file_default_values[9])

        self.GAIA_Project.setPlainText(self.file_default_values[10])

        # filling the commands combo box
        for i in self.file_commands:
            self.message_le.addItem(i)

        # set values initialisation
        self.lis = []
        self.default_texts = self.file_default_text[0:1] + self.file_default_extra_text[:] + self.file_default_text[1:]
        self.checkboxes = self.set_checkbox[0:1] + self.set_checkboxx[:] + self.set_checkbox[1:]
        self.labels = self.file_set_dir_parameter[0:1] + self.file_set_dir_extra_parameter[:] + \
                      self.file_set_dir_parameter[1:]
        self.file_set_dir_parameter = self.labels
        self.textboxes = self.set_textbox[0:1] + self.set_textboxx[:] + self.set_textbox[1:]

        z = zip(self.checkboxes, self.labels, self.textboxes, self.default_texts)

        self.checkbox_list = []
        self.set_command_value = []
        self.set_command_val = []
        i = 0

        for c, l, t, d in z:
            # print(i, c, l, t, d)
            self.Frame = QFrame(self.Update)
            self.Frame.setMinimumSize(QSize(0, 60))
            self.Frame.setFrameShape(QFrame.StyledPanel)
            self.Frame.setFrameShadow(QFrame.Raised)
            self.Frame.setObjectName("Frame" + str(i))

            self.GridLayout = QGridLayout(self.Frame)
            self.GridLayout.setContentsMargins(2, -1, 2, -1)
            self.GridLayout.setObjectName("GridLayout" + str(i))

            self.panel = QHBoxLayout()
            self.panel.setObjectName("panel" + str(i))

            if c == "_":
                self.checkbox = QFrame(self.Frame)
                self.checkbox.setObjectName(c)
            else:
                self.checkbox = QCheckBox(self.Frame)
                self.checkbox.setMaximumSize(QSize(17, 31))
                self.checkbox.setTabletTracking(True)
                self.checkbox.setChecked(False)
                self.checkbox.setObjectName(c)
                self.checkbox_list.append(self.checkbox)
            self.panel.addWidget(self.checkbox)

            self.label = QLabel(self.Frame)
            self.label.setFocusPolicy(Qt.ClickFocus)
            self.label.setObjectName(l)
            self.label.setText(
                "<html><head/><body><p><span style=\" font-size:10pt;\">" + l + "</span></p></body></html>")
            self.panel.addWidget(self.label)

            self.text = QPlainTextEdit(self.Frame)
            sizePolicy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            sizePolicy.setHorizontalStretch(0)
            sizePolicy.setVerticalStretch(0)
            sizePolicy.setHeightForWidth(self.text.sizePolicy().hasHeightForWidth())
            self.text.setSizePolicy(sizePolicy)
            self.text.setMinimumSize(QSize(100, 0))
            self.text.setMaximumSize(QSize(320, 50))
            self.text.setFocusPolicy(Qt.ClickFocus)
            self.text.setStyleSheet("color: rgb(91, 91, 91) ;\n""font: 11pt \"MS Shell Dlg 2\";")
            self.text.setFrameShape(QFrame.WinPanel)
            self.text.setFrameShadow(QFrame.Sunken)
            self.text.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.text.setOverwriteMode(False)
            self.text.setObjectName(t)
            self.text.setPlainText(d)
            self.set_command_val.append(self.text)
            self.panel.addWidget(self.text)

            self.GridLayout.addLayout(self.panel, 0, 0, 1, 1)
            self.verticalLayout_11.addWidget(self.Frame)
            i += 1

        self.Updating.setWidget(self.Update)

        self.set_rx_resp_no = [0] * (len(self.file_set_rx_list + self.file_set_rx_extra_list))
        self.set_rx_resp_ind = [0] * (len(self.file_set_rx_list + self.file_set_rx_extra_list))
        self.set_list = []
        self.set_rx_list = self.file_set_rx_list[0:1] + self.file_set_rx_extra_list + self.file_set_rx_list[1:]
        self.set_rx_index = self.file_set_rx_index[0:1] + self.file_set_rx_extra_index[0:2] + self.file_set_rx_index[1:]
        for j in range(len(self.set_rx_list)):
            for i in range(len(self.file_commands)):
                if str(self.set_rx_list[j]) == str(self.file_commands[i]):
                    self.set_rx_resp_no[j] = int(self.file_response_no[i])
                    self.set_rx_resp_ind[j] = i

        self.checked_list = []
        self.set_not_rx_list = []
        for i in range(len(self.set_rx_list)):
            if self.set_rx_list[i] == "_":
                self.set_not_rx_list.append(i)

        self.set_rx_check_list = []
        for index, set_rx_list_element in enumerate(self.set_rx_list):
            x = []
            for index1, set_rx_list_element_element in enumerate(self.set_rx_list[:index]):
                if set_rx_list_element == set_rx_list_element_element:
                    x.append(index1)
            self.set_rx_check_list.append(x)

        # clicked buttons initialisation
        self.button.clicked.connect(self.on_toggled)
        self.info_btn.clicked.connect(self.info)
        self.Send_btn.clicked.connect(self.send_start)
        self.Receive_btn.clicked.connect(self.receive_start)
        self.clear_screen_btn.clicked.connect(self.clear_screen)
        self.quit_btn.clicked.connect(QApplication.quit)

        self.check_btn.clicked.connect(self.check_start)

        self.Get_response_data_btn.clicked.connect(self.get_response_data)
        self.Set_response_data_btn.clicked.connect(self.set_response_data)

        self.Delete_btn.clicked.connect(self.delete_all_sms)
        self.Show_btn.clicked.connect(self.show_all_sms)

        self.Empty_btn.clicked.connect(self.Empty_excel)
        self.Empty_set_btn.clicked.connect(self.set_Empty_excel)

        self.auto_send_btn.clicked.connect(self.sequence_send_start)
        self.auto_receive_btn.clicked.connect(self.sequence_send_custom_receive)
        self.repeat_auto_send_btn.clicked.connect(self.repetitive_sequence_send_start)
        self.set_get_response_btn.clicked.connect(self.set_get_response_start)

        self.set_btn.clicked.connect(self.set)
        self.set_send_rx_btn.clicked.connect(self.set_send_rx_start)
        self.repeat_auto_set_btn.clicked.connect(self.repetitive_sequence_set_start)
        self.set_dvc_id_btn.clicked.connect(self.set_dvc_id)
        self.select_all_btn.clicked.connect(self.select_all)

        # receive values initialisation
        self.data = ''
        self.receive_list = []
        self.response_number = 0
        self.sms_needed_list = []
        self.all_sms = ''
        self.all_sms_list = []

        # timers threads initialisation
        self.timer = QTimer()

        self.timer1 = QTimer()
        self.timer1.setInterval(3000)
        self.timer1.timeout.connect(self.set_commands_sending)

        self.timer2 = QTimer()
        self.timer2.setInterval(3000)
        self.timer2.timeout.connect(self.send_multiple)

        self.timer3 = QTimer()
        self.timer3.setInterval(3000)
        self.timer3.timeout.connect(self.receive_only)

        self.timer4 = QTimer()
        self.timer4.setInterval(3000)
        self.timer4.timeout.connect(self.sequence_send_multiple)

        self.timer5 = QTimer()
        self.timer5.setInterval(3000)
        self.timer5.timeout.connect(self.repetitions)

        self.timer6 = QTimer()
        self.timer6.setInterval(3000)
        self.timer6.timeout.connect(self.set_send_rx_multiple)

        self.timer7 = QTimer()
        self.timer7.setInterval(3000)
        self.timer7.timeout.connect(self.set_repetitions)

        self.timer8 = QTimer()
        self.timer8.setInterval(3000)
        self.timer8.timeout.connect(self.set_get_response)

        self.timer9 = QTimer()
        self.timer9.setInterval(3000)
        self.timer9.timeout.connect(self.check)

        self.recipient = None
        self.content = None

        # port initialisation
        self.comPort = None
        self.on_toggled()
        self.serial = serial.Serial(self.comPort,
                                    9600,
                                    timeout=5,
                                    xonxoff=False,
                                    rtscts=False,
                                    bytesize=serial.EIGHTBITS,
                                    parity=serial.PARITY_NONE,
                                    stopbits=serial.STOPBITS_ONE)

        # Check Closing Excels
        self.load_get()
        self.load_set()

    # ----------------------------------COMMON-------------------------------------------------------------------------#
    @pyqtSlot()
    def info(self):
        self.output_te.append("")
        self.output_te.append("1...<b>Instruction</b>: <I>'Info'</I>")
        self.output_te.append("2...<b>Connection error or not</b>: <i>'Disconnected'/'Connected'</i>")
        self.output_te.append("3...<b>Single GET command</b>: <i>'Send' or 'Receive'</i>")
        self.output_te.append("4...<b>Quick single GET command</b>: <i>'Check'</i>")
        self.output_te.append("5...<b>Multiple GET commands</b>: <i>'Auto Send' or 'Auto Receive'</i>")
        self.output_te.append("6...<b>Repetitive multiple GET commands</b>: <i>'Repeat Sent-Receive'</i>")
        self.output_te.append("7...<b>Only set device id of multiple devices</b>: <i>'SET Device Id'</i>")
        self.output_te.append("8...<b>Toggle selection on SET command</b>: <i>'Select All'</i>")
        self.output_te.append("9...<b>Only SET commands</b>: <i>'SET'</i>")
        self.output_te.append("10..<b>SET commands with validation</b>: <i> 'SET Check' </i>")
        self.output_te.append("11..<b>Repetitive multiple SET commands with validation</b>: <i>'Repeat SET Check'</i>")
        self.output_te.append("12..<b>Runs whole GUI (SET and then GET)</b>: <i>'SET - GET Response'</i>")
        self.output_te.append("13..<b>Data analysis of GET or SET Records</b>: <i>'Analyse GET Data' or 'Analyse SET "
                              "Data'</i>")
        self.output_te.append("14..<b>Renew GET or SET Records</b>: <i>'Empty GET Data' or 'Empty SET Data'</i>")
        self.output_te.append("15..<b>Show or Delete all msgs of GSM module</b>: <i>'Show' or 'Delete'</i>")
        self.output_te.append("16..<b>Clear Command Response screen</b>: <i>'Clear Screen'</i>")
        self.output_te.append("17..<b>Close App</b>: <i>'Quit'</i>")
        self.output_te.append("\n")

    @pyqtSlot(bool)
    def on_toggled(self):
        # port initialisation
        for i in port.comports():
            if 'USB-SERIAL CH340' in str(i):
                self.comPort = str(i).split(' ')[0]

        if not self.comPort:
            self.button.setText("Disconnected")
            print('Error in hardware connection')
            self.output_te.append('\n<b>Error in hardware connection...')
        else:
            self.button.setText("Connected")
            self.output_te.append('\n' + '<b>Connected...')

    @pyqtSlot()
    def clear_screen(self):
        self.output_te.clear()

    @pyqtSlot()
    def delete_all_sms(self):
        try:
            self.output_te.append("<b>\nAll SMS in inbox deleting...")
            self.send_list = ["AT\r", 'AT+CMGD=1,4\r']
            print("VALUE TAKEN")
            self.sendMessage()
            time.sleep(10)
            print("Msgs Deleted !\n")
            self.output_te.append("<b>Deleted")
        except Exception as e:
            print('Error in delete_all_sms', e)
            self.output_te.append('<b>\nError in delete_all_sms: ' + str(e))

    @pyqtSlot()
    def show_all_sms(self):
        self.output_te.append("<b>All SMS in inbox collecting...")
        self.send_list = ['AT+CMGL="ALL"\r']
        print("VALUE TAKEN")
        try:
            for i in self.send_list:
                self.serial.write(i.encode())
                time.sleep(1)
                self.output_te.append(i)
                x = self.serial.readall()
                print(x)
                for j in str(x.decode('utf-8')).split('\n'):
                    self.output_te.append(j)
        except Exception as e:
            print('Error in show_all_msgs', e)
            self.output_te.append('<b>\nError in show_all_msgs: ' + str(e))
        print("MESSAGE SENT")
        self.disconnectPhone()
        self.output_te.append("<b>Collected all messages :)")

    @pyqtSlot()
    def check_start(self):
        self.output_te.append("\n<b>Checking...")
        self.timer9.start()

    @pyqtSlot()
    def check(self):
        if 'Choose Command' in self.message_le.currentText():
            self.output_te.append("<b>Un-select CHOOSE COMMAND :(")
        else:
            numbers = self.number.toPlainText().split(',')
            print("\nNumbers:", numbers)
            for i in numbers:
                try:
                    i = i.strip()
                    print("\nNumber:", i)
                    self.delete_all_sms()
                    self.send(i, self.message_le.currentText())
                    print("\nReceive in:", self.check_time.toPlainText())
                    time.sleep(int(self.check_time.toPlainText()))
                    self.GetAllSMS(i)
                    if not numbers[-1] == i:
                        print("\nNext number in: 5")
                        time.sleep(5)
                except Exception as e:
                    print('Error in check', e)
                    self.output_te.append('<b>\nError in check: ' + str(e) + " at number " + str(i))
        self.output_te.append("\n<b>Checking completed :)")
        self.timer9.stop()

    @pyqtSlot()
    def Empty_excel(self):
        self.writer_Record_File = pd.ExcelWriter(self.Record_File, engine='openpyxl')
        try:
            data_frame = [''] * len(self.file_response_no)
            sms_parameter = [[]] * len(self.file_response_no)
            for i in range(len(self.file_response_no)):
                if int(self.file_response_no[i]) > 0:
                    sms_parameter[i] = self.file_com_parameter[i].split(',')

                    common_data_frame = pd.DataFrame(columns=self.common_dir_parameter)  # blank common data frame
                    common_data_frame.to_excel(self.writer_Record_File, sheet_name=self.Record_File_Common_Sheet)

                    data_frame[i] = pd.DataFrame(columns=['NUMBER', 'TIME'] + sms_parameter[i])  # blank data frame
                    data_frame[i].to_excel(self.writer_Record_File, sheet_name=self.file_commands[i])

            self.writer_Record_File.save()

            print("Update Done")
            self.output_te.append("\n<b>GET Excel Empty Update Successful :)")
        except Exception as e:
            print('Error in Empty_excel', e)
            self.output_te.append("\n" + str(e) + "<b>...Close the Get File to proceed :(")

    @pyqtSlot()
    def set_Empty_excel(self):
        self.writer_set_Record_File = pd.ExcelWriter(self.set_Record_File, engine='openpyxl')
        try:
            file_set_dir_parameter = ['NUMBER', 'TIME'] + self.set_response_parameters
            data_frame = pd.DataFrame(columns=file_set_dir_parameter)
            data_frame.to_excel(self.writer_set_Record_File, columns=file_set_dir_parameter)

            self.writer_set_Record_File.save()
            print("Update Done")
            self.output_te.append("\n<b>SET Excel Empty update successful :)")
        except Exception as e:
            print('Error in set_Empty_excel', e)
            self.output_te.append("\n" + str(e) + "<b>...Close the Set File to proceed :(")

    @pyqtSlot()
    def select_all(self):
        try:
            for i in self.checkbox_list:
                if i.isChecked():
                    i.setCheckState(False)
                elif not i.isChecked():
                    i.setCheckState(True)
            self.output_te.append("<b>All the SET options toggled :)</b>")
        except Exception as e:
            print('Error in send', e)
            self.output_te.append('<b>\nError in select_all: ' + str(e))

    # --------------------------------END COMMON-----------------------------------------------------------------------#

    # --------------------------------SEND MESSAGE---------------------------------------------------------------------#

    @pyqtSlot()
    def send_start(self):
        self.output_te.append("\n<b>Sending...")
        self.timer2.start()

    @pyqtSlot()
    def send_multiple(self):
        if 'Choose Command' in self.message_le.currentText():
            print("<b>Un-select CHOOSE COMMAND :(")
            self.output_te.append("<b>Un-select CHOOSE COMMAND :(")
            self.timer2.stop()
        else:
            numbers = self.number.toPlainText().split(',')
            print("\nNumbers:", numbers)
            for i in numbers:
                try:
                    i = i.strip()
                    print("\nNumber:", i)
                    self.delete_all_sms()
                    self.send(i, self.message_le.currentText())
                    if not numbers[-1] == i:
                        print("\nNext number in: 5")
                        time.sleep(5)
                except Exception as e:
                    print('Error in sendMessage', e)
                    self.output_te.append('<b>\nError in sendMessage: ' + str(e) + " at number " + str(i))
            self.timer2.stop()

    @pyqtSlot(str, str)
    def send(self, number, set_content):
        try:
            self.send_list = ["AT\r", 'AT+CMGF=1\r', '''AT+CMGS="''' + number + '''"\r''',
                              set_content, chr(26)]
            print("VALUE TAKEN")
            self.sendMessage()

        except Exception as e:
            print('Error in send', e)
            self.output_te.append('<b>\nError in send: ' + str(e) + " at number " + number)

    @pyqtSlot()
    def sendMessage(self):
        try:
            for i in self.send_list:
                self.serial.write(i.encode())
                time.sleep(1.5)
                if i in self.send_list[:-1]:
                    print(self.serial.readline())
                else:
                    self.all_sms = self.serial.readall()
            print("MESSAGE SENT")
            self.disconnectPhone()
        except Exception as e:
            print('Error in sendMessage:', str(e))
            self.output_te.append('<b>\nError in sendMessage: ' + str(e))

    @pyqtSlot()
    def disconnectPhone(self):
        self.all_sms = str(self.all_sms.decode('utf-8'))
        print("\nResponse:", self.all_sms.strip())
        self.output_te.append(self.all_sms)
        if self.all_sms.find("+CMGS: ") > -1 or self.all_sms.find("ERROR") == -1:
            print("\nSent successful :) ")
            self.output_te.append("<b>\nSent successful :)\n")
        else:
            print("\nSent unsuccessful :( ")
            self.output_te.append("<b>\nSent unsuccessful :(\n")
        print("DISCONNECT")

    # ----------------------------------END SEND MESSAGE---------------------------------------------------------------#

    # -----------------------------------RECEIVE MESSAGE---------------------------------------------------------------#
    @pyqtSlot()
    def receive_start(self):
        self.output_te.append("\n<b>Receiving...")
        self.timer3.start()

    @pyqtSlot()
    def receive_only(self):
        try:
            self.GetAllSMS(self.number.toPlainText().split(",")[-1])
            self.timer3.stop()
        except Exception as e:
            print('Error in receive_only', str(e))
            self.output_te.append('<b>\nError in receive_only: ' + str(e))

    @pyqtSlot()
    def GetAllSMS(self, number):
        try:
            for i in range(len(self.file_commands)):
                if str(self.message_le.currentText()) == str(self.file_commands[i]):
                    self.response_number = int(self.file_response_no[i])
                    self.ind = i
                    print("\nRECEIVE:", self.message_le.currentText())
                    break
            self.receiving(number)
            self.timer3.stop()
        except Exception as e:
            print('Error in GetAllSMS:', str(e))
            self.output_te.append('<b>\nError in GetAllSMS: ' + str(e))

    @pyqtSlot(str)
    def receiving(self, number):
        try:
            self.receive()
            self.receive_1(number)
        except Exception as e:
            print('Error in receiving:', str(e))
            self.output_te.append('<b>\nError in receiving: ' + str(e))

    @pyqtSlot()
    def receive(self):
        try:
            self.receive_list = ['AT+CMGF=1\r', "AT\r", '''AT+CMGL="ALL"\r''']
            print("Searching...\n")
            self.sms_needed_list = []
            for i in self.receive_list:
                self.serial.write(i.encode())
                time.sleep(1)
                if i in self.receive_list[:-1]:
                    print(self.serial.readline())
                else:
                    self.all_sms = self.serial.readall()
                    print(self.all_sms)
                self.output_te.append(i)
            self.all_sms_list = self.all_sms.decode('utf-8').split('\n')
        except Exception as e:
            print('Error in receive:', str(e))
            self.output_te.append('<b>Error in receive: ' + str(e))

    @pyqtSlot(str)
    def receive_1(self, number):
        try:
            print("\nReceiving:", self.file_commands[self.ind], self.response_number)
            for index in range(len(self.all_sms_list)):
                if ('GAIA ' + str(self.GAIA_Project.toPlainText()) + ',' + str(
                        self.file_response_no[self.ind]) + ',') in self.all_sms_list[index]:
                    self.sms_needed_list = self.all_sms_list[
                                           index: int(index + len(list(self.file_com_parameter[self.ind].split(','))))]
            if len(self.sms_needed_list) > 1:
                self.list_sms_para = list(self.file_com_parameter[self.ind].split(','))
                self.list_sms_rsp = ''.join(self.sms_needed_list).split(',')

                for sms_parameter, sms_response in zip(self.list_sms_para, self.list_sms_rsp):
                    print(f'{sms_parameter.strip()} : {sms_response.strip()}')
                    self.output_te.append(f'<b>{sms_parameter.strip()}</b> : <i>{sms_response.strip()}</i>')

                print(self.list_sms_para, "\n", self.list_sms_rsp)
                self.Save_excel(self.ind, number, self.response_number, self.list_sms_para, self.list_sms_rsp)

                self.output_te.append("<b>Response successful :)\n")
                print("Response successful :)\n")
            else:
                self.output_te.append("<b>Response unsuccessful :(\n")
                print("Response unsuccessful :(\n")
            print("Done Receive")
            self.disconnect_Phone()
        except Exception as e:
            print('Error in receive_1:', str(e))
            self.output_te.append('<b>\nError in receive_1: ' + str(e))

    @pyqtSlot(int, str, int, str, str)
    def Save_excel(self, ind, number, response_no, sms_parameter, sms_response):
        print("\nSaving in excel...")

        response_no = int(response_no)
        if int(response_no) > 0:
            print(response_no, "Response checked: ", ind)

            previous_record = []
            try:
                current_time = datetime.datetime.now()

                common_previous_record = pd.read_excel(self.Record_File, sheet_name=self.Record_File_Common_Sheet,
                                                       index_col=[0]).dropna(how="all")
                common_count_new = len(list(common_previous_record[self.common_dir_parameter[0]]))
                common_data_frame = pd.DataFrame([[str(number), str(current_time), str(self.file_commands[ind]),
                                                   str(sms_response), str(sms_parameter)]],
                                                 columns=self.common_dir_parameter, index=[common_count_new])
                common_new_record = common_previous_record.append(common_data_frame, verify_integrity=True, sort=False)
                common_new_record.to_excel(self.writer_Record_File, sheet_name=self.Record_File_Common_Sheet,
                                           columns=self.common_dir_parameter)

                for i in range(len(self.file_response_no)):
                    previous_record.append(' ')
                    if int(self.file_response_no[i]) > 0:
                        previous_record[i] = pd.read_excel(self.Record_File, sheet_name=self.file_commands[i],
                                                           index_col=[0]).dropna(how="all")
                        previous_record[i].to_excel(self.writer_Record_File, sheet_name=self.file_commands[i])
                counter_new = len(list(previous_record[ind][sms_parameter[0]]))

                print(ind, "counter_new and common_count_new", counter_new, common_count_new, "record done")
                print(ind, "Common record generated: ", common_new_record)

                if len(self.list_sms_rsp) == len(self.list_sms_para):
                    current_time = datetime.datetime.now()
                    data_frame = pd.DataFrame([[number, current_time] + sms_response],
                                              columns=['NUMBER', 'TIME'] + sms_parameter, index=[counter_new])
                    new_record = previous_record[ind].append(data_frame, verify_integrity=True, sort=False)
                    new_record.to_excel(self.writer_Record_File, sheet_name=self.file_commands[ind],
                                        columns=['NUMBER', 'TIME'] + sms_parameter)
                    print(ind, "Specific record generated: ", data_frame, new_record)
                else:
                    self.output_te.append("Incomplete message received :(")
            except Exception as e:
                print("Exception save_excel:", str(e))
                self.output_te.append("<b>Exception save_excel: " + str(e))

            self.writer_Record_File.save()
            print("Update Done")
            self.output_te.append("\n<b>Excel update successful :)")

    @pyqtSlot()
    def disconnect_Phone(self):
        self.timer3.stop()

    # ---------------------------------END RECEIVE MESSAGE-------------------------------------------------------------#

    # --------------------------------AUTO SEND MESSAGE----------------------------------------------------------------#
    @pyqtSlot()
    def sequence_send_start(self):
        self.output_te.append("\n<b>Sequential sending started...")
        self.timer4.start()

    @pyqtSlot()
    def sequence_send_multiple(self):
        numbers = self.number.toPlainText().split(',')
        print("\nNumbers:", numbers)
        for i in numbers:
            try:
                i = i.strip()
                print("\nNumber:", i)
                self.delete_all_sms()
                self.sequence_send(i)
                if not numbers[-1] == i:
                    print("\nNext number in: 5")
                    time.sleep(5)
            except Exception as e:
                print("Exception sequence_send_multiple:", str(e))
                self.output_te.append("<b>\nException sequence_send_multiple: " + str(e) + " at number " + str(i + 1))

    @pyqtSlot(str)
    def sequence_send(self, number):
        try:
            index_from = int(self.Index_from.toPlainText())
            index_to = int(self.Index_to.toPlainText())
            response_no_list = []
            for i in range(index_from, index_to + 1):
                print("Sending: ", self.file_commands[i])

                self.response_number = int(self.file_response_no[i])
                print("Response No:", self.response_number)

                self.send(number, self.file_commands[i])

                if int(self.response_number) > 0:
                    response_no_list.append(i)
                print("Updated response list: ", response_no_list)

                print("\nNext msg in: ", int(self.file_timing[i]))
                time.sleep(int(self.file_timing[i]))

            if response_no_list:
                time.sleep(int(self.Wait_time_to_rx.toPlainText()))
                print("\nReceive in:", int(self.Wait_time_to_rx.toPlainText()))

            self.sequence_send_receive(response_no_list, number)
        except Exception as e:
            print("Exception sequence_send:", str(e))
            self.output_te.append("<b>\nException sequence_send: " + str(e))

        self.timer4.stop()

    @pyqtSlot(str, str)
    def sequence_send_receive(self, response_no_list, number):
        try:
            for i in response_no_list:
                self.ind = i
                self.auto_GetAllSMS(i, number)

            self.output_te.append("<b>All the responses received :)\n")
        except Exception as e:
            print("Exception sequence_send_receive:", str(e))
            self.output_te.append("<b>\nException sequence_send_receive: " + str(e))

    @pyqtSlot()
    def sequence_send_custom_receive(self):
        try:
            number = self.number.toPlainText()
            index_from = int(self.Index_from.toPlainText())
            index_to = int(self.Index_to.toPlainText())
            response_no_list = []
            for i in range(index_from, index_to + 1):
                response_no_list.append(i)
                self.ind = i
                self.auto_GetAllSMS(i, number)

            self.output_te.append("<b>All the responses received :)\n")
        except Exception as e:
            print("Exception sequence_send_custom_receive:", str(e))
            self.output_te.append("<b>\nException sequence_send_custom_receive: " + str(e))

    @pyqtSlot(int)
    def auto_GetAllSMS(self, response_no, number):
        try:
            self.response_number = response_no
            self.receiving(number)
        except Exception as e:
            print("Exception auto_GetAllSMS:", str(e))
            self.output_te.append("<b>Exception auto_GetAllSMS: " + str(e))

    @pyqtSlot()
    def repetitive_sequence_send_start(self):
        self.output_te.append("\n<b>Repetitive sequential sending started...")
        self.timer5.start()

    @pyqtSlot()
    def repetitions(self):
        try:
            for i in range(int(self.No_of_repetitions.toPlainText())):
                self.sequence_send_multiple()
                time.sleep(int(self.Delay_in_repetition.toPlainText()))

            self.output_te.append("<b>Repetitive sequential sending complete :)")
        except Exception as e:
            print("Exception repetitions:", str(e))
            self.output_te.append("<b>\nException repetitions: " + str(e) + " at repetition " + str(i + 1))
        self.timer5.stop()

    # --------------------------------END AUTO SEND MESSAGE------------------------------------------------------------#

    # --------------------------------SET CONFIGURATION----------------------------------------------------------------#
    @pyqtSlot()
    def set(self):
        self.output_te.append("\n<b>Setting Configurations...")
        self.timer1.start()

    @pyqtSlot()
    def set_commands_sending(self):
        try:
            self.set_commands_send()
            self.set_multiple()
            self.output_te.append("<b>Set configuration successful :)\n")
        except Exception as e:
            print("Exception sequence_send_custom_receive:", e)
            self.output_te.append("<b>\nException sequence_send_custom_receive: " + str(e))
        self.timer1.stop()

    @pyqtSlot()
    def set_commands_send(self):
        try:
            self.set_command_value = []
            for i in self.set_command_val[0:1] + self.set_command_val[6:]:
                self.set_command_value.append(i.toPlainText().strip())

            self.set_list = []
            for set_command, set_command_value in zip(self.file_var_parameter[0:1], self.set_command_value[0:1]):
                if set_command_value != '':
                    self.set_list.append(f'{set_command[ : set_command.index("[")]}{set_command_value}')

            if self.set_command_val[1].toPlainText() != '' and self.set_command_val[2].toPlainText() != '':
                self.set_list.append(
                    f'DEFAULT_MCN_{self.set_command_val[1].toPlainText().strip()},{self.set_command_val[2].toPlainText().strip()}')

            if (self.set_command_val[3].toPlainText() != '' and self.set_command_val[4].toPlainText() != ''
                    and self.set_command_val[5].toPlainText() != ''):
                self.set_list.append(f'SET_DVC_MCN_{self.set_command_val[3].toPlainText().strip()},'
                                     f'{self.set_command_val[4].toPlainText().strip()},{self.set_command_val[5].toPlainText().strip()}')

            for set_command, set_command_value in zip(self.file_var_parameter[1:], self.set_command_value[1:]):
                if set_command_value != '':
                    self.set_list.append(f'{set_command[ : set_command.index("[")]}{set_command_value}')

            self.lis = []
            print("Following commands set:")
            for i in range(len(self.set_list)):
                if self.checkbox_list[i].isChecked():
                    self.lis.append(self.set_list[i])
                    print(f'{i+1}. {self.set_list[i]}')
                    self.checked_list.append(i)
        except Exception as e:
            print("Exception set_commands_send:", e)
            self.output_te.append("<b>\nException set_commands_send: " + str(e))

    @pyqtSlot()
    def set_multiple(self):
        numbers = self.number.toPlainText().split(',')
        print("\nNumbers:",numbers)

        for i in numbers:
            try:
                i = i.strip()
                print("Number: ", i)
                if len(self.lis) == 1:
                    for j in self.lis:
                        print(j, "\nExecuted\n")
                        self.send(i, j)
                    if not numbers[-1] == i:
                        print("\nNext number in: 20")
                        time.sleep(20)
                else:
                    for j in self.lis:
                        print(j, "\nExecuted\n")
                        self.send(i, j)
                        print("\nNext set_send in: ", self.Set_send_time.toPlainText())
                        time.sleep(int(self.Set_send_time.toPlainText()))
                    if not numbers[-1] == i:
                        print("\nNext number in: 15")
                        time.sleep(15)
            except Exception as e:
                print("Exception set_multiple:", e)
                self.output_te.append("<b>\nException set_multiple: " + str(e) + " at number " + str(i + 1))
        self.output_te.append("Following commands set:")
        for k in range(len(self.lis)):
            self.output_te.append(f'{k+1}. {self.lis[k]}')

    # --------------------------------END SET CONF---------------------------------------------------------------------#

    # ----------------------------------SET & RECEIVE MESSAGE----------------------------------------------------------#
    @pyqtSlot()
    def set_send_rx_start(self):
        self.output_te.append("\n<b>Setting Configurations...")
        self.timer6.start()

    @pyqtSlot()
    def repetitive_sequence_set_start(self):
        self.output_te.append("\n<b>Repetitive sequential sending started...")
        self.timer7.start()

    @pyqtSlot()
    def set_repetitions(self):
        for i in range(int(self.set_No_of_repetitions.toPlainText())):
            try:
                self.set_send_rx_multiple()
                time.sleep(int(self.set_Delay_in_repetition.toPlainText()))
            except Exception as e:
                print("Exception set_repetitions:", e)
                self.output_te.append("<b>\nException set_repetitions: " + str(e) + " at repetition " + str(i + 1))
        self.output_te.append("<b>Repetitive sequential setting complete :)")
        self.timer7.stop()

    @pyqtSlot()
    def set_send_rx_multiple(self):
        numbers = self.number.toPlainText().split(',')
        print("\nNumbers:", numbers)
        for i in numbers:
            try:
                i = i.strip()
                print("\nNumber:", i)
                self.delete_all_sms()
                self.set_commands_send()
                self.set_single(i)
                if self.checked_list != self.set_not_rx_list:
                    self.set_sending(i)
                    self.set_receiving(i)
                if not numbers[-1] == i:
                    print("\nNext number in: 5")
                    time.sleep(5)
            except Exception as e:
                print("Exception set_send_rx_multiple:", e)
                self.output_te.append("<b>\nException set_send_rx_multiple: " + str(e) + " at number " + str(i))

        self.output_te.append("<b>Set configuration & rx successful :)\n")
        self.timer6.stop()

    @pyqtSlot(str)
    def set_sending(self, number):
        try:
            print(self.set_rx_list)
            for i in range(len(self.set_rx_list)):
                if self.checkbox_list[i].isChecked() and (i not in self.set_not_rx_list):
                    x = 0
                    if not self.set_rx_check_list[i]:
                        print("\nNext send in: ", self.Set_send_time.toPlainText())
                        time.sleep(int(self .Set_send_time.toPlainText()))
                        self.send(number, self.set_rx_list[i])
                    else:
                        for j in range(len(self.set_rx_check_list[i])):
                            if self.checkbox_list[int(self.set_rx_check_list[i][j])].isChecked():
                                x = 1
                                break
                        if x == 0:
                            print("\nNext send in: ", self.Set_send_time.toPlainText())
                            time.sleep(int(self.Set_send_time.toPlainText()))
                            self.send(number, self.set_rx_list[i])

            print("\nReceive in: ", self.Set_receive_time.toPlainText())
            time.sleep(int(self.Set_receive_time.toPlainText()))
        except Exception as e:
            print("Exception set_sending:", e)
            self.output_te.append("<b>Exception set_sending: " + str(e))

    @pyqtSlot(str)
    def set_receiving(self, number):
        set_check = [''] * len(self.set_rx_list)

        print("Set command val and set_check:", len(self.set_command_val), len(set_check))
        index=0
        try:
            for i in range(len(self.set_rx_list)):
                if self.checkbox_list[i].isChecked() and (i not in self.set_not_rx_list):  # s no. not having  checks & 7th index for passcode
                    self.list_sms_rsp = []
                    if i == 1:
                        if self.checkbox_list[i].isChecked():  # set default mcn - 1
                            set_check[i] = self.set_rx_mcn(1, self.set_command_val[1].toPlainText().strip(),
                                                           self.set_command_val[2].toPlainText().strip())
                    elif i == 2:
                        if self.checkbox_list[i].isChecked():  # set device mcn - 2
                            set_check[i] = self.set_rx_mcn(2, self.set_command_val[3].toPlainText().strip(),
                                                           self.set_command_val[4].toPlainText().strip())
                    else:
                        self.set_auto_GetAllSMS(self.set_rx_resp_ind[i], self.set_rx_resp_no[i])
                        if len(self.list_sms_rsp) < 1 or len(self.list_sms_rsp) <= int(self.set_rx_index[i]):
                            set_check[i] = 'NO RESPONSE' + str(self.list_sms_rsp)
                            self.output_te.append('<b>No response for: ' + str(self.set_rx_list[i]))
                        else:
                            if i == 0:
                                index = i
                            else:
                                index = i-2
                            print(len(self.list_sms_rsp), self.set_rx_index[i])
                            if str(self.list_sms_rsp[int(self.set_rx_index[i])]) == str(self.set_command_value[index]):
                                set_check[i] = 'SUCCESS\n' + str(self.list_sms_rsp)
                            else:
                                set_check[i] = 'FAIL\n' + str(self.list_sms_rsp)
                    time.sleep(1)
            print("Set Check: ", set_check)
            self.set_Save_excel(number, set_check)
        except Exception as e:
            print("Exception set_receiving:", e, i, set_check)
            self.output_te.append("<b>\nException set_receiving: " + str(e) + str(i) + str(set_check))

    @pyqtSlot(int, str, str)
    def set_rx_mcn(self, index, mcn_ind, mcn_num):
        try:
            y = ''
            self.set_auto_GetAllSMS(self.set_rx_resp_ind[index], self.set_rx_resp_no[index])

            x = [1, 2, 3]
            if len(self.list_sms_rsp) < 1:
                self.output_te.append('<b>No response for: ' + str(self.set_rx_list[index]))
                return 'NO RESPONSE []'
            else:
                for i, j in zip(x, self.file_set_rx_extra_index):
                    if str(mcn_ind) == str(i):
                        if int(index) == 1:
                            y = self.set_rx_default_mcn(i - 1, j)
                            break
                        elif int(index) == 2:
                            y = self.set_rx_setDvc_mcn(j, mcn_num)
                            break
                return y + '\n' + str(self.list_sms_rsp)
        except Exception as e:
            print("Exception set_rx_mcn:", e)
            self.output_te.append("<b>\nException set_rx_mcn: " + str(e))

    @pyqtSlot(int, str, str)
    def set_rx_default_mcn(self, file_ind, list_ind):
        print("True/not:", str(self.list_sms_rsp[int(list_ind)]) == str(self.file_set_default_mcn[file_ind].strip()))
        return 'SUCCESS' if str(self.list_sms_rsp[int(list_ind)]) == str(
            self.file_set_default_mcn[file_ind].strip()) else 'FAIL'

    @pyqtSlot(int, str, str)
    def set_rx_setDvc_mcn(self, list_ind, mcn_num):
        return 'SUCCESS' if str(self.list_sms_rsp[int(list_ind)]) == str(mcn_num) else 'FAIL'

    @pyqtSlot(int, int)
    def set_auto_GetAllSMS(self, ind, response_no):
        try:
            self.response_number = response_no
            self.receive()
            self.set_receive_1(ind, response_no)
        except Exception as e:
            print("Exception set_auto_GetAllSMS:", e)
            self.output_te.append("<b>\nException set_auto_GetAllSMS: " + str(e))

    @pyqtSlot(int, int)
    def set_receive_1(self, ind, response_no):
        print("\nReceiving:", ind, response_no)
        try:
            for index in range(len(self.all_sms_list)):
                if ('GAIA ' + str(self.GAIA_Project.toPlainText()).strip() + ',' + str(response_no) + ',') in \
                        self.all_sms_list[index]:
                    self.sms_needed_list = self.all_sms_list[
                                           index: int(index + len(list(self.file_com_parameter[ind].split(','))))]
            if len(self.sms_needed_list) > 1:
                self.list_sms_rsp = ''.join(self.sms_needed_list).split(',')
                print(self.list_sms_rsp)
        except Exception as e:
            print("Exception set_receive_1:", e)
            self.output_te.append("<b>\nException set_receive_1: " + str(e))

    @pyqtSlot(str)
    def set_single(self, number):
        try:
            for j in self.lis:
                print(j, "\nExecuted\n")
                self.send(number, j)

                if len(self.lis) > 1 and (not j == self.lis[-1]):
                    print("\nNext set_send in: ", self.Set_send_time.toPlainText())
                    time.sleep(int(self.Set_send_time.toPlainText()))

            self.output_te.append("\nFollowing commands set:")
            for k in range(len(self.lis)):
                self.output_te.append(f'{k+1}. {self.lis[k]}')
        except Exception as e:
            print("Exception set_single:", e)
            self.output_te.append("<b>\nException set_single: " + str(e))

    @pyqtSlot(str, str)
    def set_Save_excel(self, number, set_check):
        try:
            print("\nSaving in excel...")
            current_time = datetime.datetime.now()
            set_check = [number, current_time] + set_check
            print(set_check)
            file_set_dir_parameter = ['NUMBER', 'TIME'] + self.set_response_parameters

            previous_record = pd.read_excel(self.set_Record_File, index_col=[0]).dropna(how="all")
            count_new = len(list(previous_record[file_set_dir_parameter[0]]))
            data_frame = pd.DataFrame([set_check], index=[count_new], columns=file_set_dir_parameter)
            new_record = previous_record.append(data_frame, verify_integrity=True, sort=False)
            new_record.to_excel(self.writer_set_Record_File, columns=file_set_dir_parameter)
            self.writer_set_Record_File.save()

            print("Update Done")
            self.output_te.append("\n<b>Excel update successful :)")
        except Exception as e:
            print('Error in set_Save_excel', e)
            self.output_te.append("<b>\nClose the Set Check Database file to proceed: " + str(e))

    @pyqtSlot()
    def set_dvc_id(self):
        print("SET Dvc Id Started...")
        try:
            index_in_file = 6
            numbers = self.number.toPlainText().split(',')
            dvc_id = self.set_command_val[index_in_file].toPlainText().split(',')
            set_command = self.file_var_parameter[index_in_file-5]
            x = set_command[: set_command.index("[")]
            print("\nNumbers:", numbers)
            print("\nDvc Id:", dvc_id, "\n")
            if len(numbers) == len(dvc_id):
                for number, j in zip(numbers, dvc_id):
                    set_check = [''] * len(self.set_rx_list)
                    print("\nNumbers:", number)
                    self.delete_all_sms()

                    time.sleep(2)

                    print("\nSet_Sending:", x + j)
                    self.send(number, x + j)

                    print("\nSend in:", self.Set_send_time.toPlainText())
                    time.sleep(int(self.Set_send_time.toPlainText()))

                    i = 3
                    print("\nSending:", self.set_rx_list[i])
                    self.send(number, self.set_rx_list[i])

                    time.sleep(2)

                    self.set_auto_GetAllSMS(self.set_rx_resp_ind[i], self.set_rx_resp_no[i])
                    if len(self.list_sms_rsp) < 1 or len(self.list_sms_rsp) <= int(self.set_rx_index[i]):
                        set_check[i] = 'NO RESPONSE' + str(self.list_sms_rsp)
                        self.output_te.append('<b>No response for: ' + str(self.set_rx_list[i]))
                    else:
                        if str(self.list_sms_rsp[int(self.set_rx_index[i])]) == str(j):
                            set_check[i] = 'SUCCESS\n' + str(self.list_sms_rsp)
                        else:
                            set_check[i] = 'FAIL\n' + str(self.list_sms_rsp)
                    time.sleep(1)
                    print("\nSet Check:", set_check)
                    self.set_Save_excel(number, set_check)

                    if not numbers[-1] == number:
                        print("\nNext Receive in:", self.Set_send_time.toPlainText())
                        time.sleep(int(self.Set_send_time.toPlainText()))
            else:
                self.output_te.append("\nNumbers and Device Id not equal :(")

        except Exception as e:
            print("Exception set_dvc_id:", e)
            self.output_te.append("\nException set_dvc_id: " + str(e))

        self.output_te.append("<b>\nDevice ID set in all devices :)")

    # ------------------------------END SET & RECEIVE MESSAGE----------------------------------------------------------#
    # ------------------------------SET & GET RESPONSE-----------------------------------------------------------------#

    @pyqtSlot()
    def set_get_response_start(self):
        self.output_te.append("\n<b>Starting set get response...")
        self.timer8.start()

    @pyqtSlot()
    def set_get_response(self):
        try:
            self.set_repetitions()
            time.sleep(10)
            self.repetitions()
        except Exception as e:
            print("Exception set_get_response:", e)
            self.output_te.append("<b>\nException set_get_response: " + str(e))
        self.timer8.stop()

    # ---------------------------------END SET & GET RESPONSE----------------------------------------------------------#

    # ------------------------------------------DATA ANALYSIS----------------------------------------------------------#

    """

    1.	GET_SUMMARY_SMS: 
    [GAIA PRO,RN,Cause,GPI,Device(Mains/Battery),BAT Voltage in Volt,Last button serial number,GSM_RSSI,WiFi_RSSI,GPRS 
    Connectivity(OK/NOK),Button 1,Button 2,Button 3,Button 4,Button 5,Data sending interval,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 1 
    c)	Cause: Test that it is one of the following: Midnight, Conn Loss, P Int, P Ext, Bat Voltage, Sig Strength, 
    On Demand, Dvc Restart, Ser No Reset, Tamper Open, Tamper Close, FOTA Ongoing
    d)	Poll ID: Test that it is exactly 14 characters
    e)	Power Mode: Test that it is one of the following: I, V
    f)	Battery Voltage: Test that it is a float number of the format X.XX, and is also greater than 2.6 and less than 4.3
    g)	Total Button Count: Test that it is an integer between 0 and 10000000 (10 million)
    h)	GSM Signal Strength: Test that it is an integer between 0 and 100
    i)	WiFi Signal Strength: Test that it is an integer between 0 and 100
    j)	Data Sending Status: Test that it is one of the following: OK, NOK
    k)	Button Press Count for Buttons 1, 2, 3, 4, 5: For each one, test that it is an integer between 0 and 10000000 
    (10 million)
    l)	Data Sending Interval: Test that it is an integer between 0 and 1440 
    m)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'

    2.	GET_CONFIG_SMS:
    [GAIA PRO,RN,GPI,APN,IMEI,MCN 1,MCN 2,MCN 3,FCN,LCN,Network(WiFi/GSM),2G/4G/NB,Hardware Version,Firmware Version,
    Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 2
    c)	Poll ID: Test that it is exactly 14 characters
    d)	APN: Test that it is one of the following: gaia, m2m.gaia.co.in, jionet, airtelgprs.com, www
    e)	IMEI: Test that it is 15 digits (should be all numbers)
    f)	 MCN_1, MCN_2, MCN_3, FCN, LCN: Test that it is 10 digits (should be all numbers)
    g)	Network Mode: Test that it is one of the following: G, W
    h)	GSM Module Type: Test that it is one of the following: 2G, 4G, NB
    i)	Hardware version: Test that it is a float number with the format X.X, and is also less than 10.0
    j)	Firmware version: Test that it is a float number with the format X.X, and is also less than 10.0
    k)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'


    3.	GET_WIFI_CONFIG_SMS:
    [GAIA PRO,RN,WIFI_SSID,WIFI_PASSWORD,WIFI_URL,Wifi_Fota_Url,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 3
    c)	WiFi_SSID and WiFi_Password: Ignore checks on because they can be anything
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530' 

    12.	GET_WIFI_URL:
    [GAIA PRO,RN,WIFI_SSID,WIFI_PASSWORD,WIFI_URL,Wifi_Fota_Url,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 12
    c)	WIFI_URL: Check the WiFi_URL have either an "http://" or an "https://" at the beginning 
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530' 

    13.	GET_WIFI_FOTA_URL:
    [GAIA PRO,RN,WIFI_SSID,WIFI_PASSWORD,WIFI_URL,Wifi_Fota_Url,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 13
    c)	Wifi Fota Url: Check the WiFI_FOTA_URL have either an "http://" or an "https://" at the beginning 
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'

    4.	GET_INTERVALS_SMS:
    [GAIA PRO,RN,Data sending interval,KeepAlive interval,Network Check Time window,Summary Report SMS Waiting Interval,
    Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 4
    c)	Data sending interval, KeepAlive interval, Network Check Time Window, Summary Report SMS Waiting Interval: 
    Check that each of the interval parameters has a value between 0 and 2880
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'


    5.	GET_URL: 
    [GAIA PRO,RN,Complete URL,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 5
    c)	Complete URL: Check that the URL parameter has either an "http://" or an "https://" at the beginning 
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'


    6.	GET_CCID:
    [GAIA PRO,RN,CCID,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 6
    c)	CCID: Check that the CCID parameter is strictly a 19-digit number
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'


    7.	GET_BAD_BUTTON_LIMIT:
    [GAIA PRO,RN,Bad button limit,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 7
    c)	Bad button limit: Check that the bad button limit parameter is a number >=1.
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'


    8.	GET_FOTA_URL:
    [GAIA PRO,RN,FOTA URL,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 8
    c)	FOTA URL: Check that the FOTA URL parameter has either an "http://" or an "https://" at the beginning
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'


    9.	GET_FOTA_TIME:
    [GAIA PRO,RN,FOTA TIME,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 9
    c)	FOTA TIME:
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'


    10.	FW_UPDATE:
    [GAIA PRO,RN,FOTA_Failure,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 10
    c)	FOTA_Failure:
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'


    11.	GET_THRESHOLDS_SMS:
    [GAIA PRO,RN,Battery Alert,Signal Alert,Timestamp]
    a)	GAIA PRO: GAIA AAI
    b)	RN: 11
    c)	Battery Alert and Singnal Alert: Check that both the threshold parameters have a value between 0 and 100. 
    d)	Timestamp: Test that it has the format 'DD-MM-YY HH:MM:SS +0530'

    """

    @pyqtSlot()
    def load_get(self):
        try:
            data_frame = [''] * len(self.file_response_no)
            common_data_frame = pd.read_excel(self.Record_File, sheet_name=self.Record_File_Common_Sheet,
                                              index_col=[0]).dropna(how="all")
            common_data_frame.to_excel(self.writer_Record_File, sheet_name=self.Record_File_Common_Sheet)
            for i in range(len(self.file_response_no)):
                if int(self.file_response_no[i]) > 0:
                    data_frame[i] = pd.read_excel(self.Record_File, sheet_name=self.file_commands[i],
                                                  index_col=[0]).dropna(how="all")
                    data_frame[i].to_excel(self.writer_Record_File, sheet_name=self.file_commands[i])

            self.writer_Record_File.save()
        except Exception as e:
            print("Exception load: ", e)
            self.output_te.append("<b>Exception load_get: " + str(e) + ", close the sheet or empty it after adding a "
                                                                       "new command sheet.")

    @pyqtSlot()
    def get_response_data(self):
        try:
            self.load_get()

            if self.file_default_values[10].strip() == 'AAI':
                df1 = pd.read_excel(self.Record_File, sheet_name='GET_SUMMARY_SMS', index_col=[0]).dropna(how="all")
                df1 = df1.fillna('')
                d1 = df1.reset_index(drop=True).style \
                    .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                    .applymap(self.get_highlight_1_RN, subset='RN') \
                    .applymap(self.get_highlight_1_Cause, subset='Cause') \
                    .applymap(self.get_highlight_1_GPI, subset='GPI') \
                    .applymap(self.get_highlight_1_Mode, subset='Device(Mains/Battery)') \
                    .applymap(self.get_highlight_1_BAT, subset='BAT Voltage in Volt') \
                    .applymap(self.get_highlight_1_Count, subset=['Last button serial number', 'Button 1',
                                                                  'Button 2', 'Button 3', 'Button 4', 'Button 5']) \
                    .applymap(self.get_highlight_1_RSSI, subset=['GSM_RSSI', 'WiFi_RSSI']) \
                    .applymap(self.get_highlight_1_GPRS, subset='GPRS Connectivity(OK/NOK)') \
                    .applymap(self.get_highlight_1_Interval, subset='Data sending interval') \
                    .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
                print("df:", df1)
                d1.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_SUMMARY_SMS')
                self.writer_Record_File.save()
            elif self.file_default_values[10].strip() == 'Insight':
                df1 = pd.read_excel(self.Record_File, sheet_name='GET_SUMMARY_SMS', index_col=[0]).dropna(how="all")
                df1 = df1.fillna('')
                d1 = df1.reset_index(drop=True).style \
                    .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                    .applymap(self.get_highlight_1_RN, subset='RN') \
                    .applymap(self.get_highlight_1_Cause, subset='Cause') \
                    .applymap(self.get_highlight_1_GPI, subset='GPI') \
                    .applymap(self.get_highlight_1_Mode, subset='Device(Mains/Battery)') \
                    .applymap(self.get_highlight_1_BAT, subset='BAT Voltage in Volt') \
                    .applymap(self.get_highlight_1_Count, subset='Last Data serial number') \
                    .applymap(self.get_highlight_1_RSSI, subset=['GSM_RSSI', 'WiFi_RSSI']) \
                    .applymap(self.get_highlight_1_GPRS, subset='GPRS Connectivity: OK/NOK') \
                    .applymap(self.get_highlight_1_Interval, subset=['Data sending interval_power', 'Data sending interval_battery']) \
                    .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
                print("df:", df1)
                d1.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_SUMMARY_SMS')
                self.writer_Record_File.save()

            df2 = pd.read_excel(self.Record_File, sheet_name='GET_CONFIG_SMS', index_col=[0]).dropna(how="all")
            df2 = df2.fillna('')
            d2 = df2.reset_index(drop=True).style \
                .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                .applymap(self.get_highlight_2_RN, subset='RN') \
                .applymap(self.get_highlight_1_GPI, subset='GPI') \
                .applymap(self.get_highlight_2_APN, subset='APN') \
                .applymap(self.get_highlight_2_IMEI, subset='IMEI') \
                .applymap(self.get_highlight_2_MFL, subset=['MCN 1', 'MCN 2', 'MCN 3', 'FCN', 'LCN']) \
                .applymap(self.get_highlight_2_NetMode, subset='Network(WiFi/GSM)') \
                .applymap(self.get_highlight_2_GSM, subset='2G/4G/NB') \
                .applymap(self.get_highlight_2_Ver, subset=['Hardware Version', 'Firmware Version']) \
                .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
            print("df:", df2)
            d2.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_CONFIG_SMS')
            self.writer_Record_File.save()

            df3 = pd.read_excel(self.Record_File, sheet_name='GET_WIFI_CONFIG_SMS', index_col=[0]).dropna(how="all")
            df3 = df3.fillna('')
            d3 = df3.reset_index(drop=True).style \
                .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                .applymap(self.get_highlight_3_RN, subset='RN') \
                .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
            print("df:", df3)
            d3.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_WIFI_CONFIG_SMS')
            self.writer_Record_File.save()

            df12 = pd.read_excel(self.Record_File, sheet_name='GET_WIFI_URL', index_col=[0]).dropna(how="all")
            df12 = df12.fillna('')
            d12 = df12.reset_index(drop=True).style \
                .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                .applymap(self.get_highlight_12_RN, subset='RN') \
                .applymap(self.get_highlight_12_URL, subset='WIFI_URL') \
                .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
            print("df:", df12)
            d12.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_WIFI_URL')
            self.writer_Record_File.save()

            df13 = pd.read_excel(self.Record_File, sheet_name='GET_WIFI_FOTA_URL', index_col=[0]).dropna(how="all")
            df13 = df13.fillna('')
            d13 = df13.reset_index(drop=True).style \
                .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                .applymap(self.get_highlight_13_RN, subset='RN') \
                .applymap(self.get_highlight_12_URL, subset='Wifi_Fota_Url') \
                .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
            print("df:", df13)
            d13.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_WIFI_FOTA_URL')
            self.writer_Record_File.save()

            if self.file_default_values[10].strip() == 'AAI':
                df4 = pd.read_excel(self.Record_File, sheet_name='GET_INTERVALS_SMS', index_col=[0]).dropna(how="all")
                df4 = df4.fillna('')
                d4 = df4.reset_index(drop=True).style \
                    .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                    .applymap(self.get_highlight_4_RN, subset='RN') \
                    .applymap(self.get_highlight_4_Interval, subset=['Data sending interval', 'KeepAlive interval',
                                                                     'Network Check Time window',
                                                                     'Summary Report SMS Waiting Interval']) \
                    .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
                print("df:", df4)
                d4.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_INTERVALS_SMS')
                self.writer_Record_File.save()
            elif self.file_default_values[10].strip() == 'Insight':
                df4 = pd.read_excel(self.Record_File, sheet_name='GET_INTERVALS_SMS', index_col=[0]).dropna(how="all")
                df4 = df4.fillna('')
                d4 = df4.reset_index(drop=True).style \
                    .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                    .applymap(self.get_highlight_4_RN, subset='RN') \
                    .applymap(self.get_highlight_4_Interval, subset=['Data sending interval_power', 'Data sending interval_battery',
                                                                     'Network Check Time window',
                                                                     'Summary Report SMS sending Interval']) \
                    .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
                print("df:", df4)
                d4.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_INTERVALS_SMS')
                self.writer_Record_File.save()

            df5 = pd.read_excel(self.Record_File, sheet_name='GET_URL', index_col=[0]).dropna(how="all")
            df5 = df5.fillna('')
            d5 = df5.reset_index(drop=True).style \
                .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                .applymap(self.get_highlight_5_RN, subset='RN') \
                .applymap(self.get_highlight_12_URL, subset='Complete URL') \
                .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
            print("df:", df5)
            d5.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_URL')
            self.writer_Record_File.save()

            df6 = pd.read_excel(self.Record_File, sheet_name='GET_CCID', index_col=[0]).dropna(how="all")
            df6 = df6.fillna('')
            d6 = df6.reset_index(drop=True).style \
                .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                .applymap(self.get_highlight_6_RN, subset='RN') \
                .applymap(self.get_highlight_6_CCID, subset='CCID') \
                .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
            print("df:", df6)
            d6.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_CCID')
            self.writer_Record_File.save()

            df7 = pd.read_excel(self.Record_File, sheet_name='GET_BAD_BUTTON_LIMIT', index_col=[0]).dropna(how="all")
            df7 = df7.fillna('')
            d7 = df7.reset_index(drop=True).style \
                .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                .applymap(self.get_highlight_7_RN, subset='RN') \
                .applymap(self.get_highlight_7_Bad_Button, subset='Bad button limit') \
                .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
            print("df:", df7)
            d7.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_BAD_BUTTON_LIMIT')
            self.writer_Record_File.save()

            df8 = pd.read_excel(self.Record_File, sheet_name='GET_FOTA_URL', index_col=[0]).dropna(how="all")
            df8 = df8.fillna('')
            d8 = df8.reset_index(drop=True).style \
                .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                .applymap(self.get_highlight_8_RN, subset='RN') \
                .applymap(self.get_highlight_12_URL, subset='FOTA URL') \
                .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
            print("df:", df8)
            d8.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_FOTA_URL')
            self.writer_Record_File.save()

            df9 = pd.read_excel(self.Record_File, sheet_name='GET_FOTA_TIME', index_col=[0]).dropna(how="all")
            df9 = df9.fillna('')
            d9 = df9.reset_index(drop=True).style \
                .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                .applymap(self.get_highlight_9_RN, subset='RN') \
                .applymap(self.get_highlight_9_Fota_Time, subset='FOTA TIME') \
                .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
            print("df:", df9)
            d9.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_FOTA_TIME')
            self.writer_Record_File.save()

            df10 = pd.read_excel(self.Record_File, sheet_name='FW_UPDATE', index_col=[0]).dropna(how="all")
            df10 = df10.fillna('')
            d10 = df10.reset_index(drop=True).style \
                .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                .applymap(self.get_highlight_10_RN, subset='RN') \
                .applymap(self.get_highlight_10_Fota_Update, subset='FOTA_Status') \
                .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
            print("df:", df10)
            d10.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='FW_UPDATE')
            self.writer_Record_File.save()

            if self.file_default_values[10].strip() == 'AAI':
                df11 = pd.read_excel(self.Record_File, sheet_name='GET_THRESHOLDS_SMS', index_col=[0]).dropna(how="all")
                df11 = df11.fillna('')
                d11 = df11.reset_index(drop=True).style \
                    .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                    .applymap(self.get_highlight_11_RN, subset='RN') \
                    .applymap(self.get_highlight_11_Threshold, subset=['Battery Alert', 'Signal Alert']) \
                    .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
                print("df:", df11)
                d11.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_THRESHOLDS_SMS')
                self.writer_Record_File.save()
            elif self.file_default_values[10].strip() == 'Insight':
                df11 = pd.read_excel(self.Record_File, sheet_name='GET_THRESHOLDS_SMS', index_col=[0]).dropna(how="all")
                df11 = df11.fillna('')
                d11 = df11.reset_index(drop=True).style \
                    .applymap(self.get_highlight_GAIA, subset='GAIA PRO') \
                    .applymap(self.get_highlight_11_RN, subset='RN') \
                    .applymap(self.get_highlight_11_Threshold, subset=['Battery Threshold', 'Signal Threshold']) \
                    .applymap(self.get_highlight_11_Temp, subset=['TEMP_ALERT_THRESHOLD_UPPER value', 'TEMP_ALERT_THRESHOLD_Lower value']) \
                    .applymap(self.get_highlight_1_Timestamp, subset='Timestamp')
                print("df:", df11)
                d11.to_excel(self.writer_Record_File, engine='openpyxl', sheet_name='GET_THRESHOLDS_SMS')
                self.writer_Record_File.save()

            print("\nANALYSIS of GET Excel Sheets DONE :)")
            self.output_te.append("<b>\nANALYSIS of GET Excel Sheets DONE :)")

        except Exception as e:
            print('Error in get_response_data', e)
            self.output_te.append('<b>Error in get_response_data: ' + str(e))

    @pyqtSlot(str)
    def get_highlight_GAIA(self, cell_value):  # COMMON_PARAMETER
        print(cell_value, type(cell_value))
        try:
            GAIA_project_list = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['GAIA PROJECT'])
                                     .dropna(how="any")['GAIA PROJECT'])
            if cell_value != '':
                cell_value = str(cell_value).strip()
                for i in GAIA_project_list:
                    if i in cell_value:
                        return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_1_RN(self, cell_value):  # 1=GET_SUMMARY_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 1:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_1_Cause(self, cell_value):  # 1=GET_SUMMARY_SMS
        print(cell_value, type(cell_value))
        try:
            cause_list = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['SUMMARY CAUSE'])
                              .dropna(how="any")['SUMMARY CAUSE'])
            if cell_value != '':
                cell_value = str(cell_value).strip()
                for i in cause_list:
                    if i in cell_value:
                        return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_1_GPI(self, cell_value):  # 1=GET_SUMMARY_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if len(cell_value) == 14:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_1_Mode(self, cell_value):  # 1=GET_SUMMARY_SMS
        print(cell_value, type(cell_value))
        try:
            mode_list = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['SUMMARY MODE'])
                             .dropna(how="any")['SUMMARY MODE'])
            if cell_value != '':
                cell_value = str(cell_value).strip()
                for i in mode_list:
                    if i in cell_value:
                        return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_1_BAT(self, cell_value):  # 1=GET_SUMMARY_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = float(str(cell_value).strip())
                if cell_value >= 2.6 and cell_value <= 4.3:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_1_Count(self, cell_value):  # 1=GET_SUMMARY_SMS
        print(str(cell_value), type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) in range(0, 10000001):
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_1_RSSI(self, cell_value):  # 1=GET_SUMMARY_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = int(str(cell_value).strip())
                if cell_value in range(0, 101):
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_1_GPRS(self, cell_value):  # 1=GET_SUMMARY_SMS
        print(cell_value, type(cell_value))
        try:
            GPRS_list = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['SUMMARY GPRS'])
                             .dropna(how="any")['SUMMARY GPRS'])
            if cell_value != '':
                cell_value = str(cell_value).strip()
                for i in GPRS_list:
                    if i in cell_value:
                        return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_1_Interval(self, cell_value):  # 1=GET_SUMMARY_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) in range(0, 1441):
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_1_Timestamp(self, cell_value):  # 1=GET_SUMMARY_SMS
        print(str(cell_value), type(cell_value))
        if cell_value != '':
            cell_value = str(cell_value).strip()

            cell_datetime = str(cell_value[:-6])
            cell_zone = cell_value[-5:]
            original_datetime = '%y-%m-%d %H:%M:%S'  # DD-MM-YY HH:MM:SS
            original_zone = strftime("%z", gmtime())  # +0530
            print(cell_datetime, cell_zone, original_zone)

            try:
                date_obj = datetime.datetime.strptime(cell_datetime, original_datetime)
                print(date_obj)
                if str(cell_zone) == str(original_zone):
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            except ValueError:
                print("Incorrect data format, should be YYYY-MM-DD")
                return 'background-color: #FA8072'  # salmon red
        else:
            return 'background-color: None'

    @pyqtSlot(str)
    def get_highlight_2_RN(self, cell_value):  # 2=GET_CONFIG_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 2:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    def get_highlight_2_APN(self, cell_value):  # 2=GET_CONFIG_SMS
        print(cell_value, type(cell_value))
        try:
            APN_list = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['CONFIG APN'])
                            .dropna(how="any")['CONFIG APN'])
            if cell_value != '':
                cell_value = str(cell_value).strip()
                for i in APN_list:
                    if i in cell_value:
                        return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_2_IMEI(self, cell_value):  # 2=GET_CONFIG_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) >= 0 and len(cell_value) == 15:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_2_MFL(self, cell_value):  # 2=GET_CONFIG_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if str(cell_value) == "GAIASC" or (int(cell_value) >= 0 and len(cell_value) == 10):
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    def get_highlight_2_NetMode(self, cell_value):  # 2=GET_CONFIG_SMS
        print(cell_value, type(cell_value))
        try:
            NETmode_list = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['CONFIG NETMODE'])
                                .dropna(how="any")['CONFIG NETMODE'])
            if cell_value != '':
                cell_value = str(cell_value).strip()
                for i in NETmode_list:
                    if i in cell_value:
                        return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    def get_highlight_2_GSM(self, cell_value):  # 2=GET_CONFIG_SMS
        print(cell_value, type(cell_value))
        try:
            GSM_list = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['CONFIG GSM'])
                            .dropna(how="any")['CONFIG GSM'])
            if cell_value != '':
                cell_value = str(cell_value).strip()
                for i in GSM_list:
                    if i in cell_value:
                        return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_2_Ver(self, cell_value):  # 2=GET_CONFIG_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = float(str(cell_value).strip())
                if cell_value >= 0.0 and cell_value <= 10.0:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_3_RN(self, cell_value):  # 3=GET_WIFI_CONFIG_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 3:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_12_RN(self, cell_value):  # 12=GET_WIFI_URL
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 12:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_12_URL(self, cell_value):  # 12=GET_WIFI_URL
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                print(cell_value[:7], cell_value[:8])
                if cell_value[:7] == "http://" or cell_value[:8] == "https://":
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_13_RN(self, cell_value):  # 13=GET_WIFI_FOTA_URL
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 13:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_4_RN(self, cell_value):  # 4=GET_INTERVALS_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 4:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_4_Interval(self, cell_value):  # 4=GET_INTERVALS_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) in range(0, 2881):
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_5_RN(self, cell_value):  # 5=GET_URL
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 5:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_6_RN(self, cell_value):  # 6=GET_CCID
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 6:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_6_CCID(self, cell_value):  # 6=GET_CCID
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if ((cell_value.isdigit() and int(cell_value) >= 0)
                    or ((cell_value[-1].isalpha() or cell_value[-1].isdigit()) and cell_value[:-1].isdigit()
                        and int(cell_value[:-1]) >= 0)
                    and (len(cell_value) == 20 or len(cell_value) == 19)):
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_7_RN(self, cell_value):  # 7=GET_BAD_BUTTON_LIMIT
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 7:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_7_Bad_Button(self, cell_value):  # 7=GET_BAD_BUTTON_LIMIT
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) >= 1:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_8_RN(self, cell_value):  # 8=GET_FOTA_URL
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 8:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_9_RN(self, cell_value):  # 9=GET_FOTA_TIME
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 9:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_9_Fota_Time(self, cell_value):  # 9=GET_FOTA_TIME
        print(str(cell_value), type(cell_value))
        if cell_value != '':
            cell_value = str(cell_value).strip()

            cell_datetime = str(cell_value[-8:])
            original_datetime = '%H:%M:%S'  # HH:MM:SS
            print(cell_datetime, cell_value[0])
            try:
                date_obj = time.strptime(cell_datetime, original_datetime)
                print(date_obj)
                if int(cell_value[0]) in range(1, 8):
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            except ValueError:
                print("Incorrect time format, should be HH:MM:SS")
                return 'background-color: #FA8072'  # salmon red
        else:
            return 'background-color: None'

    @pyqtSlot(str)
    def get_highlight_10_RN(self, cell_value):  # 10=FW_UPDATE
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 10:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_10_Fota_Update(self, cell_value):  # 10=FW_UPDATE
        print(cell_value, type(cell_value))
        try:
            update_list = list(pd.read_excel(self.Database, 'Sheet1', dtype=str, usecols=['FOTA STATUS'])
                               .dropna(how="any")['FOTA STATUS'])
            if cell_value != '':
                cell_value = str(cell_value).strip()
                for i in update_list:
                    if i in cell_value:
                        return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_11_RN(self, cell_value):  # 11=GET_THRESHOLDS_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) == 11:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_11_Threshold(self, cell_value):  # 11=GET_THRESHOLDS_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if cell_value.isdigit() and int(cell_value) in range(0, 101):
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot(str)
    def get_highlight_11_Temp(self, cell_value):  # 11=GET_THRESHOLDS_SMS
        print(cell_value, type(cell_value))
        try:
            if cell_value != '':
                cell_value = str(cell_value).strip()
                if float(cell_value) <= 1000 and float(cell_value) >= -100:
                    return 'background-color: #98FB98'  # pale green
                return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    @pyqtSlot()
    def load_set(self):
        try:
            data_frame = pd.read_excel(self.set_Record_File, index_col=[0]).dropna(how="all")
            data_frame.to_excel(self.writer_set_Record_File)

            self.writer_set_Record_File.save()
        except Exception as e:
            print("Exception load:", e)
            self.output_te.append("<b>Exception load_set: " + str(e) + ", try closing the sheet.")

    @pyqtSlot()
    def set_response_data(self):
        try:
            df = pd.read_excel(self.set_Record_File, index_col=[0]).dropna(how="all")
            df = df.fillna('')

            d = df.reset_index(drop=True).style.applymap(self.set_highlight_cells, subset=self.set_response_parameters)
            print("df:", df)
            d.to_excel(self.writer_set_Record_File, engine='openpyxl')
            self.writer_set_Record_File.save()

            print("\nANALYSIS of SET Excel Sheets DONE :)")
            self.output_te.append("<b>\nANALYSIS of SET Excel Sheets DONE :)")

        except Exception as e:
            print('Error in set_response_data', e)
            self.output_te.append('<b>Error in set_response_data: ' + str(e))

    @pyqtSlot(str)
    def set_highlight_cells(self, cell_value):
        try:
            if cell_value != '':
                cell_value = str(cell_value)
                if "NO RESPONSE" in cell_value:
                    return 'background-color: #ffff80'  # light yellow
                elif "SUCCESS" in cell_value:
                    return 'background-color: #98FB98'  # pale green
                elif "FAIL" in cell_value:
                    return 'background-color: #FA8072'  # salmon red
            else:
                return 'background-color: None'
        except ValueError:
            return 'background-color: #ffff80'

    # ---------------------------------------END DATA ANALYSIS---------------------------------------------------------#


if __name__ == '__main__':
    clear()
    import sys

    app = QApplication(sys.argv)
    w = Widget()
    w.showNormal()
    sys.exit(app.exec_())

