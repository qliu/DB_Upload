#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys, os
## Import PyQt4
from PyQt4 import QtGui, QtCore
## Import xlrd for reading XLS files with Python
from xlrd import open_workbook
## Import psycopg2 for PostgreSQL database connection
import psycopg2
## Import lib and module for sending emails
import smtplib
from email.mime.text import MIMEText

# Global constants
IMAGE_ROOT = 'data/images'
## DB connection parameters

### Server "pitondc1" DB connection
LOCAL_DRIVER = "LOCAL_DRIVER"
UPLOAD_DATA_ROOT = "UPLOAD_DATA_ROOT"
ROOT_PATH = "ROOT_PATH"
DB_DATABASE = "DB_DATABASE"
DB_USER = "DB_USER"
DB_PASSWORD = "DB_PASSWORD"
DB_HOST = "DB_HOST"
DB_PORT = "DB_PORT"
DB_SCHEMA = "DB_SCHEMA"

# EMAIL SETTINGS
EMAIL_USE_TLS = "EMAIL_USE_TLS"
EMAIL_HOST = "EMAIL_HOST"
EMAIL_PORT = "EMAIL_PORT"
EMAIL_HOST_USER = "EMAIL_HOST_USER"
EMAIL_HOST_PASSWORD = "EMAIL_HOST_PASSWORD"
# Admin email
ADMIN_EMAIL_ADDRESS = "ADMIN_EMAIL_ADDRESS"
TO_EMAIL_ADDRESS = ["TO_EMAIL_ADDRESS"]

class MainWindow(QtGui.QMainWindow):
    
    def __init__(self):
        super(MainWindow, self).__init__()
        
        self.init_UI()
        
    def init_UI(self):
        '''
        Initialize GUI widgets
        '''
        
        # =========
        # Initials
        # =========        
        
        ## Initial GUI parameters
        QtGui.QToolTip.setFont(QtGui.QFont('SansSerif',8))
        
        ## Initial window parameters
        w_size_width = 720
        w_size_height = 480
        w_title = "DB Upload"
        
        
        # =========
        # Main Window
        # =========  
        
        ## Set window size
        self.resize(w_size_width,w_size_height)
        ## Move window to screen center
        self.move_center()
        ## Set window title
        self.setWindowTitle(w_title)
        ## Set window title bar icon
        self.setWindowIcon(QtGui.QIcon('%s/icon_upload.png' % IMAGE_ROOT))
        
        
        # =========
        # Menu & Status
        # =========  
        
        ## Open File
        self.openfile_menu = QtGui.QAction(QtGui.QIcon('%s/icon_openfile.png' % IMAGE_ROOT),'&Open...',self)
        self.openfile_menu.setShortcut('Ctrl+O')
        self.openfile_menu.setStatusTip('Open File')
        self.openfile_menu.triggered.connect(self.show_file_dialog)
        
        ## Quit
        self.quit_menu = QtGui.QAction(QtGui.QIcon('%s/icon_quit.png' % IMAGE_ROOT),'&Quit',self)
        self.quit_menu.setShortcut('Ctrl+Q')
        self.quit_menu.setStatusTip('Quit')
        self.quit_menu.triggered.connect(self.close)
        
        ## Upload
        self.uploadfile_menu = QtGui.QAction(QtGui.QIcon('%s/icon_uploadfile.png' % IMAGE_ROOT),'&Upload',self)
        self.uploadfile_menu.setShortcut('Ctrl+U')
        self.uploadfile_menu.triggered.connect(self.upload_file)
        self.disable_upload_menu()
        
        ## Help
        self.help_menu = QtGui.QAction(QtGui.QIcon('%s/icon_help.png' % IMAGE_ROOT),'DB Upload Help Document',self)
        self.help_menu.setShortcut('F1')
        self.help_menu.setStatusTip('DB Upload Help Document')
        self.help_menu.triggered.connect(self.help_doc)
        self.github_menu = QtGui.QAction(QtGui.QIcon('%s/icon_github.png' % IMAGE_ROOT), 'See us on GitHub',self)
        self.github_menu.setStatusTip('See us on GitHub')
        self.github_menu.triggered.connect(self.open_github)
        
        ## Menu bar
        menu_bar = self.menuBar()
        ### File
        m_file = menu_bar.addMenu('&File')
        m_file.addAction(self.openfile_menu)
        m_file.addSeparator()
        m_file.addAction(self.quit_menu)
        ### Upload
        m_upload = menu_bar.addMenu('&Upload File')
        m_upload.addAction(self.uploadfile_menu)
        ### Help
        m_help = menu_bar.addMenu('&Help')
        m_help.addAction(self.help_menu)
        m_help.addSeparator()
        m_help.addAction(self.github_menu)
        
        ## Status bar
        self.statusBar()
        
        # =========
        # Text Edit
        # =========        

        ## Label
        self.filepath_label = QtGui.QLabel(self)
        self.filepath_label.setFixedWidth(600)
        self.filepath_label.setWordWrap(True)
        self.filepath_label.setTextFormat(QtCore.Qt.RichText)
        self.filepath_label.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
        self.filepath_label.move(20,40)
        self.filepath_label.setText("Click <b>File -> Open...</b> to select a file to upload.")
        
        self.error_label = QtGui.QLabel(self)
        self.error_label.setFixedWidth(600)
        self.error_label.setWordWrap(True)
        self.error_label.setOpenExternalLinks(True)
        self.error_label.setTextFormat(QtCore.Qt.RichText)
        self.error_label.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
        self.error_label.move(20,200)

        # =========
        # Buttons
        # =========
        
        ## Quit Button
#        quit_btn = QtGui.QPushButton('Quit', self)
#        quit_btn.setToolTip('Quit')
#        quit_btn.resize(quit_btn.sizeHint())
#        quit_btn.move(self.frameSize().width()-quit_btn.frameSize().width()*1.5,
#                      self.frameSize().height()-quit_btn.frameSize().height()*2)
##        quit_btn.clicked.connect(QtCore.QCoreApplication.instance().quit)
#        quit_btn.clicked.connect(self.close)
        
        ## Upload Button
#        self.upload_btn = QtGui.QPushButton('Upload',self)
#        self.upload_btn.setToolTip('Upload')
#        self.upload_btn.resize(self.upload_btn.sizeHint())
#        self.upload_btn.move(300,160)
#        self.upload_btn.clicked.connect(self.upload_file)
#        self.upload_btn.setVisible(False)
        
        
        # =========
        # Show the main window widget
        # =========
        self.show()
        
    # =========================
    #  overriding functions
    # =========================    
        
    def closeEvent(self,event):
        '''
        Override closeEvent() to pop up a msg box before closing the main window
        '''
        reply = QtGui.QMessageBox.question(self,'Quit?',"Are you sure to quit?",
            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No, QtGui.QMessageBox.No)
        if reply == QtGui.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()    

    # =========================
    # util functions
    # =========================

    def move_center(self):
        '''
        Move window to screen center
        '''
        rec_frame = self.frameGeometry()
        center_point = QtGui.QDesktopWidget().availableGeometry().center()
        rec_frame.moveCenter(center_point)
        self.move(rec_frame.topLeft())
    
    def disable_upload_menu(self):
        self.uploadfile_menu.setEnabled(False)
        self.uploadfile_menu.setStatusTip('Upload is not available. Please select a file.')
        
    def enable_upload_menu(self):
        self.uploadfile_menu.setEnabled(True)
        self.uploadfile_menu.setStatusTip('Upload')
        
    def send_email(self,email_subject,email_message):
        '''
        Send notification email when new table added to database
        '''
        # Initialize SMTP server
        email_server = smtplib.SMTP(EMAIL_HOST,EMAIL_PORT)
        email_server.ehlo()
        email_server.starttls()
        email_server.login(EMAIL_HOST_USER,EMAIL_HOST_PASSWORD)
        # Send Email
        msg = MIMEText(email_message)
        msg['Subject'] = email_subject
        msg['From'] = ADMIN_EMAIL_ADDRESS
#        msg['To'] = TO_EMAIL_ADDRESS
        email_server.sendmail(ADMIN_EMAIL_ADDRESS,TO_EMAIL_ADDRESS,msg.as_string())
        email_server.quit()
        
    # =========================
    # main functions
    # =========================
    
    def check_file_exist(self):
        file_full_path = self.file_full_path   
        file_name = file_full_path[file_full_path.rfind("/")+1:]
        file_path = file_full_path[:file_full_path.rfind("/")+1]
        file_ext = file_name[file_name.find(".")+1:]
        table_name = file_name[:file_name.find(".")]
        self.header_full_path = "%s%s_header.%s" % (file_path,table_name,file_ext)
        self.data_full_path = "%s%s.csv" % (file_path,table_name)
        error_msg = ""
        return_dict = {
                        "is_file_exists":False,
                        "error_msg":error_msg
                       }
        if os.path.exists(file_full_path):
            if os.path.exists(self.header_full_path):
                if os.path.exists(self.data_full_path):
                    return_dict["is_file_exists"] = True
                else:
                    error_msg = "Upload data(CSV file) NOT found!<br/><br/>No such file: %s" % self.data_full_path
            else:
                error_msg = "Header file NOT found!<br/><br/>No such file: %s" % self.header_full_path
        else:
            error_msg = "File NOT found!<br/><br/>No such file: %s" % self.file_full_path
        return_dict["error_msg"] = error_msg
        return return_dict
        
    def upload_file(self):
        txt = self.filepath_label.text().replace('Select <b>Upload File -> Upload</b> when you are ready to upload file.','File uploading...')
        self.filepath_label.setText(txt)
        
        file_full_path = self.file_full_path    
        file_name = file_full_path[file_full_path.rfind("/")+1:]
        file_path = file_full_path[:file_full_path.rfind("/")+1]
        table_name = file_name[:file_name.find(".")]
        
        # Initialize variables
        headers = []
        data_types = []
        sql_createtable_columns = ""
        sql_insert_columns = ""
        
        # Open EXCEL workbook (*.xls)
        xls_workbook = open_workbook(self.header_full_path)
        # Read workbook sheets
        xls_sheet = xls_workbook.sheet_by_index(0)
        # Read headers
        for cell in xls_sheet.row(0):
            headers.append('"%s"' % cell.value.strip().replace('+','_up_').replace(' ','_').replace('/','_').replace('&','_').replace('-','_').replace('%','per').lower())
        pkey = headers[0]
        # Read data types
        for cell in xls_sheet.row(1):
            data_types.append(cell.value.strip())
        # Initialize DB connection
        dbcon_dc = psycopg2.connect(database = DB_DATABASE,
                                    user = DB_USER,
                                    host = DB_HOST,
                                    port = DB_PORT,
                                    password = DB_PASSWORD)

        # Create DB table
        try:
            cur_dc = dbcon_dc.cursor()
            ## Check if talbe exists
            exesql = "SELECT tablename FROM pg_tables WHERE schemaname = '%s'" % DB_SCHEMA
            cur_dc.execute(exesql)
            db_table_names = cur_dc.fetchall()
            if db_table_names.count((table_name,)) > 0:
                ### If table exists, pop up warning message box
                reply = QtGui.QMessageBox.warning(self,"Table Already Exists","Do you want to overwrite it?\nClick 'Yes' to overwrite existed table.\nClick 'No' to cancel upload.",
                        QtGui.QMessageBox.Yes | QtGui.QMessageBox.No, QtGui.QMessageBox.No)
                if reply == QtGui.QMessageBox.No:
                    raise Exception("Upload Warning","Table '%s' already exists. Upload has been canceled. Please rename the table and try again." % table_name)
            for i in range(0,len(headers)):
                sql_createtable_columns += "%s %s," % (headers[i],data_types[i])
                sql_insert_columns += "%s," % headers[i]
        ##  Use this when first column is the primary key
        #    exesql = "DROP TABLE IF EXISTS %s; CREATE TABLE %s (%s CONSTRAINT %s_pkey PRIMARY KEY (%s)) WITH (OIDS=FALSE); ALTER TABLE %s OWNER TO postgres;" % (table_name,table_name,sql_createtable_columns,table_name,pkey,table_name)
        ##  <- END
        ##  Use this when no primary key in the table
            exesql = "DROP TABLE IF EXISTS %s; CREATE TABLE %s (%s) WITH (OIDS=FALSE); ALTER TABLE %s OWNER TO postgres;" % (table_name,table_name,sql_createtable_columns[:-1],table_name)         
        ##  <- END
            # Execute query
            cur_dc.execute(exesql)
            # Commit all pending transactions
            dbcon_dc.commit()
            # Copy records from CSV
            file_path = file_path.replace(LOCAL_DRIVER,ROOT_PATH)
            exesql = "COPY %s FROM '%s' USING DELIMITERS ',' CSV;" % (table_name,"%s%s.csv" % (file_path,table_name))
            cur_dc.execute(exesql)         
            dbcon_dc.commit()
            error_msg = "<br/><hr><b>File uploaded successfully!</b><br/><br/>Please go to our admin site to <a href='http://pitondc1.piton.local/datacommons/admin/dcmetadata/sourcedatainventory/add/' target='_blank'>add metadata</a>."
            email_subject = "[DB_upload]New Table Added - %s" % table_name
            email_message = 'This is a notification that new tabe "%s" has been added to the database.' % table_name
            self.send_email(email_subject,email_message)
        except Exception, e:
            if e.args[0] == "Upload Warning":
                error_type, error_info = e.args
            else:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                error_type = exc_type.__name__
                error_info = exc_obj
                exesql = "DROP TABLE IF EXISTS %s;" % table_name
                dbcon_dc.rollback()
                cur_dc.execute(exesql)
                dbcon_dc.commit()
            error_msg = "<br/><hr><span style='color:red'><b>ERROR!</b></span><br/><br/><b>%s</b>: %s<br/><hr><br/>Please try again! Click <a href='http://pitondc1.piton.local/datacommons/db_upload_help_doc/#errors'>HERE</a> for help.<br/><br/>Click <b>File -> Open...</b> to select a file to upload." % (error_type,error_info)
        finally:
            # Close connection to Datacommons DB
            dbcon_dc.close()
            self.error_label.setText(error_msg)
            self.error_label.adjustSize()
            self.disable_upload_menu()
            
    def help_doc(self):
        url = "http://pitondc1.piton.local/datacommons/db_upload_help_doc/"
        QtGui.QDesktopServices.openUrl(QtCore.QUrl(url))
        
    def open_github(self):
        url = "https://github.com/qliu/DB_Upload"
        QtGui.QDesktopServices.openUrl(QtCore.QUrl(url))
        
        
    # =========================
    # GUI methods
    # =========================
    
    def show_dialog(self):
        '''
        Show Dialog input window
        '''
        text, ok = QtGui.QInputDialog.getText(self,'Input','Enter the path')
        if ok:
            self.line_text.setText(str(text))

    def show_file_dialog(self):
        '''
        Open File Dialog
        '''
        self.file_full_path = str(QtGui.QFileDialog.getOpenFileName(self,'Open File',UPLOAD_DATA_ROOT,'Excel Workbook (*.xls *.xlsx)'))
        
        if self.file_full_path:
            is_file_exists = self.check_file_exist()
            if is_file_exists["is_file_exists"]:
                file_list = "%s<br/>%s<br/>%s<br/>" % (self.file_full_path,self.header_full_path,self.data_full_path)
                # When file is ready to upload, change UI
                filepath_label_txt = "Upload file from: %s<br/><br/>Following files are ready:<br/>%s<br/>Select <b>Upload File -> Upload</b> when you are ready to upload file." % (self.file_full_path,file_list)
                self.error_label.setText("")
                self.enable_upload_menu()
            else:
                filepath_label_txt = "Upload file from: %s<br/><br/><hr><span style='color:red'><b>ERROR!</b></span><br/><br/>%s<br/><hr><br/>Please try again!<br/><br/>Click <b>File -> Open...</b> to select a file to upload." % (self.file_full_path,is_file_exists['error_msg'])
            self.filepath_label.setText(filepath_label_txt)
            self.filepath_label.adjustSize()            
      
def main():
    # Initial application and window widget
    app = QtGui.QApplication(sys.argv)
    main_window = MainWindow()

    # Mainloop here
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()