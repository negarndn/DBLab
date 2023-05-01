import datetime
from PyQt5.QtWidgets import *
import sys
from xlrd import *
from xlsxwriter import *
from PyQt5.uic import loadUiType
import pymysql
pymysql.install_as_MySQLdb()


ui, _ = loadUiType('library.ui')
login, _ = loadUiType('login.ui')


class Login(QWidget, login):
    def __init__(self):
        QWidget.__init__(self)
        self.window2 = None
        self.db = None
        self.cur = None
        self.setupUi(self)
        self.pushButton.clicked.connect(self.Handel_Login)
        self.Dark_Orange()

    def Handel_Login(self):
        # connect to db
        self.db = pymysql.connect(host='localhost', user='root', password='13791389', db='library')
        self.cur = self.db.cursor()

        username = self.lineEdit.text()
        password = self.lineEdit_2.text()

        sql = '''SELECT * FROM users'''
        self.cur.execute(sql)
        data = self.cur.fetchall()
        # print(data)
        for row in data:
            if username == row[1] and password == row[3]:
                self.window2 = MainApp()
                self.close()
                self.window2.show()
            else:
                self.label.setText('Make Sure You Enterd Your Username And Password Correctly')

    def Dark_Orange(self):
        style = open('themes/darkorange.css', 'r')
        style = style.read()
        self.setStyleSheet(style)


class MainApp(QMainWindow, ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.db = None
        self.cur = None
        self.setupUi(self)
        self.Handle_Ui_Changes()
        self.Handle_buttons()
        self.Show_Category()
        self.Show_Author()
        self.Show_Publisher()
        self.Show_Category_Combobox()
        self.Show_Author_Combobox()
        self.Show_Publisher_Combobox()
        self.Show_All_Operations()
        self.Show_All_Books()
        self.Show_All_Clients()

    def Handle_Ui_Changes(self):
        self.Hide_Themes()
        self.Dark_Orange()
        self.tabWidget.tabBar().setVisible(False)

    def Handle_buttons(self):
        self.pushButton_5.clicked.connect(self.Show_Themes)
        self.pushButton_21.clicked.connect(self.Hide_Themes)
        self.pushButton.clicked.connect(self.Open_Dey_To_Day_Tab)
        self.pushButton_2.clicked.connect(self.Open_Books_Tab)
        self.pushButton_3.clicked.connect(self.Open_Users_Tab)
        self.pushButton_4.clicked.connect(self.Open_Settings_Tab)
        self.pushButton_7.clicked.connect(self.Add_New_Book)
        self.pushButton_14.clicked.connect(self.Add_Category)
        self.pushButton_15.clicked.connect(self.Add_Author)
        self.pushButton_16.clicked.connect(self.Add_Publisher)
        self.pushButton_9.clicked.connect(self.Search_Books)
        self.pushButton_8.clicked.connect(self.Edit_Books)
        self.pushButton_10.clicked.connect(self.Delete_Books)
        self.pushButton_11.clicked.connect(self.Add_New_User)
        self.pushButton_13.clicked.connect(self.Login)
        self.pushButton_12.clicked.connect(self.Edit_User)
        self.pushButton_17.clicked.connect(self.Dark_Blue)
        self.pushButton_20.clicked.connect(self.Dark_Orange)
        self.pushButton_19.clicked.connect(self.Dark_Gray)
        self.pushButton_18.clicked.connect(self.Classic)
        self.pushButton_22.clicked.connect(self.Open_Clients_Tab)
        self.pushButton_24.clicked.connect(self.Add_New_Client)
        self.pushButton_25.clicked.connect(self.Search_Client)
        self.pushButton_23.clicked.connect(self.Edit_Clients)
        self.pushButton_26.clicked.connect(self.Delete_Client)
        self.pushButton_6.clicked.connect(self.Handle_Day_Operations)
        self.pushButton_29.clicked.connect(self.Export_Day_Operations)
        self.pushButton_27.clicked.connect(self.Export_Books)
        self.pushButton_28.clicked.connect(self.Export_Clients)

    def Show_Themes(self):
        self.groupBox_3.show()

    def Hide_Themes(self):
        self.groupBox_3.hide()
        
    def connectToDb(self):
        self.db = pymysql.connect(host='localhost', user='root', password='13791389', db='library')
        self.cur = self.db.cursor()

##########################################################
################### Opening Tabs #########################

    def Open_Dey_To_Day_Tab(self):
        self.tabWidget.setCurrentIndex(0)

    def Open_Books_Tab(self):
        self.tabWidget.setCurrentIndex(1)

    def Open_Users_Tab(self):
        self.tabWidget.setCurrentIndex(3)

    def Open_Clients_Tab(self):
        self.tabWidget.setCurrentIndex(2)

    def Open_Settings_Tab(self):
        self.tabWidget.setCurrentIndex(4)

##########################################################
################# Day Operations #########################

    def Handle_Day_Operations(self):
        self.connectToDb()
        book_title = self.lineEdit.text()
        op_type = self.comboBox.currentText()
        days = self.comboBox_2.currentIndex() + 1
        #print(days)
        op_date = datetime.date.today()
        #print(op_date)
        client_name = self.lineEdit_22.text()
        #print(client_name)
        to_time = op_date + datetime.timedelta(days=int(days))
        # print(op_date)
        # print(to_time)

        sql = '''SELECT book_name FROM book'''
        self.cur.execute(sql)
        data = self.cur.fetchall()
        sql2 = '''SELECT client_name FROM clients'''
        self.cur.execute(sql2)
        data2 = self.cur.fetchall()
        # print(data)
        flag = 0
        for row in data:
            for row2 in data2:
                if book_title == row[0]:
                    flag = 1
                    if client_name == row2[0] :
                        flag = 1
                        self.cur.execute('''
                                                    INSERT INTO dayoperations(book_name, client, op_type, days,
                                                    op_date, to_time)
                                                    VALUES(%s, %s, %s, %s, %s, %s)
                                                ''', (book_title, client_name, op_type, days, op_date, to_time))
                        self.db.commit()
                        self.statusBar().showMessage('New Operation Added!')
                        self.Show_All_Operations()
                        continue
        print(flag)
        if flag == 0:
            QMessageBox.about(self, "Error", "Please Enter Valid Data!")

    def Show_All_Operations(self):
        self.connectToDb()
        self.cur.execute('''
            SELECT book_name, client, op_type, op_date, to_time FROM dayoperations
        ''')

        data = self.cur.fetchall()
        #print(data)
        self.tableWidget.setRowCount(0)
        self.tableWidget.insertRow(0)
        for row, form in enumerate(data):
            for column, item in enumerate(form):
                # print(row)
                # print(column)
                # print(item)
                # print("########")
                self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                column = column+1
            row_position = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row_position)

##########################################################
####################### Books ############################

    def Add_New_Book(self):
        self.connectToDb()

        book_title = self.lineEdit_2.text()
        book_description = self.textEdit.toPlainText()
        book_code = self.lineEdit_3.text()
        book_category = self.comboBox_3.currentText()
        book_author = self.comboBox_4.currentText()
        book_publisher = self.comboBox_5.currentText()
        book_price = self.lineEdit_4.text()

        self.cur.execute('''
            INSERT INTO book(book_name, book_description, book_code, book_category, book_price,
            book_author, book_publisher)
            VALUES(%s, %s, %s, %s, %s, %s, %s)
        ''', (book_title, book_description, book_code, book_category, book_price, book_author, book_publisher))

        self.db.commit()
        self.statusBar().showMessage('New Book Added!')
        self.lineEdit_2.setText('')
        self.textEdit.setPlainText('')
        self.lineEdit_3.setText('')
        self.lineEdit_4.setText('')
        self.comboBox_3.setCurrentIndex(0)
        self.comboBox_4.setCurrentIndex(0)
        self.comboBox_5.setCurrentIndex(0)
        self.Show_All_Books()

    def Search_Books(self):
        self.connectToDb()

        book_title = self.lineEdit_8.text()
        sql = '''SELECT * FROM  book WHERE book_name = %s'''
        self.cur.execute(sql, [book_title])
        data = self.cur.fetchone()
        #print(data)
        self.lineEdit_7.setText(data[1])  #name
        self.textEdit_2.setPlainText(data[2])  #description
        self.lineEdit_6.setText(data[3])  #code
        self.comboBox_8.setCurrentText(data[4])  #category
        self.comboBox_6.setCurrentText(data[6])  #author
        self.comboBox_7.setCurrentText(data[7])  #publisher
        self.lineEdit_5.setText(str(data[5]))   #price

    def Edit_Books(self):
        self.connectToDb()

        book_title = self.lineEdit_7.text()
        book_description = self.textEdit_2.toPlainText()
        book_code = self.lineEdit_6.text()
        book_category = self.comboBox_8.currentText()
        book_author = self.comboBox_6.currentText()
        book_publisher = self.comboBox_7.currentText()
        book_price = self.lineEdit_5.text()
        search_book_title = self.lineEdit_8.text()

        self.cur.execute('''
            UPDATE book SET book_name=%s, book_description=%s, book_code=%s, book_category=%s,
            book_price=%s, book_author=%s, book_publisher=%s WHERE book_name=%s
        ''', (book_title, book_description, book_code, book_category, book_price, book_author,
              book_publisher, search_book_title))
        self.db.commit()
        self.Show_All_Books()
        self.statusBar().showMessage('Book Updated!')

    def Delete_Books(self):
        self.connectToDb()

        search_book_title = self.lineEdit_8.text()

        warning = QMessageBox.warning(self, 'Delete Book', 'are you sure you want to delete this book?', QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes:
            sql = '''DELETE FROM book WHERE book_name=%s'''
            self.cur.execute(sql, [search_book_title])
            self.db.commit()
            self.Show_All_Books()
            self.statusBar().showMessage('Book Deleted!')

    def Show_All_Books(self):
        self.connectToDb()

        self.cur.execute('''SELECT book_code, book_name, book_description, book_category, book_author,
         book_publisher, book_price FROM book''')
        data = self.cur.fetchall()
        self.tableWidget_5.setRowCount(0)
        self.tableWidget_5.insertRow(0)
        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_5.setItem(row, column, QTableWidgetItem(str(item)))
                column = column + 1
            row_position = self.tableWidget_5.rowCount()
            self.tableWidget_5.insertRow(row_position)

##########################################################
####################### Clients ############################

    def Add_New_Client(self):
        self.connectToDb()

        client_name = self.lineEdit_29.text()
        client_email = self.lineEdit_30.text()
        client_national_id = self.lineEdit_28.text()

        self.cur.execute('''
            INSERT INTO clients(client_name, client_email, client_national_id)
            VALUES (%s, %s, %s)
        ''', (client_name, client_email, client_national_id))
        self.db.commit()
        self.db.close()
        self.Show_All_Clients()
        self.statusBar().showMessage('New Client Added!')

    def Search_Client(self):
        self.connectToDb()

        client_national_id = self.lineEdit_31.text()

        sql = '''SELECT * FROM clients WHERE client_national_id=%s'''
        self.cur.execute(sql, [client_national_id])
        data = self.cur.fetchone()
        self.lineEdit_26.setText(data[1])
        self.lineEdit_27.setText(data[2])
        self.lineEdit_25.setText(data[3])

    def Edit_Clients(self):
        self.connectToDb()

        client_original_national_id = self.lineEdit_31.text()
        client_name = self.lineEdit_26.text()
        client_email = self.lineEdit_27.text()
        client_national_id = self.lineEdit_25.text()

        self.cur.execute('''
            UPDATE clients SET client_name=%s, client_email=%s, client_national_id=%s WHERE client_national_id=%s
        ''', (client_name, client_email, client_national_id, client_original_national_id))
        self.db.commit()
        self.db.close()
        self.Show_All_Clients()
        self.statusBar().showMessage('Client Updated!')

    def Delete_Client(self):
        self.connectToDb()

        client_national_id = self.lineEdit_31.text()

        warning = QMessageBox.warning(self, 'Delete Client', 'are you sure you want to delete this client?',
                                      QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes:
            sql = '''DELETE FROM  clients WHERE client_national_id=%s'''
            self.cur.execute(sql, [client_national_id])
            self.db.commit()
            self.db.close()
            self.Show_All_Clients()
            self.statusBar().showMessage('Client Deleted!')

    def Show_All_Clients(self):
        self.connectToDb()

        self.cur.execute('''SELECT client_name, client_email, client_national_id FROM clients''')
        data = self.cur.fetchall()
        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)
        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_6.setItem(row, column, QTableWidgetItem(str(item)))
                column = column + 1
            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)

##########################################################
####################### Users ############################

    def Add_New_User(self):
        self.connectToDb()

        username = self.lineEdit_9.text()
        email = self.lineEdit_10.text()
        password = self.lineEdit_11.text()
        repeat_password = self.lineEdit_12.text()
        #print('***************')
        if password == repeat_password:
            self.cur.execute('''
                INSERT INTO users(user_name , user_email , user_password)
                VALUES (%s, %s, %s)
            ''', (username, email, password))
            self.db.commit()
            self.statusBar().showMessage('New User Added!')
        else:
            self.label_30.setText('Please enter same passwords')

    def Login(self):
        self.connectToDb()

        username = self.lineEdit_18.text()
        password = self.lineEdit_17.text()

        sql = '''SELECT * FROM users'''
        self.cur.execute(sql)
        data = self.cur.fetchall()
        #print(data)
        for row in data:
            if username == row[1] and password == row[3]:
                self.statusBar().showMessage('Login was successful!')
                self.groupBox_4.setEnabled(True)
                self.lineEdit_15.setText(row[1])
                self.lineEdit_16.setText(row[2])
                self.lineEdit_14.setText(row[3])

    def Edit_User(self):
        self.connectToDb()

        username = self.lineEdit_15.text()
        email = self.lineEdit_16.text()
        password = self.lineEdit_14.text()
        repeat_password = self.lineEdit_13.text()
        original_username = self.lineEdit_18.text()

        if password == repeat_password:
            self.cur.execute('''
                UPDATE users SET user_name=%s, user_email=%s, user_password=%s WHERE user_name=%s
            ''', (username, email, password, original_username))
            self.db.commit()
            self.statusBar().showMessage('User Data Update!')
        else:
            self.label_31.setText('Please enter same passwords')

##########################################################
####################### Settings #########################

    def Add_Category(self):
        self.connectToDb()

        category_name = self.lineEdit_19.text()

        self.cur.execute('''
            INSERT INTO category(category_name) VALUES (%s)
        ''', (category_name,))
        self.db.commit()
       #print('category added!')
        self.statusBar().showMessage('New Category Added!')
        #self.lineEdit_19.text('')
        self.Show_Category()
        self.Show_Category_Combobox()

    def Show_Category(self):
        self.connectToDb()

        self.cur.execute(''' SELECT category_name FROM category''')
        data = self.cur.fetchall()
        #print(data)
        if data:
            self.tableWidget_2.setRowCount(0)
            self.tableWidget_2.insertRow(0)   #insert new row
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_2.setItem(row, column, QTableWidgetItem(str(item)))
                    column +=1
                row_position = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(row_position)

    def Add_Author(self):
        self.connectToDb()

        author_name = self.lineEdit_20.text()

        self.cur.execute('''
                    INSERT INTO authors(author_name) VALUES (%s)
                ''', (author_name,))
        self.db.commit()
        # print('author added!')
        self.statusBar().showMessage('New Author Added!')
        #self.lineEdit_20('')
        self.Show_Author()
        self.Show_Author_Combobox()

    def Show_Author(self):
        self.connectToDb()

        self.cur.execute(''' SELECT author_name FROM authors''')
        data = self.cur.fetchall()
        #print(data)
        if data:
            self.tableWidget_3.setRowCount(0)
            self.tableWidget_3.insertRow(0)  # insert new row
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_3.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
                row_position = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_position)

    def Add_Publisher(self):
        self.connectToDb()

        publisher_name = self.lineEdit_21.text()

        self.cur.execute('''
                            INSERT INTO publisher(publisher_name) VALUES (%s)
                        ''', (publisher_name,))
        self.db.commit()
        # print('publisher added!')
        self.statusBar().showMessage('New Publisher Added!')
        #self.lineEdit_21('')
        self.Show_Publisher()
        self.Show_Publisher_Combobox()

    def Show_Publisher(self):
        self.connectToDb()

        self.cur.execute(''' SELECT publisher_name FROM publisher''')
        data = self.cur.fetchall()
        #print(data)
        if data:
            self.tableWidget_4.setRowCount(0)
            self.tableWidget_4.insertRow(0)  # insert new row
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_4.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
                row_position = self.tableWidget_4.rowCount()
                self.tableWidget_4.insertRow(row_position)

##########################################################
############# Show Settings Data in UI ###################

    def Show_Category_Combobox(self):
        self.connectToDb()

        self.cur.execute('''SELECT category_name FROM category''')
        data = self.cur.fetchall()
        self.comboBox_3.clear()
        for category in data:
            self.comboBox_3.addItem(category[0])
            self.comboBox_8.addItem(category[0])

    def Show_Author_Combobox(self):
        self.connectToDb()

        self.cur.execute('''SELECT author_name FROM authors''')
        data = self.cur.fetchall()
        self.comboBox_4.clear()
        for author in data:
            self.comboBox_4.addItem(author[0])
            self.comboBox_6.addItem(author[0])

    def Show_Publisher_Combobox(self):
        self.connectToDb()

        self.cur.execute('''SELECT publisher_name FROM publisher''')
        data = self.cur.fetchall()
        self.comboBox_5.clear()
        for publisher in data:
            self.comboBox_5.addItem(publisher[0])
            self.comboBox_7.addItem(publisher[0])

##########################################################
####################### UI Themes ########################

    def Dark_Blue(self):
        style = open('themes/darkblue.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Dark_Orange(self):
        style = open('themes/darkorange.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Dark_Gray(self):
        style = open('themes/darkgary.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Classic(self):
        style = open('themes/classic.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

##########################################################
####################### Export Data ######################

    def Export_Day_Operations(self):
        self.connectToDb()

        self.cur.execute('''
                    SELECT book_name, client, op_type, op_date, to_time FROM dayoperations
                ''')
        data = self.cur.fetchall()
        wb = Workbook('day_operations.xlsx')
        sheet1 = wb.add_worksheet()
        sheet1.write(0, 0, 'book title')
        sheet1.write(0, 1, 'client name')
        sheet1.write(0, 2, 'type')
        sheet1.write(0, 3, 'from')
        sheet1.write(0, 4, 'to')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        self.statusBar().showMessage('Report Created Successfully!')

    def Export_Books(self):
        self.connectToDb()

        self.cur.execute('''SELECT book_code, book_name, book_description, book_category, book_author,
                 book_publisher, book_price FROM book''')
        data = self.cur.fetchall()

        wb = Workbook('all_books.xlsx')
        sheet1 = wb.add_worksheet()
        sheet1.write(0, 0, 'book code')
        sheet1.write(0, 1, 'book name')
        sheet1.write(0, 2, 'book description')
        sheet1.write(0, 3, 'book category')
        sheet1.write(0, 4, 'book author')
        sheet1.write(0, 5, 'book publisher')
        sheet1.write(0, 6, 'book price')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        self.statusBar().showMessage('Report Created Successfully!')

    def Export_Clients(self):
        self.connectToDb()

        self.cur.execute('''SELECT client_name, client_email, client_national_id FROM clients''')
        data = self.cur.fetchall()

        wb = Workbook('all_clients.xlsx')
        sheet1 = wb.add_worksheet()
        sheet1.write(0, 0, 'client name')
        sheet1.write(0, 1, 'client email')
        sheet1.write(0, 2, 'client national id')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        self.statusBar().showMessage('Report Created Successfully!')


def main():
    app = QApplication(sys.argv)
    window = Login()
    window.show()
    app.exec_()


if __name__ == "__main__":
    main()

