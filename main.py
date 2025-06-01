
import sys
from PyQt6 import QtWidgets, uic
import pyodbc
from PyQt6.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QHeaderView
import sqlite3
from PyQt6.QtWidgets import QMessageBox

server = 'DESKTOP-1RI4AT8\AHMEDSERVER'
database = 'DATABASECREATION'  # Name of your Northwind database
use_windows_authentication =True  # Set to True to use Windows Authentication
username = 'sa'  # Specify a username if not using Windows Authentication
password = 'Fall2022.dbms'  # Specify a password if not using Windows Authentication
if use_windows_authentication:
    connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;'
else:
    connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
connection = pyodbc.connect(connection_string)
cursor = connection.cursor()
savedid=0
users = [
    {"id": "1", "password": "pass123", "role": "Employee"},
    {"id": "2", "password": "emp234", "role": "Employee"},
    {"id": "admin1", "password": "admin123", "role": "Office Management"},
    {"id": "admin2", "password": "admin234", "role": "Office Management"},
]
   

    #popups
def show_attendance_marked():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Icon.Information)
    msg.setWindowTitle("Success")
    msg.setText("Attendance marked successfully!")
    msg.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg.exec()

def show_submitted():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Icon.Information)
    msg.setWindowTitle("Success")
    msg.setText("Submission successful!")
    msg.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg.exec()

def show_record_deleted():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Icon.Warning)
    msg.setWindowTitle("Record Deleted")
    msg.setText("The record has been deleted.")
    msg.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg.exec()







class MainApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()


        self.selected_role = None
        self.open_login_screen()

    def open_login_screen(self):

        self.login_window = uic.loadUi('login.ui')
        self.login_window.show()

        self.login_window.loginButton.clicked.connect(self.check_login)

    def check_login(self):
        user_id = self.login_window.lineEdit.text()  # Replace with actual field names in login.ui
        password = self.login_window.lineEdit_2.text()

        # Verify credentials from the list
        user = next((u for u in users if u["id"] == user_id and u["password"] == password), None)
        if user:
            self.login_window.close()
            if user["role"] == "Employee":
                self.open_employee_view(user["id"])
            elif user["role"] == "Office Management":
                self.open_office_management()

    def open_employee_view(self,employee_id):
        self.employee_window = uic.loadUi('Employee_view.ui')
        self.employee_window.show()
        self.employee_window.pushButton.clicked.connect(self.open_absence_request_form)
        self.employee_window.pushButton_2.clicked.connect(lambda:self.attendancerecord(employee_id))
        self.populate_employee_details(employee_id)

    def open_office_management(self):

        self.office_management_window = uic.loadUi('office_view.ui')
        self.office_management_window.setWindowTitle("Office Management View")
        self.office_management_window.show()
        self.office_management_window.generatePaySlipButton.clicked.connect(self.payslip_opener)
        self.office_management_window.viewAttendanceButton.clicked.connect(self.open_attendanceform)
        self.office_management_window.deleteButton.clicked.connect(self.deleterecord)
        self.office_management_window.tableWidget1.itemClicked.connect(self.savedata)
        self.office_management_window.pushButton.clicked.connect(self.addsalary)
        self.office_management_window.editButton.clicked.connect(self.editer)

       
        self.officeviewpopulate()

    def savedata(self,item):
        global savedid
        savedid=self.office_management_window.tableWidget1.item(item.row(), 0).text()
        print(savedid)

    def deleterecord(self):
        connection = pyodbc.connect(
        connection_string)
        cursor = connection.cursor()
        query = """
                DECLARE @Employee_ID INT = ?;
                
                -- Delete from Payroll table
                DELETE FROM Payroll
                WHERE Salary_ID IN (SELECT Salary_ID FROM Salary WHERE Employee_ID = @Employee_ID);

                -- Delete from Salary table
                DELETE FROM Salary
                WHERE Employee_ID = @Employee_ID;

                -- Delete from Attendance table
                DELETE FROM Attendance
                WHERE Employee_ID = @Employee_ID;

                -- Delete from Employee_Department table
                DELETE FROM Employee_Department
                WHERE Employee_ID = @Employee_ID;

                -- Finally, delete from Employees table
                DELETE FROM Employees
                WHERE Employee_ID = @Employee_ID;
            """

            # Execute the query with the employee ID as a parameter
        print(savedid)
        cursor.execute(query, savedid)
        connection.commit()
        show_record_deleted()
    def addsalary(self):
        self.salaryadder = uic.loadUi('addnewsalary.ui')
        self.salaryadder.show()
        self.salaryadder.submitButton.clicked.connect(self.addsalaryvalues)
       


    def addsalaryvalues(self):
        connection = pyodbc.connect(
        connection_string)
        cursor = connection.cursor()
        employeeid=self.salaryadder.textInput1.text()
        amount=self.salaryadder.textInput2.text()
        tax=self.salaryadder.textInput3.text()
        bonus=self.salaryadder.textInput4.text()
        startdate=self.salaryadder.dateEdit.date().toString("yyyy-MM-dd")
        enddate=self.salaryadder.dateEdit_2.date().toString("yyyy-MM-dd")
        cursor.execute("SELECT ISNULL(MAX(Salary_ID), 0) FROM Salary")
        max_salary_id = cursor.fetchone()[0]
        new_salary_id = max_salary_id + 1
        query = '''
            INSERT INTO Salary (Salary_ID, Employee_ID, Amount,Bonus_Amount,Tax_Amount, Start_Date,End_Date)
            VALUES (?, ?, ?, ?,?,?,?)
            '''

            # Execute query with parameters
        cursor.execute(query, (int(new_salary_id), employeeid ,amount, tax,bonus,str(startdate),str(enddate)))
        connection.commit()
        show_submitted()

    def editer(self):
        self.editerwindow = uic.loadUi('edit.ui')
        self.editerwindow.show()
        connection = pyodbc.connect(
        connection_string)
        cursor = connection.cursor()
        fetch_query = """
                    SELECT 
                        e.First_Name,
                        e.Last_Name,
                        e.Hire_Date,
                        e.Gender,
                        e.Age,
                        e.phone_number,
                        e.email,
                        e.address,
                        d.Department_Name
                    FROM 
                        Employees e
                    JOIN 
                        Employee_Department ed ON e.Employee_ID = ed.Employee_ID
                    JOIN 
                        Departments d ON ed.Department_ID = d.Department_ID
                    WHERE 
                        e.Employee_ID = ?;
                """
        cursor.execute(fetch_query, (savedid))
        result = cursor.fetchone()
        
                
        if result:
# Populate fields
            self.editerwindow.line_edit_first_name.setText(result[0])
            self.editerwindow.line_edit_last_name.setText(result[1])
            self.editerwindow.line_edit_hire_date.setText(result[2].strftime("%Y-%m-%d"))
            self.editerwindow.line_edit_gender.setText(result[3])
            self.editerwindow.line_edit_age.setText(str(result[4]))
            self.editerwindow.line_edit_phone_number.setText(str(result[5]))
            self.editerwindow.line_edit_email.setText(result[6])
            self.editerwindow.line_edit_address.setText(result[7])
            self.editerwindow.line_edit_department_name.setText(result[8])       
        connection.close()
        self.editerwindow.button_update.clicked.connect(self.updatevalues)
        
    def updatevalues(self):
        connection = pyodbc.connect(
        connection_string)
        cursor = connection.cursor()
        first_name=self.editerwindow.line_edit_first_name.text()
        last_name=self.editerwindow.line_edit_last_name.text()
        hire=self.editerwindow.line_edit_hire_date.text()
        gender=self.editerwindow.line_edit_gender.text()
        age= self.editerwindow.line_edit_age.text()
        phone_number=self.editerwindow.line_edit_phone_number.text()
        email=self.editerwindow.line_edit_email.text()
        adress=self.editerwindow.line_edit_address.text()
        department=self.editerwindow.line_edit_department_name.text()
        update_employee_query = """
                UPDATE Employees
                SET 
                    First_Name = ?,
                    Last_Name = ?,
                    Hire_Date = ?,
                    Gender = ?,
                    Age = ?,
                    phone_number = ?,
                    email = ?,
                    address = ?
                WHERE Employee_ID = ?;
            """
        cursor.execute(update_employee_query, (first_name, last_name, hire, gender, age, phone_number, email, adress, savedid))
        fetch_department_id_query = """
                SELECT Department_ID FROM Departments WHERE Department_Name = ?;
            """
        cursor.execute(fetch_department_id_query, (department,))
        result = cursor.fetchone()
        department_id = result[0]
        update_department_query = """
                    UPDATE Employee_Department
                    SET Department_ID = ?
                    WHERE Employee_ID = ?;
                """
        cursor.execute(update_department_query, (department_id, savedid))
        connection.commit()
        show_submitted()
        connection.close()





    def payslip_opener(self):
        self.absence_request_window = uic.loadUi('payslip.ui')
        self.absence_request_window.show()
        selected_year = self.office_management_window.dateEdit.date().year()
        selected_month = self.office_management_window.dateEdit.date().month()
        connection = pyodbc.connect(
        connection_string)
        cursor = connection.cursor()
        query = """
    DECLARE @selectedYear INT = ?;
    DECLARE @selectedMonth INT = ?;

    INSERT INTO Payroll (Payroll_ID, Salary_ID, attendance, tax_deducted, bonus_added, Total_Amount, Date)
    SELECT 
        ROW_NUMBER() OVER (ORDER BY s.Salary_ID) + (SELECT ISNULL(MAX(Payroll_ID), 0) FROM Payroll) AS Payroll_ID,  -- Generate custom Payroll_ID
        s.Salary_ID,
        (COUNT(a.Attendance_ID) * 100.0 / CAST(DAY(EOMONTH(DATEFROMPARTS(@selectedYear, @selectedMonth, 1))) AS FLOAT)) AS attendance_percentage,
        s.Tax_Amount AS tax_deducted,
        s.Bonus_Amount AS bonus_added,
        s.Amount,
        DATEFROMPARTS(@selectedYear, @selectedMonth, 1) AS Date
    FROM Salary s
    LEFT JOIN Attendance a 
        ON s.Employee_ID = a.Employee_ID 
        AND MONTH(a.Date) = @selectedMonth 
        AND YEAR(a.Date) = @selectedYear
        AND a.Status = 'Present'
    WHERE NOT EXISTS (
        SELECT 1 FROM Payroll p
        WHERE p.Salary_ID = s.Salary_ID 
        AND MONTH(p.Date) = @selectedMonth 
        AND YEAR(p.Date) = @selectedYear
    )
    GROUP BY s.Salary_ID, s.Amount, s.Tax_Amount, s.Bonus_Amount;
"""

        params = (selected_year, selected_month)
        cursor.execute(query, params)
        connection.commit()
        query2 = """
    DECLARE @selectedYear INT = ?;
    DECLARE @selectedMonth INT = ?;

    SELECT 
        Payroll_ID,
        Salary_ID,
        attendance,
        tax_deducted,
        bonus_added,
        Total_Amount,
        Date
    FROM 
        Payroll
    WHERE 
        MONTH(Date) = @selectedMonth 
        AND YEAR(Date) = @selectedYear;
"""

        params2 = (selected_year, selected_month)
        cursor.execute(query2, params2)
        for row_index, row_data in enumerate(cursor.fetchall()):
            self.absence_request_window.tableWidget.insertRow(row_index)
            for col_index, cell_data in enumerate(row_data):
                item = QTableWidgetItem(str(cell_data))
                self.absence_request_window.tableWidget.setItem(row_index, col_index, item)


    def open_absence_request_form(self):

        self.absenseform = uic.loadUi('absenceRequestForm.ui')
        self.absenseform.show()
        self.absenseform.pushButton.clicked.connect(show_submitted)
  
    def open_attendanceform(self):
        self.attendanceform = uic.loadUi('attenance_mark.ui')
        self.attendanceform.show()
        self.attendanceform.button_markAttendance.clicked.connect(self.savingattendance)
    def savingattendance(self):
        text = self.attendanceform.lineEdit.text()
        selected_date = self.attendanceform.dateEdit.date().toString('yyyy-MM-dd')
        selected_item = self.attendanceform.comboBox.currentText()
        print(text,selected_date,selected_item)
        print("water")
        connection = pyodbc.connect(
        connection_string)
        cursor = connection.cursor()

        query_get_max_id = "SELECT ISNULL(MAX(Attendance_ID), 0) FROM Attendance;"
        cursor.execute(query_get_max_id)
        max_id = cursor.fetchone()[0]  # Get the result
        new_attendance_id = max_id + 1  # Increment the ID

        query = "INSERT INTO Attendance (Attendance_ID,Employee_ID, Date, Status) VALUES (?,?, ?,?);"
        cursor.execute(query,(new_attendance_id,text,selected_date,selected_item))
        connection.commit()
        connection.close()
        show_attendance_marked()



    def populate_employee_details(self, employee_id):
        connection = pyodbc.connect(
            connection_string
        )
        cursor = connection.cursor()
    # Fetch employee details for a specific ID
        query = "SELECT Employee_ID, First_Name,last_Name, Hire_Date,Gender, Age,phone_number,email,address FROM employees WHERE Employee_ID = ?"
        cursor.execute(query, (employee_id,))
        result = cursor.fetchone()
        query2 = "SELECT e.Employee_ID, e.First_Name, e.Last_Name,(CAST(SUM(CASE WHEN a.Status = 'Present' THEN 1 ELSE 0 END) AS FLOAT) /COUNT(a.Date)) * 100 AS Attendance_Percentage FROM  Employees e JOIN  Attendance a ON e.Employee_ID = a.Employee_ID WHERE  e.Employee_ID = ? GROUP BY  e.Employee_ID, e.First_Name, e.Last_Name;"
        cursor.execute(query2, (employee_id,))
        result2 = cursor.fetchone()
        query3 = "SELECT  e.Employee_ID, e.First_Name, e.Last_Name, d.Department_Name, d.Location,d.Description FROM  Employees e JOIN Employee_Department ed ON e.Employee_ID = ed.Employee_ID JOIN  Departments d ON ed.Department_ID = d.Department_ID WHERE  e.Employee_ID = ?;"
        cursor.execute(query3, (employee_id,))
        result3 = cursor.fetchone()
        query4 = "SELECT TOP 1 Amount  FROM Salary  WHERE Employee_ID = ?  ORDER BY Start_Date DESC;"
        cursor.execute(query4, (employee_id,))
        result4 = cursor.fetchone()
        
        


        if result:
        # Assuming your QTextBrowsers are named appropriately in the .ui file
            self.employee_window.textBrowser_2.setText(result[1])  # Position
            self.employee_window.textBrowser_6.setText(result[2])  # Department
            self.employee_window.textBrowser_3.setText(str(result[7]))  # Email
            self.employee_window.textBrowser_7.setText(str(result[6]))
            self.employee_window.textBrowser_8.setText(str(result[8]))
            self.employee_window.lineEdit.setText(str(result[5]))
            self.employee_window.dateEdit.setDate(result[3])
            if result[4]=='Female':
                self.employee_window.radioButton_2.setChecked(True)
            elif result[4]=='Male':
                self.employee_window.radioButton.setChecked(True)
            self.employee_window.textBrowser_9.setText(str(result2[3])) 
            self.employee_window.textBrowser_4.setText(result3[3])
            self.employee_window.textBrowser_5.setText(str(result4[0]))

        connection.close()

    def officeviewpopulate(self):
        connection = pyodbc.connect(
            connection_string
        )
        cursor = connection.cursor()
    # Fetch employee details for a specific ID
        query = "SELECT  e.Employee_ID,  e.First_Name,  e.Last_Name,  d.Department_Name,  s.Amount AS Salary FROM Employees e LEFT JOIN  Employee_Department ed ON e.Employee_ID = ed.Employee_ID LEFT JOIN  Departments d ON ed.Department_ID = d.Department_ID LEFT JOIN  Salary s ON e.Employee_ID = s.Employee_ID AND s.Start_Date = ( SELECT MAX(Start_Date)  FROM Salary WHERE Employee_ID = e.Employee_ID);"
        cursor.execute(query)
        for row_index, row_data in enumerate(cursor.fetchall()):
            self.office_management_window.tableWidget1.insertRow(row_index)
            for col_index, cell_data in enumerate(row_data):
                item = QTableWidgetItem(str(cell_data))
                self.office_management_window.tableWidget1.setItem(row_index, col_index, item)

        # Close the database connection
        connection.close()

    def attendancerecord(self,employee_id):
        self.attendancerecord = uic.loadUi('attendancerecord.ui')
        self.attendancerecord.show()
        connection = pyodbc.connect(connection_string)
        cursor = connection.cursor()
        query = """
    SELECT 
        Date, 
        Status
    FROM 
        Attendance
    WHERE 
        Employee_ID = ? 
    ORDER BY 
        Date DESC;
    """
        cursor.execute(query, (employee_id,))
        results = cursor.fetchall()
        self.attendancerecord.tableWidget.setRowCount(0)

        for row_index, row_data in enumerate(results):
            self.attendancerecord.tableWidget.insertRow(row_index)  # Insert a new row for each attendance record
            for col_index, data in enumerate(row_data):
                item = QtWidgets.QTableWidgetItem(str(data))  # Convert each data item to a string
                self.attendancerecord.tableWidget.setItem(row_index, col_index, item)

    # Close the database connection
    connection.close()







def main():
    app = QtWidgets.QApplication(sys.argv)
    main_app = MainApp()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
