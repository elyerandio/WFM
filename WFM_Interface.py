#***********************************
# Program Name : WFM_Interface.py
# Date Written : October 08, 2013
# Description  : a program to interface the employees schedules in WFM into Orisoft
# Author       : Eleazer L. Erandio
#************************************

import sys
import os
from PySide.QtCore import *
from PySide.QtGui import *
import pyodbc
from datetime import *
import re
import ConfigParser
from WFMReport import *


class WFMInterface(QMainWindow):
    def __init__(self, parent=None):
        super(WFMInterface, self).__init__(parent)
        self.resize(200,100)
        form = WFMInterfaceForm()
        self.setCentralWidget(form)
        self.setWindowTitle('WFM Interface')
        #form.show()

class WFMInterfaceForm(QDialog):

    def __init__(self, parent=None):
        global dateFromPrev, dateToPrev

        # get the to-date of last process, then add 1 day to get the next process date
        (pyear, pmonth, pday) = dateToPrev.split('-')
        nextDate = date(int(pyear), int(pmonth), int(pday))
        nextDate = nextDate + timedelta(1)
        super(WFMInterfaceForm, self).__init__(parent)

        fromLabel = QLabel('From')
        fromLabel.setAlignment(Qt.AlignHCenter)
        toLabel = QLabel('To')
        toLabel.setAlignment(Qt.AlignHCenter)
        dateLabel = QLabel('Date')
        dateLabel.setAlignment(Qt.AlignRight)
        self.dateEditFrom = QDateTimeEdit(date(nextDate.year, nextDate.month, nextDate.day))
        self.dateEditFrom.setCalendarPopup(True)
        self.dateEditFrom.setMaximumDate(QDate.currentDate())
        self.dateEditTo = QDateTimeEdit(date(nextDate.year, nextDate.month, nextDate.day))
        self.dateEditTo.setCalendarPopup(True)
        self.labelStatus = QLabel('')
        self.labelSaveAs = QLabel('Save As')
        self.chOverWrite = QCheckBox('Overwrite?')
        self.chOverWrite.setChecked(True)

        self.processButton = QPushButton('Process')
        self.cancelButton = QPushButton('Cancel')
        buttonLayout = QHBoxLayout()
        buttonLayout.addSpacing(2)
        buttonLayout.addWidget(self.processButton)
        buttonLayout.addWidget(self.cancelButton)

        layout = QGridLayout()
        layout.addWidget(fromLabel, 0, 1)
        layout.addWidget(toLabel, 0, 2)
        layout.addWidget(dateLabel, 1, 0)
        layout.addWidget(self.dateEditFrom, 1, 1)
        layout.addWidget(self.dateEditTo, 1, 2)
        layout.addWidget(self.chOverWrite, 2, 1)
        layout.addLayout(buttonLayout, 3, 1, 1, 2)
        layout.addWidget(self.labelStatus, 4, 0, 1, 3)

        self.dateEditFrom.dateChanged.connect(self.setDateTo)
        self.dateEditTo.dateChanged.connect(self.setDateFrom)
        self.connect(self.processButton, SIGNAL('clicked()'), self.process)
        self.connect(self.cancelButton, SIGNAL('clicked()'), self.canceled)

        self.setLayout(layout)
        self.setWindowTitle('WFM Inteface')


    def setDateFrom(self):
        # don't allow dateEditTo.date to be less than dateEditFrom.date
        # set dateEditFrom.date to dateEditTo.date
        if self.dateEditTo.date() < self.dateEditFrom.date():
            self.dateEditFrom.setDate(self.dateEditTo.date())

        if self.dateEditTo.date() > QDate.currentDate():
            self.dateEditTo.setDate(QDate.currentDate())

    def setDateTo(self):
        # set dateEditTo.date to dateEditFrom.date
        self.dateEditTo.setDate(self.dateEditFrom.date())

        if self.dateEditFrom.date() > QDate.currentDate():
            self.dateEditFrom.setDate(QDate.currentDate())


    def process(self):

        # show the confirmation message
        flags = QMessageBox.StandardButton.Yes
        flags |= QMessageBox.StandardButton.No
        question = "Do you really want to process right now?"
        response = QMessageBox.question(self, "Confirm Process", question, flags, QMessageBox.No)
        if response == QMessageBox.No:
            return

        cur = connOriTMS.cursor()
        # truncate the WFM Exception table
        cur.execute('Truncate table dbo.user_wfm_exception')

        self.getValidSchedTypes()
        self.getDaysRange()
        self.getGroupSchedule()
        self.getActiveEmployees()
        self.getSchedules()
        self.saveSchedules()
        self.labelStatus.setText('Process finished!')
        self.saveIni()
        self.processButton.setEnabled(False)
        self.cancelButton.setText('Exit')
        self.viewExceptionReport()

    def createEmployee(self):

        emp = {}
        emp['lastname'] = ''
        emp['firstname'] = ''
        emp['shift_schedule'] = ''
        emp['restday_schedule'] = ''
        emp['workhours'] = 9

        return emp

    def getValidSchedTypes(self):
        """
        Creates a list of valid schedule types from Orisoft
        """
        cur = connOriTMS.cursor()
        cur.execute("Select schedule_type_code from schedule_type")
        for rec in cur:
            validSchedType.add(rec[0])



    def getDaysRange(self):
        """
        Creates a list of valid dates from DateFrom to DateTo in the entry panel
        """

        global daysRange
        dateFrom = self.dateEditFrom.date()
        dateTo = self.dateEditTo.date()

        self.labelStatus.setText('Getting date range')
        self.repaint()

        # convert QDate format to Python Date
        dateFrom = dateFrom.toPython()
        dateTo = dateTo.toPython()

        currDay = dateFrom
        daysRange = []
        while currDay <= dateTo:
            daysRange.append(currDay.isoformat())
            currDay = currDay + timedelta(1)                # add 1 day to current day



    def getGroupSchedule(self):

        dateFrom = self.dateEditFrom.date()
        dateTo = self.dateEditTo.date()

        # convert QDate format to Python Date
        dateFrom = dateFrom.toPython()
        dateTo = dateTo.toPython()

        self.labelStatus.setText('Getting Group Schedules')
        self.repaint()

        cur = connOriTMS.cursor()
        cur.execute("Select id, work_group, work_period_id from group_schedule_hd " \
            "where year(convert(datetime, work_period_id)) between year('%s') and year('%s') " \
            "and month(convert(datetime, work_period_id)) between month('%s') and month('%s')" % (dateFrom, dateTo, dateFrom, dateTo))
        for rec in cur:
            refer_id = rec[0]
            work_group = rec[1]
            work_period = rec[2]                # format is mm/dd/yyyy

            # convert work_period format to mm/yyyy
            period = work_period.split('/')[0] + '/' + work_period.split('/')[2]

            if work_group not in groupSchedule:
                groupSchedule[work_group] = {}

            groupSchedule[work_group][period] = refer_id

    def getActiveEmployees(self):
        cur = connOriTMS.cursor()
        #cur.execute("Select eb.employee_id, eb.employee_no, eb.employee_name from employee_biodata eb, employee_employment ee" \
        #" where eb.employee_id = ee.employee_id and ee.employee_status = 'A' ")
        cur.execute("Select employee_no, employee_name, work_group_code from employee_badge where employee_status = 'A' ")

        for rec in cur:
            employee_no = rec[0]
            fullname = rec[1]
            workgroup = rec[2]

            if fullname.count(',') > 1:
                (lastname, firstname, whatever) = fullname.split(',', 2)
            elif fullname.count(',') == 1:
                (lastname, firstname) = fullname.split(',')
            else:
                lastname = fullname
                firstname = ''

            employees[employee_no] = self.createEmployee()
            employees[employee_no]['lastname'] = lastname
            employees[employee_no]['firstname'] = firstname
            employees[employee_no]['workgroup'] = workgroup
            employees[employee_no]['workhours'] = 9


    def getSchedules(self):

        dateStart = self.dateEditFrom.date()
        dateEnd = self.dateEditTo.date()

        self.labelStatus.setText('Getting schedules from WFM Database.')
        self.repaint()

        cur = connWFM.cursor()
        query = "select rs.payroll, rs.rdate, r.shift, rs.start, rs.finish, rs.hours from roster r join roster_staff rs on r.[key] = rs.roster_key" \
            " where r.start between '%s' and '%s' order by payroll, rdate" % (dateStart.toPython(), dateEnd.toPython())
        cur.execute(query)
        #cur.execute("select rs.payroll, rs.rdate, r.shift, rs.start, rs.finish from roster r join roster_staff rs on r.[key] = rs.roster_key" \
        #    " where r.start between '%s' and '%s' order by payroll, rdate" % (dateStart.toPython(), dateEnd.toPython()))
        for rec in cur:
            emp = rec[0]
            rdate = rec[1].date()
            shift = rec[2]
            time_start = rec[3].time().isoformat()[:2]
            time_end = rec[4].time().isoformat()[:2]
            hours = rec[5]

            if emp in employees:
                if 'sched' not in employees[emp]:
                    employees[emp]['sched'] = {}

                employees[emp]['sched'][rdate.isoformat()] = time_start + time_end
                employees[emp]['workhours'] = hours

            self.labelStatus.setText('Fetching schedule of employee : ' + str(emp) + ' from WFM database.')
            self.repaint()


    # Saves the employees daily schedules to table Employee_Schedule in Orisoft DB
    def saveSchedules(self):

        # get the next new record ID of table Employee_Schedule
        cur = connOriTMS.cursor()
        query = "select ctrlctr from ofcctrlid where ctrlcol = 'employee_schedule'"
        cur.execute(query)
        saveID = cur.fetchone()[0]
        currID = saveID

        self.labelStatus.setText('Saving schedules to Orisoft TMS.')
        self.repaint()

        # loop through the employees daily schedule hash
        for emp in sorted(employees):
            if 'sched' in employees[emp]:
                workgroup = employees[emp]['workgroup']
                fullname = employees[emp]['lastname'] + ', ' + employees[emp]['firstname']
                currDay = self.dateEditFrom.date().toPython()
                schedType = employees[emp]['shift_schedule']
                now = datetime.now()
                if workgroup not in groupSchedule:
                    # No WorkGroup schedule
                    # save exemption record to user_wfm_exception table for reporting purpose
                    query = "Insert into user_wfm_exception(EMPLOYEE_NO, EMPLOYEE_NAME, SCHEDULE_DATE, SCHEDULE_TYPE, WORK_GROUP, REMARKS, CREATED_BY, CREATED_DATE) " \
                    "VALUES('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % (emp, fullname, daysRange[0], schedType, workgroup, 'No workgroup schedule', 'WFM_IFACE', now.strftime('%Y-%m-%d %H:%M:%S'))
                    cur.execute(query)
                    continue

                for currDay in daysRange:
                    if currDay in employees[emp]['sched']:
                        schedType = employees[emp]['sched'][currDay]
                    else:
                        # no schedule for current day, assume rest day
                        if employees[emp]['workhours'] == 9:
                            schedType = 'RD08'
                        elif employees[emp]['workhours'] == 12:
                            schedType = 'RD11'
                        else:
                            schedType = 'REST'

                    # verify schedType if valid
                    if schedType not in validSchedType:
                        # save exemption record to user_wfm_exception table for reporting purpose
                        query = "Insert into user_wfm_exception(EMPLOYEE_NO, EMPLOYEE_NAME, SCHEDULE_DATE, SCHEDULE_TYPE, WORK_GROUP, REMARKS, CREATED_BY, CREATED_DATE) " \
                        "VALUES('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % (emp, fullname, currDay, schedType, workgroup, 'ScheduleType is invalid.', 'WFM_IFACE', now.strftime('%Y-%m-%d %H:%M:%S'))
                        cur.execute(query)
                        continue

                    # get REFER_ID from groupSchedule hash
                    (year,month,day) = currDay.split('-')
                    key = '/'.join([month, year])
                    referID = groupSchedule[workgroup].get(key, 'NULL')
                    now = datetime.now()
                    if referID == 'NULL':
                        # No WorkGroup schedule
                        # save exemption record to user_wfm_exception table for reporting purpose
                        cur.execute("Insert into user_wfm_exception(EMPLOYEE_NO, EMPLOYEE_NAME, SCHEDULE_DATE, SCHEDULE_TYPE, WORK_GROUP, REMARKS, CREATED_BY, CREATED_DATE) " \
                         "VALUES('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % (emp, fullname, currDay, schedType, workgroup, 'No workgroup schedule', 'WFM_IFACE', now.strftime('%Y-%m-%d %H:%M:%S')))
                        continue

                    # repeatCount variable is used to avert an infinite loop, the while loop will exit if the value is more than 1
                    repeatCount = 0
                    while 1:
                        repeatCount += 1
                        try:
                            query = "INSERT INTO employee_schedule (ID, REFER_ID, BADGE_NO, EMPLOYEE_NO, SCHEDULE_DATE, SEQ_NO, SCHEDULE_TYPE, CREATED_BY, CREATED_DATE)" \
                                " VALUES('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % \
                                (str(currID), str(referID), emp, emp, currDay, '1', schedType, 'WFM_IFACE', now.strftime('%Y-%m-%d %H:%M:%S'))
                            cur.execute(query)
                            currID += 1
                            break;
                        except pyodbc.IntegrityError, e:
                           # Duplicate record error
                           # check the overwrite data checkbox
                           if self.chOverWrite.isChecked():
                               cur.execute("Delete from EMPLOYEE_SCHEDULE where BADGE_NO = '%s' and SCHEDULE_DATE = '%s' and SEQ_NO = %d" % (emp, currDay, 1))
                           else:
                               break

                           # check if repeated already
                           if repeatCount > 1:
                               break

                    self.labelStatus.setText('Saving schedule of employee : ' + str(emp) + ' to Orisoft database.')
                    self.repaint()


        # save new next_record ID of table employee_schedule
        if saveID <> currID:
            cur.execute("update ofcctrlid set ctrlctr = '%s' where ctrlcol = 'employee_schedule'" % (currID))

        connOriTMS.commit()

    def viewExceptionReport(self):
        cursor = connOriTMS.cursor()
        cursor.execute("Select * from USER_WFM_EXCEPTION order by EMPLOYEE_NAME, SCHEDULE_DATE")
        exceptionReport = []
        for rec in cursor:
            line = []
            line.append(rec[1])                 # employee no
            line.append(rec[2])                 # employee name
            line.append(rec[3])                 # schedule date
            line.append(rec[4])                 # schedule type
            line.append(rec[5])                 # workgroup
            line.append(rec[6])                 # remarks
            line.append(rec[7])                 # created by
            line.append(rec[8])                 # created date

            exceptionReport.append(line)

        cursor.close()
        if not exceptionReport:
            self.labelStatus.setText(self.labelStatus.text() + ' No exception report.')
            return

        rept = WfmReport(exceptionReport,['EMPLOYEE NO', 'EMPLOYEE NAME', 'SCHEDULE DATE', 'SCHEDULE_TYPE', 'WORKGROUP', 'REMARKS', 'CREATED BY', 'CREATED DATE'], self)
        rept.resize(800,600)
        rept.setWindowTitle("WFM Interface Exception Report")
        rept.exec_()

    def canceled(self):
        # show the confirmation message
        if self.cancelButton.text() == 'Cancel':
            flags = QMessageBox.StandardButton.Yes
            flags |= QMessageBox.StandardButton.No
            question = "Do you really want to cancel?"
            response = QMessageBox.question(self, "Confirm Cancel", question, flags, QMessageBox.No)
            if response == QMessageBox.No:
                return

        self.abort()

    def abort(self):
        connOriTMS.close()
        self.reject()
        app.exit(1)

    def saveIni(self):

        # name of configuration file
        iniFile = 'WFM_Interface.ini'
        config = ConfigParser.ConfigParser()
        config.read(iniFile)

        # save dateFrom/dateTo
        config.set('History','datefrom',self.dateEditFrom.date().toPython())
        config.set('History', 'dateto', self.dateEditTo.date().toPython())

        ini = open(iniFile, 'w')
        config.write(ini)
        ini.close()


def readIni():
    global orisoftDsn, orisoftUser, orisoftPwd
    global wfmDsn, wfmUser, wfmPwd
    global dateFromPrev, dateToPrev

    # name of configuration file
    iniFile = 'WFM_Interface.ini'
    # test if config file exists
    if not os.path.exists(iniFile):
        QMessageBox.critical(None,'Config File Missing', "The configuration file '%s' in '%s' not found!" % \
            (iniFile, os.getcwd()))
        app.exit(1)

    try:
        config = ConfigParser.ConfigParser()
        config.read(iniFile)

        # read Orisoft TMS settings
        orisoftDsn = config.get('OrisoftTMSDSN', 'dsn')
        orisoftUser = config.get('OrisoftTMSDSN', 'uid')
        orisoftPwd = config.get('OrisoftTMSDSN', 'pwd')

        wfmDsn = config.get('WFM_DSN', 'dsn')
        wfmUser = config.get('WFM_DSN', 'uid')
        wfmPwd = config.get('WFM_DSN', 'pwd')
        # read last date processed
        dateFromPrev = config.get('History', 'datefrom')
        dateToPrev = config.get('History', 'dateto')

    except ConfigParser.NoSectionError, e:
        QMessageBox.critical(None,'Config File Error', str(e))
        app.exit(1)

    except ConfigParser.NoOptionError, e:
        QMessageBox.critical(None, 'Config File Error', str(e))
        app.exit(1)


app = QApplication(sys.argv)
try:
    readIni()
    # connection for Orisoft TMS Database
    connOriTMS = pyodbc.connect('DSN=%s; UID=%s; PWD=%s' % (orisoftDsn, orisoftUser, orisoftPwd))
    connWFM = pyodbc.connect('DSN=%s; UID=%s; PWD=%s' %(wfmDsn, wfmUser, wfmPwd))

except pyodbc.Error, e:
    QMessageBox.critical(None,'WFM-Interface Connection Error', str(e))
    sys.exit(1)


employees = {}                          # employees daily attendance
daysRange = []                          # list of days from DateFrom to DateTo in the entry panel
groupSchedule = {}                      # workgroup schedule table
validSchedType = set()                  # a set of valid schedule types in Orisoft

orisoftDsn, orisoftUser, orisoftPwd
wfmDsn, wfmUser, wfmPwd
dateFromPrev, dateToPrev

form = WFMInterface()
form.show()
sys.exit(app.exec_())
