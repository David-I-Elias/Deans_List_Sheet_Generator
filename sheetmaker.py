import math
import xlwt

class sheetMaker:
    # workbook = xlsxwriter.Workbook("DeansList.xlsx", {'nan_inf_to_errors': True})
    # campusSheet = workbook.add_worksheet("On Campus List")
    # onlineSheet = workbook.add_worksheet("Online List")

    workbook = None
    masterSheet = None
    onlineSheet = None

    def __init__(self, firstName, midName, lastName, email, semester, phone):
        self.firstName = firstName
        self.midName = midName
        self.lastName = lastName
        self.email = email
        self.semester = semester
        self.phone = phone



    def run(self):
        self.workbook = xlwt.Workbook({'nan_inf_to_errors': True})
        self.masterSheet = self.workbook.add_sheet("Master List")

        self.masterSheet.write(0, 0, 'FIRST NAME')
        self.masterSheet.write(0, 1, 'MIDDLE NAME')
        self.masterSheet.write(0, 2, 'LAST NAME')
        self.masterSheet.write(0, 3, 'EMAIL')
        self.masterSheet.write(0, 4, 'SEMESTER REQUESTED')
        self.masterSheet.write(0, 5, 'PHONE')

        for i in range(len(self.firstName)):
            self.masterSheet.write(i + 1, 0, self.firstName[i])

            # fix the #NAN writes
            try:
                if math.isnan(float(self.midName[i])):
                    self.masterSheet.write(i + 1, 1, ' ')
            except ValueError:
                self.masterSheet.write(i + 1, 1, self.midName[i])

            self.masterSheet.write(i + 1, 2, self.lastName[i])
            self.masterSheet.write(i + 1, 3, self.email[i])
            self.masterSheet.write(i + 1, 4, self.semester[i])
            self.masterSheet.write(i + 1, 5, self.phone[i])

        return self.workbook

