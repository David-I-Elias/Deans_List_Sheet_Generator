import pandas as pd
import os
import glob


class csvParser:
    firstName = []
    midName = []
    lastName = []
    email = []
    semester = []
    phone = []
    concatdf = {}

    firstNameKey = "sendername"
    midNameKey = "middle-name-fill-in-only-if-you-want-your-middle-name-on-your-certificate"
    lastNameKey = "last-name"
    emailKey = "senderemail"
    semesterKey = "deans-list-semester-requested"
    phoneKey = "phone"

    def run(self, path):
        all_files = glob.glob(os.path.join(path, "*.csv"))

        df_from_each_file = (pd.read_csv(f) for f in all_files)
        self.concatdf = pd.concat(df_from_each_file, ignore_index=True)

        self.firstName = self.concatdf[self.firstNameKey]
        self.midName = self.concatdf[self.midNameKey]
        self.lastName = self.concatdf[self.lastNameKey]
        self.email = self.concatdf[self.emailKey]
        self.semester = self.concatdf[self.semesterKey]
        self.phone = self.concatdf[self.phoneKey]

