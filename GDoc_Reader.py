import re
import csv
import os
import datetime
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import client_list

class ReadSpreadsheet:

    def __init__(self, chosenClient, ourClients):
        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        self.KEY_FILE_LOCATION = 'client_secrets.json'
        self.chosenClient = chosenClient
        self.clients = ourClients
        self.sheets = []
        self.data = []
        self.cleanData = []

    def initialize(self):
        # Initializes an API V4 service object.
        #
        # Returns:
        #   An authorized API V4 service object.

        credentials = ServiceAccountCredentials.from_json_keyfile_name(
            self.KEY_FILE_LOCATION, self.SCOPES)

        # Build the service object.
        self.serviceObject = build('sheets', 'v4', credentials=credentials)

    def getSheets(self):
        spreadsheetidname = self.clients[self.chosenClient]

        result = self.serviceObject.spreadsheets().get(spreadsheetId=spreadsheetidname).execute()

        mysheets = result.get('sheets')

        for each in mysheets:
            self.sheets.append(each.get('properties').get('title'))

    def getData(self):
        spreadsheetidname = self.clients[self.chosenClient]
        majordimensionname = 'COLUMNS'
        rangenames = []
        for sheet in self.sheets:
            rangename = '%s!A:ZZ' % sheet
            rangenames.append(rangename)

        for ranges in rangenames:
            result = self.serviceObject.spreadsheets().values().get(
                spreadsheetId=spreadsheetidname, range=ranges, majorDimension=majordimensionname).execute()

            self.data.append(result.get('values', []))

    def emailCleaner(self):
        expression = re.compile(ur'[\w\.-]+@[\w\.-]+\.[a-zA-Z]+', re.UNICODE)
        for l1 in self.data:
            for l2 in l1:
                for l3 in l2:
                    match = re.search(expression, l3)
                    if match:
                        email = match.group(0)
                        if not any(email in s for s in self.cleanData):
                            self.cleanData.append(email)

    def csvWriter(self):
        todaysDate = datetime.date.today().strftime('%y%m%d')

        fileName = "Mailing_List_%s_%s.csv" % (self.chosenClient.replace(" ", "-"), todaysDate)

        with open(os.path.expanduser(os.path.join("~/Desktop", fileName)), 'wb') as csvfile:
            mailing_writer = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
            for i in self.cleanData:
                mailing_writer.writerow([i])

    def runAll(self):
        self.initialize()
        self.getSheets()
        self.getData()
        self.emailCleaner()
        self.csvWriter()


if __name__ == '__main__':
    for each in client_list.ourClients:
        a = ReadSpreadsheet(each, client_list.ourClients)
        a.runAll()
