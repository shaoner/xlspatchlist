#! /usr/bin/env python

import sys

try:
    import xlwt, xlrd
except:
     print 'Required modules: python-xlrd and python-xlwt'
     sys.exit(1)

import argparse
import subprocess
import datetime
import re
import os.path

class Display:
    CYAN = '\033[95m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    ORANGE = '\033[93m'
    RED = '\033[91m'
    ENDC = '\033[0m'

    @staticmethod
    def header(message):
        line =  Display.CYAN + '-' * 80 + Display.ENDC
        print line
        print "* " + message
        print line

    @staticmethod
    def row(message):
        print Display.BLUE + '*' + Diplay.ENDC + ' ' + message

    @staticmethod
    def err(message):
        print '[{0}E{1}] {2}'.format(Display.RED, Display.ENDC, message);

    @staticmethod
    def ok(message):
        print '[{0}OK{1}] {2}'.format(Display.GREEN, Display.ENDC, message);

class Spreadsheet:
    """
    XLS Spreadsheet proxy
    """

    tableHeader = xlwt.easyxf("font: bold 1, height 250; pattern: pattern solid, fore-colour light_orange; border: left thick, bottom thick, right thick, top thick")
    tableBody = xlwt.easyxf("border: left thick, bottom thick, right thick, top thick")
    linkStyle = xlwt.easyxf("font: bold 1, colour blue; border: left thick, bottom thick, right thick, top thick")

    def __init__(self, fSpreadSheet, sSheetName):
        """ Constructor """
        self.fSpreadSheet = fSpreadSheet
        self.workbook = None
        self.worksheet = None
        self.currow = 0
        # If file already exists
        # data are appended
        if os.path.exists(fSpreadSheet):
            self.open(sSheetName)
        else:
            self.create(sSheetName)

    def open(self, name):
        """ Open an excel workbook with a worksheet <name> """
        book = xlrd.open_workbook(self.fSpreadSheet)
        dXlsData = { }
        for sheet in book.sheets():
            dXlsData[sheet.name] = [ ]
            for rownum in range(1, sheet.nrows):
                dXlsData[sheet.name].append(sheet.row_values(rownum))
        # New workbook
        self.workbook = xlwt.Workbook()
        # Copy data into it
        for sheetname, lData in dXlsData.iteritems():
            sheet = self.workbook.add_sheet(sheetname)
            # Save the current worksheet if it already exists
            if name == sheetname:
                self.worksheet = sheet
                self.currow = len(lData) + 1
            self.write_header(sheet)
            iRow = 1
            for lRow in lData:
                for k in range(0, 6):
                    sheet.write(iRow, k, lRow[k], Spreadsheet.tableBody)
                iRow += 1
        # Create a new sheet if not found
        if self.worksheet == None:
            self.new_sheet(name)

    def create(self, name):
        """ Create a new workbook with worksheet <name> """
        self.workbook = xlwt.Workbook()
        self.new_sheet(name)

    def new_sheet(self, name):
        """ Create a new excel worksheet """
        self.worksheet = self.workbook.add_sheet(name)
        self.write_header(self.worksheet)
        self.currow = 1

    def display(self, text, size, end = False, header = False):
        nLen = len(text)
        if nLen >= size:
            text = text[:(size - 3)] + '...'
        if header:
            stytext = '\033[94m' + text + '\033[0m'
        else:
            stytext = text;
        print '|' + stytext + ' ' * (size - len(text)),
        if end:
            print

    def write_header(self, sheet):
        """ Write a table header """
        if sheet == None:
            raise Exception("No worksheet")
        sheet.col(0).width = 42 * 256
        sheet.col(1).width = 25 * 256
        sheet.col(2).width = 15 * 256
        sheet.col(3).width = 15 * 256
        sheet.col(4).width = 8 * 256
        sheet.col(5).width = 50 * 256
        sheet.write(0, 0, 'SHA', Spreadsheet.tableHeader)
        sheet.write(0, 1, 'Author', Spreadsheet.tableHeader)
        sheet.write(0, 2, 'Reference', Spreadsheet.tableHeader)
        sheet.write(0, 3, 'Component', Spreadsheet.tableHeader)
        sheet.write(0, 4, 'Type', Spreadsheet.tableHeader)
        sheet.write(0, 5, 'Message', Spreadsheet.tableHeader)
        print '-' * 155
        self.display('SHA', 42, header = True)
        self.display('Author', 25, header = True)
        self.display('Reference', 15, header = True)
        self.display('Component', 15, header = True)
        self.display('Type', 8, header = True)
        self.display('Message', 50, end = True, header = True)
        print '-' * 155

    def write_commit(self, commit):
        """ Write a table row """
        if self.worksheet == None:
            raise Exception("No worksheet")
        row = self.currow
        #print('[ ' + commit.sha + ' - ' + commit.author + ' - ' + commit.ref + ' - ' + commit.component + ' - ' + commit.stype + ' ]')
        self.worksheet.write(row, 0, commit.sha, Spreadsheet.tableBody)
        self.worksheet.write(row, 1, commit.author, Spreadsheet.tableBody)
        if commit.url:
            link = xlwt.Formula('HYPERLINK("' + commit.url + '"; "' + commit.ref + '")')
            self.worksheet.write(row, 2, link, Spreadsheet.linkStyle)
        else:
            self.worksheet.write(row, 2, commit.ref, Spreadsheet.tableBody)
        self.worksheet.write(row, 3, commit.component, Spreadsheet.tableBody)
        self.worksheet.write(row, 4, commit.stype, Spreadsheet.tableBody)
        self.worksheet.write(row, 5, commit.message, Spreadsheet.tableBody)
        self.currow += 1
        self.display(commit.sha, 42)
        self.display(commit.author, 25)
        self.display(commit.ref, 15)
        self.display(commit.component, 15)
        self.display(commit.stype, 8)
        self.display(commit.message, 50, True)

    def save(self):
        """ Save the current workbook """
        self.workbook.save(self.fSpreadSheet)

class CommitParser:
    """
    Represents a commit line with its attributes
    """
    def __init__(self, sha, author, stype, ref, component, message):
        """ Constructor """
        self.sha = sha
        self.author = author
        self.stype = stype
        self.ref = ref
        self.component = re.sub('^ | $', '', component)
        message = message.replace('\n', ' ')
        self.message = re.sub('^ | $', '', message)
        try:
            bz = ref.split("BZ#")
            if len(bz) > 1:
                self.url = "https://bugzilla.stlinux.com/show_bug.cgi?id=" + bz[1]
            else:
                raise
        except:
            self.url = None

    def to_tuple(self):
        return (self.sha, self.author, self.stype, self.ref, self.component, self.message)

    @staticmethod
    def get_type_from_input():
        aStatus = [ 'FIX', 'FEATURE', 'UPDATE' ]
        j = 1
        for s in aStatus:
            print '[' + str(j) + '] ' + s
            j += 1
        n = 0
        lStatus = len(aStatus)
        while not (n >= 1 and n < lStatus):
            nin = raw_input('Choose: ')
            n = int(nin)
        return aStatus[n]

    @staticmethod
    def get_ref_from_input():
        inRef = raw_input('Bugzilla number (BZ#XXX) or reference? (type <enter> if not) ')
        if inRef == '':
            return 'NONE'
        return inRef

    @staticmethod
    def get_component_from_input():
        inComponent = raw_input('CustomXXXXXX or HPK or etc? (type <enter> if not) ')
        if inComponent == '':
            return 'GENERIC'
        return inComponent

    @staticmethod
    def find_attr(aSubject):
        """ Find attributes """
        sType = ''
        sRef = ''
        sComponent = ''
        sMessage = ''
        for tok in aSubject:
            if re.match('(UPDATE)|(FEATURE)|(FIX)', tok):
                sType = tok
            elif re.match('(BZ\#[0-9]+)|(BRISSEC\#[0-9]+)|(NONE)', tok):
                sRef = tok
            elif re.match('([A-Z0-9]{3,15})', tok):
                sComponent = tok
            else:
                sMessage += tok
        if sType == '':
            print 'Line: ' + str(aSubject)
            sType = CommitParser.get_type_from_input()
        if sRef == '':
            print 'Line: ' + str(aSubject)
            sRef = CommitParser.get_ref_from_input()
        if sComponent == '':
            print 'Line: ' + str(aSubject)
            sComponent = CommitParser.get_component_from_input()
        while sMessage == '':
            sMessage = raw_input('Message? ')
        return (sType, sRef, sComponent, sMessage)

    @staticmethod
    def get_attr(line):
        """ Get attributes by parsing a log line """
        # This line should look like
        # SHA|AUTHOR|TYPE-REF-COMPONENT: MESSAGE
        # TYPE = FIX|FEATURE|UPDATE
        # REF = BZ#XXX|RNDXXX|NONE
        # COMPONENT = CUSTOMYYYYYY|HPK|GENERIC
        aInfo = line.split('|')
        if len(aInfo) < 3:
            raise SyntaxError
        sSha = aInfo[0]
        sAuthor = aInfo[1]
        sCmMessage = aInfo[2]
        aSAndMessage = sCmMessage.split(':')
        sType = ''
        sRef = ''
        sComponent = ''
        sMessage = ''
        if len(aSAndMessage) < 2:
            # This should be a SyntaxError, but I'm not a dictator.
            return (sSha, sAuthor, 'FIX', 'NONE', 'GENERIC', sCmMessage)
        # aSubject should contain
        # [ TYPE, REF, COMPONENT, MESSAGE ]
        # if not, then it is a commit from division
        # or with a bad format
        aSubject = aSAndMessage[0].split('-')
        # <component>: <message>
        if len(aSubject) == 1 or len(aSubject) == 2:
            sMessage = aSAndMessage[1]
            try:
                sRef = re.search('(RnDHV[0-9]{8}?)', sMessage).group(1)
            except:
                sRef = 'NONE'
            return (sSha, sAuthor, 'FIX', sRef, aSAndMessage[0], sMessage)
        # <A>-<B>-<C>: <message>
        aSubject.append(aSAndMessage[1])
        if len(aSubject) != 4:
            # More that 4 elements so, need a user interaction
            return (sSha, sAuthor) + CommitParser.find_attr(aSubject)
        sType = aSubject[0]
        sRef = aSubject[1]
        sComponent = aSubject[2]
        sMessage = aSubject[3]
        return (sSha, sAuthor, sType, sRef, sComponent, sMessage)

def main():
    datePrefix = datetime.datetime.now().strftime("%Y%m%d")
    parser = argparse.ArgumentParser()
    parser.add_argument("commit", type = str,
                        help = "Commit name (release_X..release_Y, <since>..<until>, HEAD)")
    parser.add_argument("sheet", type = str, default = 'SDK', choices = [ 'SDK', 'KERNEL', 'MULTICOM'],
                        help = "The worksheet name")
    parser.add_argument("-o", "--output", type = str, nargs = '?', default = 'PATCH_LIST_' + datePrefix + '.xls',
                        help = "XLS file to output")
    parser.add_argument("-q", "--quiet", action="count",
                        help="Quiet mode on stdout")
    args = parser.parse_args()
    # Run git log
    proc = subprocess.Popen([ 'git','log', '--pretty=format:%H|%an|%s', '--no-merges', '--reverse', args.commit ],
                            stdout = subprocess.PIPE)
    # Get output into an array of each git log line
    aOutput = proc.stdout.readlines()
    Display.header("Generating worksheet " + args.sheet + " in excel file " + args.output)
    # Create a new Spreadsheet
    ss = Spreadsheet(args.output, args.sheet)
    # Write the commit list
    for line in aOutput:
        #    print line
        line = line.decode('utf-8', errors = 'ignore')
        try:
            sSha, sAuthor, sType, sRef, sComponent, sMessage = CommitParser.get_attr(line)
        except SyntaxError:
            Display.err('Syntax error for ' + line)
            sSha, sAuthor, sType, sRef, sComponent, sMessage = ('NONE', 'NONE', 'FIX', 'NONE', 'GENERIC', line)
        except:
            Display.err('Unexpected error at line ' + line)
            raise
            return 1
        else:
            c = CommitParser(sSha, sAuthor, sType, sRef, sComponent, sMessage)
            ss.write_commit(c)
    print
    Display.ok('Patch list created!')
    # Save the patch list
    ss.save()
    return 0

if __name__ == "__main__":
    iRet = main()
    sys.exit(iRet)
