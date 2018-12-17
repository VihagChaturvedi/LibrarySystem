import random
import xlwt
#from xlwt import Workbook
import xlrd
#bugs:the program asks for key when restarted
#solutions:make it a function and activate seperately
a = True
student_name = []
class_section = []
admin_no = []
codes = []
book_borrowed = []
book_copy_amount = []
class library():
    def borrow_a_book(self):
        try:
            print("You are now borrowing a book")
            print("The following books are available:\n", booklist, "\nCopies:", booklist_copies)
            book_borrower = input("Which book would you like to borrow?").upper()
            book_copy_amount_prompt = int(input("How many Copies would you like to borrow?"))
            tempval = booklist.index(book_borrower, 0, len(booklist))
            currval = booklist_copies.__getitem__(tempval)
            currval = currval - book_copy_amount_prompt
            booklist_copies[tempval] = currval
            print("Please provide the following information to continue:")
            fullname = input("Full Name:")
            # input begins here
            classandsection = input("Class and Section'in roman numerical':")
            # class_section.append(classandsection)
            admissionno_optional = int(input("Admission No.:"))
            verifier = input("Please insist the librarian to review the details:")
            if verifier == verifier_code:
                print("Your book has been borrowed successfully!")
                code_generator = random.randint(00000, 99999)
                code_issuer = print(
                    "Thank You for your time! Please keep this code safely as it would be critical while returning your book:",
                    code_generator)
                codes.append(code_generator)
                book_borrowed.append(book_borrower)
                student_name.append(fullname)
                book_copy_amount.append(book_copy_amount_prompt)
                admin_no.append(admissionno_optional)
            else:
                print("Invalid verification procedure encountered.Please restart your request")
                pass
        except ValueError:
            print("Resetting Program due to invalid input")
    def return_a_book(self):
        try:
            print("You are now returning a book")
            returner = int(input("Welcome! Please provide your Verification code:"))
            if returner in codes:
                code_book_relator =codes.index(returner, 0, len(codes))
                bookret = book_borrowed.__getitem__(code_book_relator)
                copy_check = int(input("How many copies had you borrowed?"))
                booklist_copies.append(copy_check)
                copy_book_relator = booklist_copies.index(copy_check, 0, len(booklist_copies))
                print("Your book",code_book_relator,"has been returned successfully")
                copyret =book_borrowed.index(int(copy_book_relator))
                booklist_copies.append(copyret)
                codes.remove(code_book_relator)
            else:
                print("Incorrect Code given, please re-enter your credentials")
        except ValueError:
            print("Resetting program due to invalid input")
    def administartion_access(self):
        try:
            security = input("HighSecurity Password:")
            if security == access_key:
                splitter = input("Would you like to open the Database?").lower()
                if splitter == 'yes':
                    print("Access Granted")
                else:
                    print("No Problem!")
            else:
                print("Resetting program.")
        except ValueError:
            print("Resetting Program due to invalid input")
    def exit_sequence(self):
        try:
            exiter = input("Please enter the code to exit the program")
            if exiter == exit_code:
                print('Thank you for using LibraryManagement System by Vihag!')
                caller = library
                filename = "library_data.xls"
                try:
                    with open(filename) as file1:
                        wb = xlwt.Workbook()
                        # add_sheet is used to create sheet.
                        sheet1 = wb.add_sheet('Sheet_1')
                        sheet1.write(1, 0, student_name)
                        sheet1.write(1, 1, class_section)
                        sheet1.write(1, 2, admin_no)
                        wb.save(filename)
                except IOError:
                    wb = xlwt.Workbook()
                    # add_sheet is used to create sheet.
                    sheet1 = wb.add_sheet('Sheet_1')
                    sheet1.write(1, 0, student_name)
                    sheet1.write(1, 1, class_section)
                    sheet1.write(1, 2, admin_no)
                    wb.save(filename)
                quit()
        except ValueError:
            print("Exit sequence unsuccessful")
        except IOError:
            print("Exit sequence unsuccessful")
    def report_student(self):
        try:
            print("You are now reporting a student.")
            student_name = input("Name:")
            student_class = input("Class and section")
            student_admin_no = int(input("Admission no.:"))
            student_reason = input("Reason:")
            print("Student reported.")
            wb=xlwt.Workbook("library_report_a_student.xlsx")
            sheet2=wb.add_sheet('Sheet_1')
            sheet2.write(1,0,student_name)
            sheet2.write(1,2,student_admin_no)
            sheet2.write(1,3,student_class)
            sheet2.write(1,4,student_reason)
        except IOError:
            print("Complain registration unsuccessful")
        except ValueError:
            print("Invalid Input provided")
    def report_damage(self):
        try:
            print("You are now reporting damage of school's property in the library.")
            item_report = input("What is the damage?").lower()
            print("Please fill in the following detail to continue:")
            student_name = input("Name:")
            student_class = input("Class and section")
            student_admin_no = int(input("Admission no.:"))
            if item_report == 'stool':
                print("Your fine will be a total of Rs.200")
            elif item_report == 'table':
                print("Your fine will be a total of Rs.1000")
            elif item_report == 'glass':
                print("Your fine will be a total of Rs.1500")
            else:
                print("Damage to be sorted with the incharge")
            wb=xlwt.Workbook("library_damage.xlsx")
            sheet3=wb.add_sheet("Sheet_3")
            sheet3.write(1,0,student_name)
            sheet3.write(1,2,student_class)
            sheet3.write(1,3,student_admin_no)
            sheet3.write(1,4,item_report)
        except IOError:
            print("Complain Registration unsuccessful")
        except ValueError:
            print("Complain Registration unsuccessful due to improper input")
    def append_from_file(self,student_naam):
        if student_naam!="":
            student_name.copy()
        else:
            try:
                with open("library_data.xls") as file:
                    wb = xlrd.open_workbook("library_data.xls")
                    # add_sheet is used to create sheet.
                    sheet = wb.sheet_by_index(0)
                    print(sheet.cell_value(0, 0), "value in cell")

                    # For row 0 and column 0
                    if sheet.cell_value(0, 0)==" ":
                        student_naam =list(sheet.cell_value(0, 0))

                    if sheet.cell_value(0, 0)==" ":
                        class_section =[]
                    else:
                        class_section = sheet.cell_value(1, 0)
                    if sheet.cell_value(0, 0)==" ":
                        admin_no =[]
                    else:
                        admin_no = list(sheet.cell_value(2, 0))
            except IOError:
                pass
    def sequence_key(self):
        access_key = 'MeridianMadhapur567'
        sequence_key_looper=True
        while sequence_key_looper==True:
            high_authority_key='HigherAuthorityOfMeridian'
            tries=0
            access_key_checker = input("Please provided the access key in order to begin the program:")
            try:
                if access_key == access_key_checker:
                    sequence_key_looper=False
                else:
                    print("Invalid Key")
                    tries=tries+1
                    if tries==3:
                        high_authority=input("You have given three consecutive incorrect keys. Please enter HighAuthority passcode:")
                        if high_authority!=high_authority_key:
                            quit()
                        else:
                            pass
            except ValueError:
                print("Please restart your request as invalid input procedure ws encountered.")
while a == True:
    access_key='MeridianMadhapur567'
    print('Hello and Welcome to LibraryAssistant!')
    booklist=["CHEMISTRY CLASS 8","PHYSICS CLASS 11","BIOLOGY CLASS 11","CHEMISTRY CLASS 3","CHEMISTRY CLASS 4" ]
    booklist_copies=[5,5,5,5,5]
    verifier_code='1357924680'
    exit_code='borroweraccessed'
    copy_1=book_copy_amount
    launcher = library()
    bridge = True
    while bridge == True:
        launcher.sequence_key()
        try:
                print("The syntaxes are as follows:")
                print("bb= borrow a book\n", "rb= return a book\n", "rs= report a student\n", "rd= report damage\n",
                      "es= exit")
                guide = input("What would you like to do in the library using the given syntaxes?").lower()
                if guide == 'bb':
                    launcher.borrow_a_book()
                elif guide == 'rb':
                    launcher.return_a_book()
                elif guide == 'administration_access':
                    launcher.administartion_access()
                elif guide == 'rs':
                    launcher.report_student()
                elif guide == 'rd':
                    launcher.report_damage()
                elif guide == 'es':
                    launcher.exit_sequence()
                else:
                    bridge = False
                    print("Your input doesn't match the syntaxes provided.")
        except ValueError:
            print("Your input doesn't match the syntaxes provided.")