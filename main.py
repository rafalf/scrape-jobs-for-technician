import invoice_xlsx_handler
import jobs_sheet_xlsx_handler

job1 = ["03-I-16", "mrichardson", "16552315", "1/20", "Camira Place, Newnham", "", "YES", "", "", "", "", ""]
job2 = ["04-I-16", "mrichardson", "16544987", "17", "Walkers Avenue, Newnham", "1", "End User - Underground",
        "Battery Backup", "", "YES", "", "sure", "we can"]

invoice_xlsx_handler.create_file_and_write_to([job1, job2])
# invoice_xlsx_handler.add_to_file([job1, job2])
# print(invoice_xlsx_handler.read_from_file())


job1 = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17"]

jobs_sheet_xlsx_handler.create_file_and_write_to([job1])
# jobs_sheet_xlsx_handler.add_to_file([job1])
