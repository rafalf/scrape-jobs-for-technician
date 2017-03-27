## Files:
__Must exist__: (Don't delete)
* template.xlsx - the template for the invoice.xlsx
* scraped.csv - scraped data is saved into this file along with the scraped time

## Scripts arguments:
* -v : verbose mode 
* -l : log output in the log file
* --run-mode : email or print. print - runs once a day and prints out jobs.xlsx sheet, email - runs every x minute and send an email
if new assignments found or any assignment has been changed.
* --test : to run in the test mode; in this mode when run together with __--run-mode email__, an email is sent out to: __test_email:__ from
label.conf

## Run script:
* configure a daily schedule task with the following arguments: ```python run.py --run-mode print -l```
* configure a schedule task that runs every x minute with the following arguments: ```python run.py --run-mode email -l```

## To test:
the best way to test it is to:
* delete all entries in scraped.csv file except for the headings
* run ```run.py --run-mode print``` - it should create a new job file jobs.xlsx and print it out, jobs should be saved into the invoice sheet
* remove a couple of jobs from scraped.csv and run ```run.py --run-mode print --test```. the script should collect all jobs once
again and since we deleted some from the scraped.csv, an email should be sent out the __test_email:__ with the new assignments.

## Technicians:
In label.conf file, fill data each technician:</br>
name_1::Si*** Ti****</br>
username_1::s*******</br>
pass_1::*******</br>
email_1::******</br>

## Install:
pip install emails</br>
pip install -U selenium</br>
pip install openpyxl</br>