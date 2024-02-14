import random
import smtplib
from csv import writer

# In order to run this program, please save the 'team_roster' Excel file as a csv in the same folder.
# Overwrite the last file to update the information that will be read in this program.

# This python file is dependent on the names entered and will assign tasks without repeating until everyone has contributed.
# Names in the excel will receive an email delegating their task and due date for 5S activity.
# The tasks assigned are base on Boeing 5S Lean Assessment goals.

####################################################################################################

# Please enter your email credentials below:
sender = 'xxxxxxxxxx@email.com'
password = 'xxxxxxxxxxxxxx'

####################################################################################################

team_list = [] # listed as the first column from the csv file
email_list = [] # listed as the second column from the csv file
task_list = [] # listed as the third column from the csv file
date_list = [] # listed as the forth column from the csv file

print("\nReading through EO North Employee Roster...")

with open('team_roster.csv', "r") as csv_file:
    for line in csv_file:
        team_list.append(line.split(",")[0])
        email_list.append(line.split(",")[1])
        task_list.append(line.split(",")[2])
        date_list.append(line.split(",")[3])

print("Randomly assigning tasks and due dates to team...\n")

random.shuffle(task_list) # randomly assigns tasks while keeping names and emails in order
random.shuffle(date_list) # randomly assigns due dates while keeping names and emails in order

count = 0

open("results.txt", "w").close() #clears the current results file

for address in email_list:
    with open('results.txt', 'a') as text_file:
        prompt = team_list[count] + " has been assigned to the task " + task_list[count] + " with a completion date of " + date_list[count]
        text_file.write(prompt)
        print(prompt)

    s = smtplib.SMTP('smtp-mail.outlook.com', 587) # creates SMTP session based on Outlook server location and port
    s.starttls() # start TLS for security (Transport Layer Security) encrypts all the SMTP commands
    s.login(sender, password) # Authentication account credentials provided at the top of this Python file
    message = "Hello, " + team_list[count] + ".\n\n You have been assigned to the EO North 5S task of " + task_list[count] + ".\n Please complete this task by " + date_list[count] + ".\n For any questions please reach out to jason.johnson-yurchak@boeing.com.\n\n Thanks!" # message to be sent
    s.sendmail(sender, address, message) # sending the mail
    s.quit() # terminating the session

    count += 1