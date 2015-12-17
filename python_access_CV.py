#John W Python Code
import csv, pyodbc

# MDB is the path to your database, change it based on where the file is located
MDB = ''; DRV = '{Microsoft Access Driver (*.mdb)}'; PWD = 'pw'
con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV,MDB,PWD))
cur = con.cursor()

print("This program is used to score graduate student CV data.")
firstname=str(input("Enter your first name."))
lastname=str(input("Enter your last name."))
middleini=str(input("Enter your middle initial."))
gpa=float(input("Enter your un-weighted GPA. (ex. 3.99)"))
gre=int(input("Enter your GRE score out of 340."))
pubnum=int(input("Enter the number of publications you have."))
classteach=int(input("Enter the number of classes you have conducted as an instructor."))
classta=int(input("Enter the number of classes you have acted as TA."))
score=float((((gpa*.4)+(gre*.2)+(pubnum*.2)+(classteach*.1)+(classta*.1))/73.6)*100)

print(firstname,middleini,lastname)
print("GPA: ",gpa,"GRE score: ",gre)
print("Number of publications: ",pubnum)
print("Classes taught: ",classteach,"Classes as TA: ",classta)

gpa=str(gpa)
gre=str(gre)
pubnum=str(pubnum)
classteach=str(classteach)
classta=str(classta)

cur.execute('''INSERT INTO gradstudents(First_Name,Last_Name,Middle_Initial,GPA,GRE_Score,Number_of_Publications,Classes_was_Instructor,Classes_was_TA)
VALUES(?,?,?,?,?,?,?,?)''',(firstname,lastname,middleini,gpa,gre,pubnum,classteach,classta))
cur.commit()

print(firstname,middleini,lastname," Score: ",score)
print("-------------------------------------------------------------------------"
print("Current Ranking: Highest to Lowest")

counter=1
cur.execute("SELECT First_Name, Last_Name, Score FROM gradstudents_Query");
while counter==1:
    row = cur.fetchone()
    if row is None:
        break
    print(row)

cur.close()
con.close()


