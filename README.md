# Prize-Bond-Automation-Script-Selenium
### Step-1:
1. Create an xlsx sheet which contain 2 columns, 1st row 1st column contain heading **Bond Numbers**,1st row 2nd column contain **Bond Status**,
2nd (row till end) 1st column contain **123456 (6-digit bond numbers)**,2nd (row till end) 2nd column is empty , it will **automatically filled 
after the code execution**,(check xlsx sheet for guidance in Sheet Folder).

2. Enter All your Bond Numbers once in the xlsx sheet you created, save it.

### Step-2:
  Clone the Code <br/>git clone https://github.com/SaadAhmed00/Prize-Bond-Automation-Script-Selenium.git

### Step-3:
  1. Put your saved xlsx file in Sheet folder.
  2. Change path with respect to your xlsx sheet path in line 15 and line 52<br />
     path = "Your Path" <br />
     my_wb.save("Your Path")

### Step-4:
  1. Write your bond amount (e.g 200,750,1500,7500 etc) on line number 33. <br />
  select.select_by_visible_text('**Your Bond Amount in signle inverted comma**') <br/>
  Example: select.select_by_visible_text('**750**')
  
### Step-5:
  Run the Code and wait for the magic.

### Step-6:
  1. After the execution completed open your xlsx file , you will see the results in front of every bond numbers.
  2. Select those numbers whose status are **Congratulations! You have won 1 prize(s).**. These are your bond numbers who won the prize.

### Step-7:
  Before the next execution of code clean your xlsx file **entire rows below the Bond Status Column** and save it and then run the code for next bond announcement.
