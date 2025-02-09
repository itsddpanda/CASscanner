# CASscanner
To scan detailed CAS shared via Karvy (https://mfs.kfintech.com/investor/General/ConsolidatedAccountStatement) 
Currently it only scan CAS statement (link to generate above) and creates a workbook, storing fund details in seperate excel sheets.

## A Scanner Again? Why??
It is a make shift attempt to get away from all the insane apps, away from the need to share your personal identifiable information with them. No doubt there are many portfolio scanner out there and **all are _better_ than this**.
But I just couldn't stand giving up my data when we can do it on our own. 
Financial Literacy is important and it sucks to see all are using this as a platform to make money while for everything else we got free apps.
Thus after sabbatical of 12 years I picked up coding again, with what i can remember, since i know a bit about excel thus started it in VBA.

Most of you reading this **are smarter** than me and I know first hand there is **no limit of talent** in India.
So use this, in whatever way you find fit. 
      Want to make a better one? Fork this **and just do it**
      Insipired to create your own, **go for it**
      This is lame attempt and you can do it better, I bet you can and you need to show it to the world

And at last all you you can do is to just improve the code here please do, its just me in my free time, trying to re-learn and code (thank god for chatgpt).
And if that is not even possible, share it with anyone who can use excel and want get away from these apps.

A quick user guide is on its way! Link will be updated soon.

### How to use?
Before you start, 
If you are importing-
1. Create a new excel file, save as macro enabled excel. Enable Developer tools.
2. Import the form and modules (There is repetetive code or extra code that is not used, feel free to delete with caution)
3. Enable the referencces, as per below screenshot. Follow Run steps from below
![](https://github.com/itsddpanda/CASscanner/blob/main/Project%20Refs.png)

### Run it aka use the excel sheet
Either you imported the modules (and form) or using the excel here in follow the steps and you should be golden

### Pre-Requisites
You would need CAS, - Get CAS from Karvy(https://mfs.kfintech.com/investor/General/ConsolidatedAccountStatement).
Since this file is password protected, you would need to unlock it and make it password free i.e. opens without password. Google "pdf password removal". 
You will get many online tools for it.

Secondly, goto (https://www.amfiindia.com/nav-history-download) and download Click on "Download Complete NAV Report in Text Format", save it and give it a name "NAVAll.txt".

Please do not miss these two.

#### 0.  Starting Fresh or Restart
Ensure you are starting fresh or if something went wrong in between and you are restarting. 

Simply follow next line
Goto developer tools or press ALT + F8. 
Select Macro Named **"Restart"** and hit RUN to ensure you are starting fresh.

#### 1.  Step 1
Now you are all set to start. Follow above step and Select Macro Named **"Step1_SelectPDFFile"**
- Select the unlocked PDF File as explained in Pre-Requisites
- Provide Name of the excel sheet where data will be stored (tested with .xlsx)
- Select NAVAll.txt (if you get an error saying file not found, place the file in the same folder as your macro excel sheet and restart)
If all 3 are provided now, you can Launch the app or

### 2 to 4 Remaining steps
Rest of the steps - keep on selecting the next step (2,3 and 4) in order and you should be ok.
