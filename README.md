# Automating Outlook using Python win32com
## This project contains a script which can be run to carry out the below:
* Save all attachments on emails in a folder
* Save all email bodies in a folder as csv
* Move emails from folder to subfolder

### The below script can be called from the command line or can be embedded in another program so it is easy to schedule using a task scheduler (or crontab) as well as triggering manual runs as an embedded functionality in another program. All is explained below.

"Can you automate the saving down of attachments from a daily file I receive from x?"
This is often the first question I'm asked when I tell people I can help them automate their most menial tasks. And the short answer is: **Yes**.

There are two approaches. The VBA method which will require you to write this code in the VBE within Outlook. This is impossible to trigger via an e-mail event as Microsoft removed the ability to trigger a script on the back of one mid 2018. It is possible to use the same Outlook objects in python by way of a very poorly documented library called **[win32com](https://docs.microsoft.com/en-us/office/vba/api/overview/outlook/object-model)**. It was tough to get it to work, but once you get it to work it works without any of the funky hit-and-miss randomness anyone who's ever used VBA will have experienced. It was like a dream and here's how to do it.

## Prerequisites
Python 3 and **pip-install pywin32**

This will allow you to do
``` python
import win32com.client
```

## Main Script
There is one main function that has been written such that you can call it from the command line or a GUI. The only difference between them is how an error message is displayed. If It is from the command line success and errors are printed to the console whereas if called from a GUI the function will return the error message (presumably to be displayed somewhere on screen).

``` python
    def runOlScript(outdest, filefmt, olreadfolder, olprocessedfolder, guiEntry, proc):
    if _right(outdest,1) != '/':
        outdest = outdest + '/'

    # To Do: Deal with anything starting with re: or fw:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    try:
        inbox = outlook.Folders.Item(1).Folders['Inbox'].Folders[olreadfolder]
    except Exception as e:
        _scriptOutput(str(e) + " | Could not find Outlook Folder [{}]".format(olreadfolder), guiEntry)
    try:
        outlook.Folders.Item(1).Folders['Inbox'].Folders[olreadfolder].Folders[olprocessedfolder]
    except Exception as e:
        _scriptOutput(str(e) + " | Could not find Outlook Folder [{}]".format(olprocessedfolder), guiEntry)

    messages = inbox.Items
    if len(messages) == 0:
        _scriptOutput( 'No emails found in folder [{}]'.format(olreadfolder), guiEntry)
    
    mailCounter = 0
    for msg in list(messages):
        bProcessed = False
        if proc == 'olatt':
            for atmt in msg.Attachments:
                if filefmt == 'blank' or str.lower(_right(atmt.FileName, len(filefmt))) == str.lower(filefmt):
                    tmpFileName = outdest + msg.Subject + '_' + atmt.FileName
                    try:
                        atmt.SaveAsFile(tmpFileName)
                        print('File Successfully Saved [{}]'.format(tmpFileName))
                        bProcessed = True
                    except Exception as e:
                        _scriptOutput(str(e) + ' | File NOT saved [{}]'.format(tmpFileName), guiEntry)
        if proc == 'olbody':
            listbody = msg.Body.split("\r\n")
            tmpFileName = outdest + msg.Subject + '_' + msg.CreationTime.strftime("%Y%m%d") + '.csv'
            bProcessed = True
            with open(tmpFileName, 'w', newline='') as file:
                writer = csv.writer(file)
                for row in listbody:
                    writer.writerow([row])
        if bProcessed:
            mailCounter += 1
            msg.Move(outlook.Folders.Item(1).Folders['Inbox'].Folders[olreadfolder].Folders[olprocessedfolder])
        
    return 'Succesfully processed {} emails!'.format(mailCounter) if mailCounter > 0 else 'No emails processed'
```

runOlScript has 6 function arguments:
1. outdest - the path to the folder you wish to save the attachments or email bodies
2. filefmt - the file format you would **only** like to save. Helpful if you only want to save the .csv files across many e-mails
3. olreadfolder - the name of the outlook folder which has the e-mails you wish to process
4. olprocessedfolder - the name of the subfolder (which must be nested under the olreadfolder) you wish you move these sucessfully processed e-mails to
5. guiEntry - True if called from a GUI. False if called from a cmd prompt
6. proc - Name of the process being run "olatt" is save attachments only, "olbody" will save the text found in the body of the emails as a .csv each with a datestamp

## Calling This From a Command Line
You will need to pass the 5 function by way of cmd line argument and parse them programatically to read them into the main funtion call, is illustrated below.

``` python
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-ola","--olatt",default=False)
    parser.add_argument("-olb","--olbody",default=False)
    parser.add_argument("-out","--outdest",default="")
    parser.add_argument("-olf1","--olfolder",default="")
    parser.add_argument("-olf2","--olprocfolder",default="")
    parser.add_argument("-typ","--olfiletype",default="blank")
    args = parser.parse_args()
    
    if (not args.olatt and not args.olbody):
        sys.exit('No process choice made, choose between ol attachments saver (--olatt) and ol mail body saver (--olbody)!')
    if args.outdest == '':
        sys.exit('No out destination defined using --outdest defined!')
    if args.olfolder == '':
        sys.exit('No outlook folder to search mails for defined using --olfolder!')
    if args.olprocfolder == '':
        sys.exit('No outlook folder to move processed mails to defined using --olprocfolder!')
    proc = 'olatt' if args.olatt else 'olbody' if args.olbody else ''

    runOlScript(args.outdest, args.olfiletype, args.olfolder, args.olprocfolder, False, proc )
```

## Automating The Run
Setting up the command line line you require down as a bat file and attaching this to a task in the Windows Task Scheduler is the easiest way to do this. This is what my .bat file looks like which is set to run this every day at 5pm:

The first path is where your Python interpreter lives, the second path is where your script is saved and the rest of them the 5 arguments required to get this script to run.

``` bash
C:\GitProjects\dir1\dir2\Scripts\python.exe "C:/GitProjects/dir1/ol_script.py" --olbody True --olfolder CSVTester --olprocfolder CSVTesterProcessed --olfiletype csv --outdest "C:/Users/Username/Documents/CSV OutDir/"
```

## Support
For any questions, do not hesitate to e-mail me at Hridai@ThatAutomation.co

Windows Documentation: https://docs.microsoft.com/en-us/office/vba/api/overview/outlook/object-model

Full script available for download here at the top of this page!
