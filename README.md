# Automating Outlook using Python win32com
Saving attachments from a daily e-mail or moving an e-mail from one folder to another is super easy using Python!
You can carry out literally any repetitive task using the Win32com library and the below will demonstrate a basic example which carries out the below:
* Goes through emails in a folder one by one and saves the attachment to a windows directory of your choice
* Moves the processed e-mail to another Outlook folder

## This project contains a script which can be run to carry out the below:
* Save all attachments on emails in a folder
* Save all email bodies in a folder as csv
* Move emails from folder to subfolder

### The below script can be called from the command line or can be embedded in another program so it is easy to schedule using a task scheduler (or crontab) as well as triggering manual runs as an embedded functionality in another program. All is explained below.

"Can you automate the saving down of attachments from a daily file I receive from x?"
This is often the first question I'm asked when I tell people I can help them automate their most menial tasks. And the short answer is: **Yes!**.

There are two approaches. The VBA method which will require you to write this code in the VBE within Outlook. This is impossible to trigger via an e-mail event as Microsoft removed the ability to trigger a script on the back of one mid 2018. Task Scheduler is my method of choice to automate this functionality. It is possible to use the same Outlook objects in python by way of a very poorly documented library called **[win32com](https://docs.microsoft.com/en-us/office/vba/api/overview/outlook/object-model)**. It was tough to get it to work, but once you figure it out, it works without any of the funky hit-and-miss randomness anyone who's ever used VBA will have experienced. **It works like a dream and here's how to get started**.

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
    outdest = os.path.normpath(outdest)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = _find_subfolder(outlook.Folders, olreadfolder)
    if inbox is None:
        sys.exit(f'No Folder {olreadfolder} found!!! Exiting.')
    procbox = _find_subfolder(outlook.Folders, olprocessedfolder)
    if procbox is None:
        sys.exit(f'Folder {olprocessedfolder} not found!!! Exiting.')

    messages = inbox.Items
    if len(messages) == 0:
        _scriptOutput( 'No emails found in folder [{}]'.format(olreadfolder), gui_entry)
    mail_counter = 0
    for msg in list(messages):
        b_processed = False
        if proc == 'olatt':
            for atmt in msg.Attachments:
                if filefmt == 'blank' or str.lower(_right(atmt.FileName, len(filefmt))) == str.lower(filefmt):
                    temp_filename = os.path.normpath(os.path.join(outdest, f'{msg.Subject} {atmt.FileName}'))
                    try:
                        atmt.SaveAsFile(temp_filename)
                        print('File Successfully Saved [{}]'.format(temp_filename))
                        b_processed = True
                    except Exception as e:
                        _scriptOutput(str(e) + ' | File NOT saved [{}]'.format(temp_filename), gui_entry)
        if proc == 'olbody':
            listbody = msg.Body.split("\r\n")
            temp_filename = os.path.normpath(os.path.join(outdest, f'{msg.Subject} {msg.CreationTime.strftime("%Y%m%d")} .csv'))
            b_processed = True
            with open(temp_filename, 'w', newline='') as file:
                writer = csv.writer(file)
                for row in listbody:
                    writer.writerow([row])
        if b_processed and procbox is not None:
            mail_counter += 1
            msg.Move(procbox)

    return 'Succesfully processed {} emails!'.format(mail_counter) if mail_counter > 0 else 'No emails processed'
```

runOlScript has 6 function arguments:
1. outdest - the path to the folder you wish to save the attachments or email bodies
2. filefmt - the file format you would **only** like to save. Helpful if you only want to save the .csv files across many e-mails
3. olreadfolder - the name of the outlook folder which has the e-mails you wish to process (must be a uniquely named folder!)
4. olprocessedfolder - the name of the subfolder you wish you move these sucessfully processed e-mails to (must be a uniquely named folder!)
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
For any questions do not hesitate to e-mail me at HridaiTrivedy@Gmail.com

Windows Documentation: https://docs.microsoft.com/en-us/office/vba/api/overview/outlook/object-model

Full script available for download here at the top of this page!
