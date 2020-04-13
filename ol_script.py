import sys, os.path, argparse, csv
import win32com.client
from win32com.client.gencache import EnsureDispatch as Dispatch
import datetime

def _right(s, amount):
    return s[-amount:]

def _scriptOutput(s, guiEntry):
    if guiEntry:
        return s
    else:
        sys.exit(s)

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