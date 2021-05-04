import sys
import os
import argparse
import csv
import win32com.client

def _right(s, amount):
    return s[-amount:]

def _scriptOutput(s, gui_entry):
    if gui_entry:
        return s
    else:
        sys.exit(s)

def run_ol_Script(outdest, filefmt, olreadfolder, olprocessedfolder, gui_entry, proc):
    outdest = os.path.normpath(outdest)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = None
    for folder in outlook.Folders:
        try:
            inbox = folder.Folders['Inbox'].Folders[olreadfolder]
            break
        except Exception as e:
            print(e)
    if inbox is None:
        sys.exit(f'No Folder {olreadfolder} found!!! Exiting.')
    procbox = olprocessedfolder
    if procbox is not None:
        procbox = None
        for folder in outlook.Folders:
            try:
                procbox = folder.Folders['Inbox'].Folders[olreadfolder].Folders[olprocessedfolder]
                break
            except Exception as e:
                print(e)
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

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-ola","--olatt",nargs='?',default=False)
    parser.add_argument("-olb","--olbody",nargs='?',default=False)
    parser.add_argument("-out","--outdest",default="")
    parser.add_argument("-olf1","--olfolder",default="")
    parser.add_argument("-olf2","--olprocfolder",default=None)
    parser.add_argument("-typ","--olfiletype",default="blank")
    args = parser.parse_args()
    
    b_olatt = True if args.olatt is None else False
    b_olbody = True if args.olbody is None else False
    
    if (not b_olatt and not b_olbody):
        sys.exit('No process choice made, choose between ol attachments saver (--olatt) and ol mail body saver (--olbody)!')
    if args.outdest == '':
        sys.exit('No out destination defined using --outdest defined!')
    if args.olfolder == '':
        sys.exit('No outlook folder to search mails for defined using --olfolder!')
    if args.olprocfolder == '':
        sys.exit('No outlook folder to move processed mails to defined using --olprocfolder!')
    proc = 'olatt' if b_olatt else 'olbody' if b_olbody else ''

    run_ol_Script(args.outdest, args.olfiletype, args.olfolder, args.olprocfolder, False, proc)