import wx, wx.adv, email, getpass, imaplib, os, sys, re, smtplib, shutil, time, subprocess, string, socket
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

Mail_Login      = "your-gmail-login@gmail.com"
Mail_Passw      = "your-gmail-password"
Mail_Folder     = "DwlNow"
Mail_From       = "your-gmail-address@gmail.com"
Mail_To         = "your-work-address@company.com"
Mail_Server     = "imap.gmail.com"
Mail_SMTP       = "smtp.gmail.com"
Folder_toSend   = "./to_send/"
Folder_Received = './received/'
Folder_Sent     = "./sent/"

def create_menu_item(menu, label, func):
    item = wx.MenuItem(menu, -1, label)
    menu.Bind(wx.EVT_MENU, func, id=item.GetId())
    menu.AppendItem(item)
    return item 

class TaskBarIcon(wx.adv.TaskBarIcon):
    def __init__(self):
        super(TaskBarIcon, self).__init__()
        self.set_icon('./img/icon.png')
        #self.Bind(wx.EVT_TASKBAR_LEFT_DOWN, self.on_left_down)

        self.Timer = wx.Timer(self, 100)  # One ID per timer, as to not cause conflict
        self.Timer.Start(25000) # time in ms to check for files to send
        self.Bind(wx.EVT_TIMER, self.on_timer, self.Timer)

        self.Timer2 = wx.Timer(self, 200) # One ID per timer, as to not cause conflict
        self.Timer2.Start(360000) # time in ms to check for new emails to get
        self.Bind(wx.EVT_TIMER, self.check, self.Timer2)      

    def CreatePopupMenu(self):
        menu = wx.Menu()
        create_menu_item(menu, 'Get!',  self.on_get)
        create_menu_item(menu, 'Send!', self.on_send)
#       create_menu_item(menu, 'Send a file!', self.on_send_file)
        menu.AppendSeparator()
        create_menu_item(menu, 'Open Folder', self.open_folder)
        create_menu_item(menu, 'Exit', self.on_exit)
        return menu

    def check(self, event):
        self.set_icon('./img/icon-dwl.png')
        try:
            if getmail(self): subprocess.Popen('explorer ".\\received"', shell=True)
        except socket.gaierror:
           self.ShowBalloon("Status:", "Offline.", msec=5, flags=0)
        self.set_icon('./img/icon.png')
        if len(os.listdir(Folder_toSend)) != 0: check_items_to_send(self)

    def on_timer(self, event):
        if len(os.listdir(Folder_toSend)) != 0: check_items_to_send(self)

    def set_icon(self, path): self.SetIcon(wx.Icon(wx.Bitmap(path)), 'Mail Downloader')
    def on_left_down(self, event): self.check(event) 
    def on_get(self, event):  self.check(event)
    def on_send(self, event): check_items_to_send(self)
    def on_exit(self, event): wx.CallAfter(self.Destroy)
    def open_folder(self, event): subprocess.Popen('explorer .', shell=True)

    def on_send_file(self, event):
        dlg = wx.FileDialog(self, "Choose a file", './img/icon.png', "", "*.*", wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            f = open(os.path.join(self.dirname, self.filename), 'r')
            self.control.SetValue(f.read())
            f.close()
            dlg.Destroy()
    
def check_items_to_send(self):
    if len(os.listdir(Folder_toSend)) != 0:
        self.ShowBalloon("Sending", "Uploading files!!", msec=10, flags=0)
        sendmail(self)
    else:
        self.ShowBalloon("Warning:", "Nothing to send!", msec=10, flags=0)

def getmail(self):
    m = imaplib.IMAP4_SSL(Mail_Server)
    m.login(Mail_Login, Mail_Passw)
    m.select(Mail_Folder) # mailbox

    resp, items = m.search(None, "ALL") # filter using the IMAP rules (http://www.example-code.com/csharp/imap-search-critera.asp)
    items = items[0].split() # get mails id

    try: items = items[0]
    except IndexError: return

    for emailid in items:
        resp, data = m.fetch(emailid, "(RFC822)")    # fetch mail, "`(RFC822)`" means "get the whole stuff"; but you can ask for headers only, etc
        email_body = data[0][1]                      # getting the mail content
        mail = email.message_from_string(email_body) # parsing the mail content to get a mail object
        if mail.get_content_maintype() != 'multipart': continue #Check for attachments
        self.ShowBalloon("Status:", "Downloading!", msec=5, flags=0)

        for part in mail.walk(): # we use walk to create a generator so we can iterate on the parts and forget about the recursive headach
            if part.get_content_maintype() == 'multipart': continue # multipart are just containers, so we skip them
            if part.get('Content-Disposition') is None: continue    # is this part an attachment ?
            
            filename = ''.join(c for c in part.get_filename() if c in "-_.() %s%s" % (string.ascii_letters, string.digits))
            att_path = os.path.join(Folder_Received, filename)

            if not os.path.isfile(att_path):
                fp = open(att_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
        m.store(items[0], '+FLAGS', '\\Deleted')
        m.expunge()
    return(1)

def sendmail(self):
    msg = MIMEMultipart()
    msg['To'], msg['From'] = Mail_To, Mail_From

    Assunto = "UPP " + str(len(os.listdir(Folder_toSend))) + ": "
    
    for fname in os.listdir(Folder_toSend):
        path = os.path.join(Folder_toSend, fname)
        dest = os.path.join(Folder_Sent,     fname)

        if not os.path.isfile(path): continue

        part = MIMEApplication(open(path,"rb").read())
        part.add_header('Content-Disposition', 'attachment', filename=fname)
        msg.attach(part)

        shutil.move(path, dest)

        Assunto += str(fname) + "; "

    msg['Subject'] = Assunto
    part = MIMEText('text', "plain") # Mensagem no corpo do email
    part.set_payload('Attachments.') # Mensagem no corpo do email
    msg.attach(part)                 # Mensagem no corpo do email

    session = smtplib.SMTP(Mail_SMTP, 587)
    session.ehlo()
    session.starttls()
    session.ehlo
    session.login(Mail_Login, Mail_Passw)
    session.sendmail(Mail_From, Mail_To, msg.as_string())
    session.quit()

def main():
    app = wx.App()
    TaskBarIcon()
    app.MainLoop()

if __name__ == '__main__': main()