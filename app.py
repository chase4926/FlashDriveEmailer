#!/usr/bin/python

"""
Delta College Found Flash Drive Emailer

----
Requirements:
- Python 2.7
- cx_Freeze (to build)
----

Template Variables:
T_NAME = Tech's Name
T_EMAIL = Tech's Email
C_DATE = Current Date
"""


import os, time, ssl, random, re, smtplib, webbrowser, requests, lxml.html

from email.mime.text import MIMEText
from Tkinter import *
from ttk import *
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.poolmanager import PoolManager


def find_data_file(filename):
  if getattr(sys, 'frozen', False):
    # The application is frozen
    datadir = os.path.dirname(sys.executable)
  else:
    # The application is not frozen
    # Change this bit to match where you store your data files:
    datadir = os.path.dirname(__file__)
  return os.path.join(datadir, filename)



# Window width and height
width = 1050
height = 600

# URL for flash drive site
flash_drive_url = "https://applications.delta.edu/OITStudent/Lists/Flash%20Drives/All%20Check%20In.aspx"

# Variables for use with directory scraper
URL = 'https://applications.delta.edu/OITStudent/_layouts/Picker.aspx?MultiSelect=False&CustomProperty=User%3B%3B15%3B%3B%3BFalse&DialogTitle=Select People&DialogImage=%2F_layouts%2Fimages%2Fppeople.gif&PickerDialogType=Microsoft.SharePoint.WebControls.PeoplePickerDialog%2C Microsoft.SharePoint%2C Version%3D12.0.0.0%2C Culture%3Dneutral%2C PublicKeyToken%3D71e9bce111e9429c&EntitySeparator=%3B&DefaultSearch='
URL_LOGIN = 'https://applications.delta.edu/CookieAuth.dll?Logon'

payload = {
  'curl':'Z2F',
  'flags':'0',
  'forcedownlevel':'0',
  'formdir':'3',
  'trusted':'0',
  'SubmitCreds':'Log+On'}

user_agents = ("Mozilla/5.0 (Windows NT 6.1; WOW64; rv:40.0) Gecko/20100101 Firefox/40.1",
  "Mozilla/5.0 (Windows NT 6.3; rv:36.0) Gecko/20100101 Firefox/36.0",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10; rv:33.0) Gecko/20100101 Firefox/33.0",
  "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.1 Safari/537.36",
  "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.0 Safari/537.36")

ua = user_agents[random.randrange(len(user_agents))]
headers = {'user-agent': ua}


class MyAdapter(HTTPAdapter):
  # Grabbed this from a website, needed to support this SSL protocol
  def init_poolmanager(self, connections, maxsize, block=False):
    self.poolmanager = PoolManager(num_pools=connections,
      maxsize=maxsize,
      block=block,
      ssl_version=ssl.PROTOCOL_TLSv1)


def grab_directory_html(username, password, searchterm):
  # Make a copy of the payload and add the username and password to it
  full_payload = dict(payload)
  full_payload['username'] = username
  full_payload['password'] = password
  session = requests.Session()
  # Forces https to use ssl.PROTOCOL_TLSv1
  session.mount("https://", MyAdapter())
  # Obtain an authenication cookie from the login page
  session.post(URL_LOGIN, data=full_payload, headers=headers, verify=find_data_file("cacert.pem"))
  # Grab the desired page & return it
  return session.get(URL + searchterm, headers=headers).text


def extract_names_from_html(html):
  tree = lxml.html.fromstring(html)
  table = tree.find_class("ms-pickerresulttable")
  if table == []:
    # Bad username / password probably
    return "Failed"
  rows = table[0].findall('tr')
  result = []
  for row in rows:
    tds = row.find_class("ms-pb")
    if len(tds) == 0:
      continue
    result.append([
      # Name
      re.search('<\/a>.*<\/td>', lxml.html.tostring(tds[0])).group(0)[4:-5].replace("&lt;",'<').replace("&gt;",'>'),
      # Email
      tds[3].text,
      # Directory
      tds[4].text])
  return result


class App(Frame):
  def __init__(self, parent):
    Frame.__init__(self, parent)
    self.parent = parent
    
    # Entry Variables
    #   Owner
    self.owner_first_name = StringVar()
    self.owner_first_name.trace("w", lambda name, index, mode,
        sv=self.owner_first_name: self.update_owner_email())
    self.owner_last_name = StringVar()
    self.owner_last_name.trace("w", lambda name, index, mode,
        sv=self.owner_last_name: self.update_owner_email())
    self.owner_email = StringVar()
    #   Student Tech
    self.tech_first_name = StringVar()
    self.tech_first_name.trace("w", lambda name, index, mode,
        sv=self.tech_first_name: self.update_tech_email())
    self.tech_last_name = StringVar()
    self.tech_last_name.trace("w", lambda name, index, mode,
        sv=self.tech_last_name: self.update_tech_email())
    self.tech_email = StringVar()
    self.tech_email.trace("w", lambda name, index, mode,
        sv=self.tech_email: self.update_preview())
    self.tech_password = StringVar()
    # Create all of the buttons, labels, entries, etc
    self.createWidgets()
    for child in self.winfo_children(): child.grid_configure(padx=2, pady=2)
    parent.columnconfigure(0, weight=1)
    parent.rowconfigure(0, weight=1)
    self.columnconfigure(0, weight=1)
    self.columnconfigure(1, weight=2)
    self.columnconfigure(2, weight=1)
    self.rowconfigure(0, weight=1)
    self.directory_dict = {}
    self.update_preview()

  def update_owner_email(self):
    # Automatically generate the email from (first)(last)@delta.edu
    self.owner_email.set("%s%s@delta.edu" %
        (self.owner_first_name.get().lower(),
         self.owner_last_name.get().lower()))
    self.update_preview()

  def update_tech_email(self):
    # Automatically generate the email from (first)(last)@delta.edu
    self.tech_email.set("%s%s@delta.edu" %
        (self.tech_first_name.get().lower(),
         self.tech_last_name.get().lower()))
    self.update_preview()

  def get_tech_name(self):
    return "%s %s" % (self.tech_first_name.get(), self.tech_last_name.get())

  def update_preview(self):
    self.preview_text.delete(1.0, END)
    self.preview_text.insert(END, get_email_message(self.get_tech_name(), self.tech_email.get()))

  def get_preview_text(self):
    return self.preview_text.get(1.0, END)

  def create_popup_window(self, message):
    # Creates a popup window with automatic dimensions based on the message
    popup = Toplevel()
    popup.columnconfigure(0, weight=1)
    popup.rowconfigure(0, weight=1)
    popup_label = Label(popup, text=message)
    popup_label.grid(column=0, row=0, sticky=(N, W, E, S), padx=5, pady=5)
    w = popup_label.winfo_reqwidth() + 25
    h = popup_label.winfo_reqheight() + 50
    x = (self.parent.winfo_screenwidth() - w)   / 2
    y = (self.parent.winfo_screenheight() - h) / 2
    popup.geometry("%dx%d+%d+%d" % (w, h, x, y))

  def send_email_button(self):
    if not send_email(self.get_preview_text(), self.get_email_list(), self.tech_email.get(), self.tech_password.get()):
      # Set popup to notify about Authenication failure
      self.create_popup_window("Authentication failure!\nCheck the email / password.")
    else:
      # Set popup to notify Email success
      self.create_popup_window("The email has been successfully sent!")

  def get_email_list(self):
    current_indices = self.listbox.curselection()
    if len(current_indices) > 1:
      # Return all selected items
      return [self.directory_dict[x][1] for x in current_indices]
    else:
      # If one item or nothing is selected, return the manual entry field
      return [self.owner_email.get()]

  def open_flash_drives_page(self, browser=None):
    if browser:
      try:
        webbrowser.get(browser).open_new_tab(flash_drive_url)
      except:
        self.create_popup_window("Couldn't locate Firefox!\nOpened with default browser.")
        webbrowser.open(flash_drive_url)
    else:
      webbrowser.open(flash_drive_url)

  def retrieve_button(self):
    self.listbox_clear()
    if self.tech_email.get() and self.tech_password.get():
      html = grab_directory_html(self.tech_email.get(),
                                 self.tech_password.get(),
                                 self.owner_first_name.get()+self.owner_last_name.get())
      names = extract_names_from_html(html)
      if names == "Failed":
        self.create_popup_window("Failed to grab listing!\nCheck your username / password.")
      else:
        self.directory_dict = names
        self.update_listbox()
    else:
      self.create_popup_window("You need to enter a valid username/password first!")

  def update_listbox(self):
    self.listbox.delete(0, END)
    for idx, val in enumerate(self.directory_dict):
      self.listbox.insert(idx+1, val[0])

  def enable_owner_email(self):
    self.owner_email_entry["state"] = "enabled"

  def disable_owner_email(self):
    self.owner_email_entry.delete(0, END)
    self.owner_email_entry["state"] = "disabled"

  def listbox_selection_change(self, event=None):
    # Called when the user selects a different item in the listbox
    current_indices = self.listbox.curselection()
    if len(current_indices) > 1:
      # Disable manual email entry when multiple emails are selected
      self.disable_owner_email()
    elif len(current_indices) == 1:
      # Enable manual entry when only one is selected
      self.enable_owner_email()
      self.owner_email.set(self.directory_dict[current_indices[0]][1])
    else:
      # Enable manual entry when none are selected
      self.enable_owner_email()
      self.update_owner_email()

  def listbox_clear(self, event=None):
    self.listbox.selection_clear(0, self.listbox.size() - 1)
    self.enable_owner_email()
    self.update_owner_email()

  def listbox_select_all(self, event=None):
    if self.listbox.size() > 0:
      self.listbox.selection_set(0, self.listbox.size() - 1)
      if self.listbox.size() > 1:
        self.disable_owner_email()
      else:
        self.owner_email.set(self.directory_dict[0][1])

  def yview(self, *args):
    apply(self.listbox.yview, args)

  def createWidgets(self):
    self.grid(column=0, row=0, sticky=(N, W, E, S))
    
    # Frame to hold the left side of the app
    frame_left = Frame(self, borderwidth=1)
    frame_left.grid(column=0, row=0, sticky=(N, W, E, S))
    frame_left.columnconfigure(0, weight=1)
    frame_left.columnconfigure(1, weight=4)
    # Frame to hold the middle of the app
    frame_mid = Frame(self, borderwidth=1)
    frame_mid.grid(column=1, row=0, sticky=(N, W, E, S))
    frame_mid.columnconfigure(0, weight=1)
    frame_mid.rowconfigure(1, weight=1)
    # Frame to hold the right side of the app
    frame_right = Frame(self, borderwidth=1)
    frame_right.grid(column=2, row=0, sticky=(N, W, E, S))
    frame_right.columnconfigure(0, weight=1)
    frame_right.rowconfigure(1, weight=1)
    
    # ==LEFT SIDE==
    # USB Owner's First Name
    name_label1 = Label(frame_left, text="Owner's First Name:")
    name_label1.grid(column=0, row=0, sticky=W)
    
    name_entry1 = Entry(frame_left, textvariable=self.owner_first_name)
    name_entry1.grid(column=1, row=0, sticky=(W, E))
    
    # USB Owner's Last Name
    name_label2 = Label(frame_left, text="Owner's Last Name:")
    name_label2.grid(column=0, row=1, sticky=W)
    
    name_entry2 = Entry(frame_left, textvariable=self.owner_last_name)
    name_entry2.grid(column=1, row=1, sticky=(W, E))
    
    # USB Owner's Delta Email
    name_label3 = Label(frame_left, text="Owner's Delta Email:")
    name_label3.grid(column=0, row=2, sticky=W)
    
    self.owner_email_entry = Entry(frame_left, textvariable=self.owner_email)
    self.owner_email_entry.grid(column=1, row=2, sticky=(W, E))
    
    # Seperator
    sep = Separator(frame_left, orient=HORIZONTAL)
    sep.grid(column=0, row=3, sticky=(W, E), columnspan=2)
    
    # Student Tech's First Name
    name_label4 = Label(frame_left, text="Tech's First Name:")
    name_label4.grid(column=0, row=4, sticky=W)
    
    name_entry4 = Entry(frame_left, textvariable=self.tech_first_name)
    name_entry4.grid(column=1, row=4, sticky=(W, E))
    
    # Student Tech's Last Name
    name_label5 = Label(frame_left, text="Tech's Last Name:")
    name_label5.grid(column=0, row=5, sticky=W)
    
    name_entry5 = Entry(frame_left, textvariable=self.tech_last_name)
    name_entry5.grid(column=1, row=5, sticky=(W, E))
    
    # Student Tech's Delta Email
    name_label6 = Label(frame_left, text="Tech's Delta Email:")
    name_label6.grid(column=0, row=6, sticky=W)
    
    name_entry6 = Entry(frame_left, textvariable=self.tech_email)
    name_entry6.grid(column=1, row=6, sticky=(W, E))
    
    # Student Tech's Delta Password
    name_label7 = Label(frame_left, text="Tech's Delta Password:")
    name_label7.grid(column=0, row=7, sticky=W)
    
    name_entry7 = Entry(frame_left, show="*", textvariable=self.tech_password)
    name_entry7.grid(column=1, row=7, sticky=(W, E))
    
    # ==MIDDLE==
    # Retrieve button (Retrieves listings for the listbox)
    retrieve_button = Button(frame_mid, text="Retrieve Listing", command=self.retrieve_button)
    retrieve_button.grid(column=0, row=0, sticky=W)
    
    # Listbox of possible owners
    listbox_frame = Frame(frame_mid)
    listbox_frame.grid(column=0, row=1, sticky=(N, W, E, S))
    listbox_frame.columnconfigure(0, weight=1)
    listbox_frame.rowconfigure(0, weight=1)
    
    scrollbar = Scrollbar(listbox_frame, orient=VERTICAL, command=self.yview)
    scrollbar.grid(column=1, row=0, sticky=(N, S))
    
    self.listbox = Listbox(listbox_frame, width=40, selectmode="multiple", yscrollcommand=scrollbar.set)
    self.listbox.grid(column=0, row=0, sticky=(N, W, E, S))
    self.listbox.bind("<<ListboxSelect>>", self.listbox_selection_change)
    
    
    # ==RIGHT SIDE==
    # Email preview
    preview_label = Label(frame_right, text="Email Preview:")
    preview_label.grid(column=0, row=0, sticky=W)
    
    self.preview_text = Text(frame_right, wrap="word", width=40)
    self.preview_text.grid(column=0, row=1, sticky=(N, W, E, S))
    
    for child in frame_left.winfo_children(): child.grid_configure(padx=4, pady=4)
    for child in frame_mid.winfo_children(): child.grid_configure(padx=4, pady=4)
    for child in frame_right.winfo_children(): child.grid_configure(padx=4, pady=4)
    
    # ==BOTTOM==
    
    # Button to open flash drive website
    flash_drive_button1 = Button(self, text="Open Flash Drive Page", command=self.open_flash_drives_page)
    flash_drive_button1.grid(column=0, row=1, stick=W)
    
    flash_drive_button1 = Button(self, text="Open Flash Drive Page in Firefox", command=lambda: self.open_flash_drives_page('firefox'))
    flash_drive_button1.grid(column=0, row=1, stick=E)
    
    # Select All button
    select_button = Button(self, text="Select All", command=self.listbox_select_all)
    select_button.grid(column=1, row=1, sticky=W)
    # Clear button
    clear_button = Button(self, text="Clear Selection", command=self.listbox_clear)
    clear_button.grid(column=1, row=1, sticky=E)
    
    # Button to close the window
    close_button = Button(self, text="Close", command=self.quit)
    close_button.grid(column=2, row=1, sticky=E)
    # Button to send the email
    send_button = Button(self, text="Send Email", command=lambda: self.send_email_button())
    send_button.grid(column=2, row=1, sticky=W)


def get_email_message(tech_name, from_email):
  with open("template.txt") as f:
    body = f.read()
  body = body.replace("T_NAME", tech_name)
  body = body.replace("T_EMAIL", from_email)
  body = body.replace("C_DATE", time.strftime("%m/%d/%Y"))
  return body


def send_email(message, to_email, from_email, password):
  msg = MIMEText(message)
  msg['Subject'] = "Missing Flash Drive Two Month Notice"
  msg['From'] = from_email
  if type(to_email) != list:
    # If sending to a single recipient, add the To: header
    msg['To'] = to_email
    to_email = [to_email]
  server = smtplib.SMTP('smtp.office365.com', 587)
  server.ehlo()
  server.starttls()
  try:
    server.login(from_email, password)
    server.sendmail(from_email, to_email, msg.as_string())
  except smtplib.SMTPAuthenticationError:
    # Authenication failed
    server.quit()
    return False
  else:
    # Message sent successfully
    server.quit()
    return True


def main():
  root = Tk()
  root.title("Flash Drive Emailer")
  x = (root.winfo_screenwidth() - width)   / 2
  y = (root.winfo_screenheight() - height) / 2
  root.geometry("%dx%d+%d+%d" % (width, height, x, y))
  app = App(root)
  root.mainloop()


if __name__ == "__main__":
  main()
