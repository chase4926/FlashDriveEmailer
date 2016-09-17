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

import time, smtplib, webbrowser

from email.mime.text import MIMEText
from Tkinter import *
from ttk import *


# Window width and height
width = 700
height = 400
# URL for flash drive site
flash_drive_url = "https://applications.delta.edu/OITStudent/Lists/Flash%20Drives/All%20Check%20In.aspx"


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
    self.columnconfigure(1, weight=1)
    self.rowconfigure(0, weight=1)
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
    if not send_email(self.get_preview_text(), self.owner_email.get(), self.tech_email.get(), self.tech_password.get()):
      # Set popup to notify about Authenication failure
      self.create_popup_window("Authentication failure!\nCheck the email / password.")
    else:
      # Set popup to notify Email success
      self.create_popup_window("The email has been successfully sent!")

  def open_flash_drives_page(self, browser=None):
    if browser:
      try:
        webbrowser.get(browser).open_new_tab(flash_drive_url)
      except:
        self.create_popup_window("Couldn't locate Firefox!\nOpened with default browser.")
        webbrowser.open(flash_drive_url)
    else:
      webbrowser.open(flash_drive_url)

  def createWidgets(self):
    self.grid(column=0, row=0, sticky=(N, W, E, S))
    
    # Frame to hold the left side of the app
    frame_left = Frame(self, borderwidth=1)
    frame_left.grid(column=0, row=0, sticky=(N, W, E, S))
    frame_left.columnconfigure(0, weight=1)
    frame_left.columnconfigure(1, weight=4)
    # Frame to hold the right side of the app
    frame_right = Frame(self, borderwidth=1)
    frame_right.grid(column=1, row=0, sticky=(N, W, E, S))
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
    
    name_entry3 = Entry(frame_left, textvariable=self.owner_email)
    name_entry3.grid(column=1, row=2, sticky=(W, E))
    
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
    
    # ==RIGHT SIDE==
    # Frame to hold the email preview
    preview_label = Label(frame_right, text="Email Preview:")
    preview_label.grid(column=0, row=0, sticky=W)
    
    self.preview_text = Text(frame_right, wrap="word", width=40)
    self.preview_text.grid(column=0, row=1, sticky=(N, W, E, S))
    
    for child in frame_left.winfo_children(): child.grid_configure(padx=4, pady=4)
    for child in frame_right.winfo_children(): child.grid_configure(padx=4, pady=4)
    
    # ==BOTTOM==
    
    # Button to open flash drive website
    flash_drive_button1 = Button(self, text="Open Flash Drive Page", command=self.open_flash_drives_page)
    flash_drive_button1.grid(column=0, row=1, stick=W)
    
    flash_drive_button1 = Button(self, text="Open Flash Drive Page in Firefox", command=lambda: self.open_flash_drives_page('firefox'))
    flash_drive_button1.grid(column=0, row=1, stick=E)
    
    # Button to close the window
    close_button = Button(self, text="Close", command=self.quit)
    close_button.grid(column=1, row=1, sticky=E)
    # Button to send the email
    send_button = Button(self, text="Send Email", command=lambda: self.send_email_button())
    send_button.grid(column=1, row=1, sticky=W)


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
  msg['To'] = to_email
  server = smtplib.SMTP('smtp.office365.com', 587)
  server.ehlo()
  server.starttls()
  try:
    server.login(from_email, password)
    server.sendmail(from_email, [to_email], msg.as_string())
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
