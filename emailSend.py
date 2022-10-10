import sys

import win32com.client as win32
import datetime
import os


# ------------------------------------------------------------------------ #
#                   Sending Confirmation Generator                         #
#                        By Zain Zameer                                    #
#                                                                          #
# ------------------------------------------------------------------------ #

def email(file_paths):
    # get date
    date = datetime.datetime.now()

    # email generation
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = 'Sending Confirmation <<Plant>> - ' + date.strftime("%d %b %Y")

    mail.To = ""
    mail.cc = ""

    # Add the image of the Summary from the Report
    attachment = mail.Attachments.Add(os.getcwd() + "\\report.png")
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "report_img")

    desktop_path = "D:\Sending Confirmation"
    file_path = desktop_path + "\\Sending Confirmation Report " + date.strftime("%d %b %I %M %p") + ".xlsx"

    # Add the sending confirmation report Excel File
    mail.Attachments.Add(Source=file_path)
    mail.Attachments.Add(Source=file_paths[0])
    mail.Attachments.Add(Source=file_paths[1])

    mail.HTMLBody = r"""<p style="color:#0047AB;">Dear All,<br><br> Please find the attached GDN & loading 
                    confirmation for &lt;&lt;Plant&gt;&gt; on """ + date.strftime("%d %b") + """</p><br><br> 
                    <br><br>Sending Confirmation<br><br> <img src="cid:report_img"> """

    mail.display(False)
    sys.exit()
