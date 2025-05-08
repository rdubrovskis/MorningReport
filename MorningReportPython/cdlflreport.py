from email import header
import smtplib
import requests
import sys
import datetime
import json
import configparser
#Note: A Python package for JIRA is available that would've streamlined this without needing to mess with JSON manually
#However for some reason the VPN on my workstation is not allowing it to import correctly, so I had to just use JSON

from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

#Manually setting argument for testing purposes
#sys.argv = ['morning']

#Creating a variable with the argument value
jobtype=str

#Check the arguments passed into the script execution. If no arguments were passed, set job type to 'morning'
if len(sys.argv) > 1:
    jobtype = sys.argv[1]
else:
    jobtype = 'morning'

#Parse .ini file for settings
config=configparser.ConfigParser()
config.read('cdlflreport.ini')

# Set header parameters for JIRA Rest request. Authorization field must include JIRA API token

bearer = 'Bearer ' + config.get('main','token')
headers = {
    "Authorization": bearer,
    "Accept": "application/json"
}

# Project Key
projectKey = 'ABCD'

startID = config.get('main','startid')

#Setting the REST URL for JIRA. The included JQL asks for all tickets with ID greater than the ticket ID recorded in the .ini, ordered by ID descending, including only the fields relevant for the email
url="https://jira.atlassian.ITCOMPANY.com/rest/api/2/search?jql=project=ABCD AND Key>%s ORDER BY Key DESC&fields=key,summary,status,customfield_23110" % (startID)

# Send request and get response
response = requests.get(url,headers=headers)


# Decode Json string to Python
json_data = json.loads(response.text)

#Defining some strings to adjust the e-mail contents based on morning or afternoon report being sent
status=str

if jobtype == 'afternoon':
    status = 'status'
else:
    status = 'list'
    

#Build first part of HTML for e-mail
html = """
<html>
  <head>
    <style type="text/css">
         body
         {
             margin: 0px;
             font-size: 16px;
             font-family: Arial, sans-serif;
             color:black;
         }

        </style>
        <META HTTP-EQUIV="Content-Type" Content="application/vnd.ms-excel; charset=UTF-8">
        <style>

         td.issuekey,
         td.issuetype,
         td.status {
             mso-style-parent: "";
             mso-number-format: \@;
             text-align: left;
         }
         br
         {
             mso-data-placement:same-cell;
         }

         td
         {
             vertical-align: top;
         }
        </style>
  </head>
  <body>
    <p>Good __jobtype__,<br>
      <br />
      Below is the __status__ of tickets created by First Line today:</p>
      <table border="1" cellpadding="3" cellspacing="1" width="100%">
        <thead>
            <tr>
                <th scope="col class="colHeaderLink headerrow-issuekey">Key</th>
                <th scope="col" class="colHeaderLink headerrow-customfield_23110">Squad</th>
                <th scope="col" class="colHeaderLink headerrow-summary">Summary</th>
                <th scope="col" class="colHeaderLink headerrow-status">Status</th>
            </tr>
        </thead>
        <tbody>"""

html = html.replace('__jobtype__',jobtype)
html = html.replace('__status__', status)

currentDate = datetime.date.today().strftime('%Y-%m-%d')
subjectLine = '%s of tickets Created By First Line - %s' % (status.capitalize(), currentDate)

#Parsing issues retrieved from the REST response into a list
tRows =[]
issues = json_data["issues"]

#In case the Squad field is null, set it to a blank string
def convert_None(d):
     return '' if d is None else d["value"]

 #Loop through the list of issues and create a list of strings containing HTML for table rows with their information
for i in issues:
    fields = i['fields']
    squad = convert_None(fields["customfield_23110"])
    tRow = """<tr class="issuerow">
                    <td class="issuekey"><a href="https://jira.atlassian.teliacompany.net/browse/%s">%s</a></td>
                    <td class="customfield_23110">%s</td>
                    <td class="summary">%s</td>
                    <td class="status">%s</td>
              </tr>
    
    """ % (i["key"], i["key"], squad, fields["summary"],fields["status"]["name"])
    tRows.append(tRow)

#Append all the rows generated above to the HTML string
for r in tRows:
    html += r

#Append the last part of the HTML string
html += """</tbody>
</table>
<p>Regards,<br>
      <br>First Line Support Team</p>
</body>
</html>
"""

#Create Email object
relay = config.get('main','relay')

class EmailBuilder(object):
    port = 25
    smtp_server = relay
    sender_email = "FirstLineTeam@ITCOMPANY.com"
    receiver_email = "Stakeholders@ITCOMPANY.com"
    message = MIMEMultipart()
    message["Subject"] = subjectLine
    message["From"] = sender_email
    message["To"] = receiver_email
    part1 = MIMEText(html, "html")
    message.attach(part1)

#This prints the contents of the HTML string into an HTML file. Used for testing and verifying the final contents
# File_object = open(r"email.html", "w+")
# File_object.write(html)
# File_object.close()

#Send the email object to relay via SMTP
with smtplib.SMTP(relay, 25) as server:
    server.sendmail(
        EmailBuilder.sender_email, EmailBuilder.receiver_email, EmailBuilder.message.as_string()
    )

#Getting the current day value to update the .ini file on weekends without sending an afternoon e-mail
currentDay = datetime.date.weekday(datetime.date.today())
print (currentDay)

#Update .ini with latest ticket ID if afternoon job or running on a weekend so that next day's morning report starts after the latest reported ticket
if jobtype == 'afternoon' or currentDay > 4:
    config.set('main','startID', issues[0]['key'])
    with open('cdlflreport.ini', 'w') as configfile:
        config.write(configfile)


exit()


