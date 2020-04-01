############################################ Libraries Import #####################################
from smtplib import SMTP
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
from io import TextIOWrapper
from shutil import move
from datetime import datetime
import pyodbc, os, re, sys, csv
from smtplib import SMTPException

import os

filename = os.path.realpath(__file__)
cwd = os.path.dirname(filename)

############################################ DB Connection #########################################
Pitboss_connection_string  = ('Driver={SQL Server};'
                      'Server=sql2318-fm1-in.amr.corp.intel.com,3181;'
                      'Database=Pitboss_Station_Tracking;'
                      'Trusted_Connection=no;'
                      'UID=Pitboss_Station_Trac_so;'
                      'PWD=f4i31U35gTnVb5x;')

############################################ Get VT list ############################################
def getVT():
  try:
    conn = pyodbc.connect(eval('Pitboss_connection_string'))
    cursor = conn.cursor()
    query = cursor.execute("SELECT distinct VTName FROM Host_VT_Mapping")
    rows = [row for row in query]
    VT_list = [row[0] for row in rows]
    return (VT_list)

  except Exception as e:
    print (e)

#################################### query database to get hostnames ##################################
def getHosts(VT_lead):
  print ('VT lead is: ', VT_lead)

  try:
    conn = pyodbc.connect(eval('Pitboss_connection_string'))
    cursor = conn.cursor()
    output = cursor.execute("SELECT distinct HostName FROM Host_VT_Mapping where VTName =?",VT_lead)
    rows = [row for row in output]
    hostnames = [row[0] for row in rows]
    return hostnames

  except Exception as e:
        print(e)

################################### query database to get Software list #################################
def getSoft():
  
  try:
    conn = pyodbc.connect(eval('Pitboss_connection_string'))
    cursor = conn.cursor()
    output = cursor.execute("SELECT DISTINCT SOFTWARE_NAME FROM Health_check")
    rows = [row for row in output]
    Soft_list = [row[0] for row in rows]

    soft_chunks = []
    for i in range(0, len(Soft_list),8):
      soft_chunks.append(Soft_list[i:i +8])

    return soft_chunks

  except Exception as e:
        print(e)

########################################### Get VT Lead email #######################################
def getEmailCC(VT_lead):
  try:
      # DB connection
      conn = pyodbc.connect(eval('Pitboss_connection_string'))
      cursor = conn.cursor()
      cursor.execute('''SELECT DISTINCT Email, CC_EMAIL FROM Host_VT_Mapping where VTName = ?''', VT_lead)
      for row in cursor:
        email = [i for i in row]

      return (email[0], email[1])

  except Exception as e:
    print (e) 
################################### Get SQL table using Stored Procedure #############################
def getSQLTable(hostnames, Soft_list):
  #print ('Hosts are: ', hostnames)
  try:
      # DB connection
      conn = pyodbc.connect(eval('Pitboss_connection_string'))
      cursor = conn.cursor()

      # Query execution
      query = "exec dbo.[dashboard_converter_table] '{host}','{soft}' "
      host = ','.join(hostnames)
      soft = ','.join(Soft_list)
      query = query.format(host = host, soft = soft)
      cursor.execute(query)
      filename_csv = cwd + r'\SP_output.csv'

      with open(filename_csv,'w') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow([i[0] for i in cursor.description])
        csv_writer.writerows(cursor)
        
      return (filename_csv)

  except Exception as e:
      print(e)
############################################   CSV to HTML  ##########################################
def getHTML(file):

  def split_by(name, number):
     output = ''
     name = name.replace(' ', '_')
     for i, x in enumerate(name, 0):
         if (i+1)%number == 0:
             output += '\n'
         output += x
     return output

  table = '' 
  with open(file) as csvFile: 
      reader = csv.DictReader(csvFile, delimiter=',')    
      th = ''
      for header in reader.fieldnames:
          header = split_by(header, 10)
          th += '''<th width="200" height="80" bgcolor="#1CA8AD" 
          style="word-wrap:break-word;max-width:200px;min-width:200px">
          {}</th>'''.format(header)
      table = '<tr>{}</tr>'.format(th) 
      for row in reader:  
          table_row = '<tr>' 
          for fn in reader.fieldnames:
              if 'FAIL' in row[fn]: 
                 table_row += '<td width="200" height="50" bgcolor="#DF6B63" style="color:#FFFFFF">{}</td>'.format(row[fn]) 
              elif 'PASS' in row[fn]:
                 table_row += '<td width="200" height="50" bgcolor="#3CB083" style="color:#FFFFFF">{}</td>'.format(row[fn]) 
              else:
                 table_row += '<td width="200" height="50" bgcolor="#50DAEC">{}</td>'.format(row[fn])           
          table_row += '</tr>' 
          table += table_row
  html = """
  <head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>html title</title>
  <style type="text/css" media="screen">
  
  table{
      font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
      background-color: #000000;
      empty-cells:hide;
      Border:1px 'black';
    }
  tr {
    background-color: #4CAF50;
    color:#4CAF50
	  font-family:"Courier New";
  }
  td.cell{
      background-color: 'white';
      color:'red'
  }

  </style>
  </head>
  <html>
  <body> 
      <table style="border: black 1px;"> 
      %s
      </table>
    <br>
  </body>
  </html>""" % table
  return html

########################################### Send Mail ###################################################
def sendEmail(html,email,cc, VT):
  sender = r'sql_xcvr <sql_xcvr@intel.com>'
  receivers = [email]
  print ('TO: ',receivers)
  cc_receivers = cc.split(';')
  print ('CC: ',cc_receivers)
  name = email.replace(r'@intel.com', '')
  html_mail = """\
  <html>
    <head></head>
    <body>
      <p>Hi {name},<br>
        Please find health status on your Stations for pitboss software. <br>
      </p>
    </body>
  </html>
  """.format(name=VT)
  
  ending_html = """
  <p>
  This is auto-generated email. 
  For queries contact Mohammad.Faizal.Rudzuan@intel.com (or) PSG PVE VIS psg.pve.vis@intel.com
  </p>
  """
  html_mail = html_mail+html + ending_html
  msg = MIMEMultipart(
      "alternative", None, [MIMEText(html_mail,'html')])

  msg['Subject'] = 'Testing mail for installed software status tracking'
  msg['From'] = sender
  msg['To'] = ",".join(receivers)
  msg['CC'] = ",".join(cc_receivers)

  try:
      smtp = SMTP('smtp.intel.com')
      print(sender,  receivers + cc_receivers)
      smtp.sendmail(sender, receivers + cc_receivers, msg.as_string())       
      print ("Successfully sent email")
  except SMTPException:
      print ("Error: unable to send email")

  smtp.quit()
##################################### Main Function  ###################################
if __name__ == "__main__":

  #VT_list = getVT()
  VT_list = ['Yugeshwari']
  Soft_list = getSoft()

  for VT in VT_list:
    hosts = getHosts(VT)
    email, cc = getEmailCC(VT)
    html = ''
    for soft in Soft_list:
      file = getSQLTable(hosts, soft)
      html = html+ getHTML(file)
    
    sendEmail(html,email,cc, VT)
    with open(cwd + r'\html_table.html', 'w') as fh:
     fh.write(html)
    print ('**************************************************************************')

#########################################################################################
