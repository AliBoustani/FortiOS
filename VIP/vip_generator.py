from openpyxl import load_workbook
from netmiko import ConnectHandler
import getpass



#Specify your excel document! This is my test sample
wb = load_workbook(filename='MyDoc.xlsx')

#The sheet name in your excel file
ws = wb['Sheet1']

#Determining the row & column size
maxrow = ws.max_row
maxcols = ws.max_column

#Enter your FortiGate IP address and credentials
host = raw_input("Please Enter your FortiGate device IP address: ")
username = raw_input("Please Enter your FortiGate username: ")
password = getpass.getpass()


#Creating connection profile for your device
My_fortigate = {
    'device_type': 'fortinet',
    'ip': host,
    'username': username,
    'password': password,
    }


#Establishing SSH connection
net_connect = ConnectHandler(**My_fortigate)

#The function!
def fg_vip_conf(vip_name, intip, extip, port):

    #VIP script
    commands_vip = ['config firewall vip', 'edit '+ vip_name+port,
'set extip '+ extip ,'set extintf wan1','set mappedip '+ intip,'set portforward enable','set protocol tcp','set extport '+port,'set mappedport '+port, 'end']
    
    #VIP Group script
    commands_vipgrp = ['config firewall vipgrp', 'edit '+ vip_name,'set interface wan1', 'append member ' + vip_name+port, 'end']

    output_vip = net_connect.send_config_set(commands_vip)
    output_vipgrp = net_connect.send_config_set(commands_vipgrp)


   # print(output_vip)
   # print(output_vipgrp)




#Applying th function for each column in each row
#It starts from second raw! First raw is just about each column informations.
for j in range (2, maxrow+1):
    dic1 = dict((i, str(ws.cell(row=j, column=i).value)) for i in range(1, maxcols+1))
    for k in range(4, maxcols+1):  #Considering the specified ports start at column 4 !
      if dic1[k] != 'None':
          fg_vip_conf(dic1[1], dic1[2], dic1[3], dic1[k])
