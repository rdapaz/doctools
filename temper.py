import re

data = """
CPM Prod Server Migrations
AWS (Low Change Monday 9th July) - Completed
Server  Schedule    VPC Notes
PER2HWDMSR01    8:00am 9/7/2018 Production  250GB
PER2MEAP01  8:00am 9/7/2018 Production  260GB

Each server will go through the pre-flight checks:
1.  Backups Completed
2.  RDP Enabled (Cache admin credentials)
3.  AV Disabled
4.  CD/DVD Disconnected
5.  Page File not System Managed
6.  Root Free Disk Space > 12GB
7.  Shutdown VM
Instance types as per spreadsheet created in AWS once sync has completed.
Checks:
1.  Server can be RDPed to
2.  DNS Records updated
3.  Antivirus activated again
4.  Contact tech owner for testing
5.  Configure Backups
Estimated return to service: Tuesday 10th July 2018 @ 8:00am


 
V2V Malaga (Low Change Friday 13th July) - Completed
Server  Schedule    Notes
PER1TFC02   8:00am 13/7/2018    136GB
PER1TFS02   8:00am 13/7/2018    110GB
PER2ICA01   8:00am 13/7/2018    65GB
PER2RCA01   8:00am 13/7/2018    40GB

1.  Backups completed
2.  Confirm local login credentials
3.  Note the existing IP addressing information and NIC type
4.  Each server shutdown gracefully
5.  V2V Converter job setup for each server
a.  Destination LUN for each needs to be nominated
6.  Once completed, delete the NIC
7.  Recreate the NIC (not connected)
8.  Boot the VM
9.  Assign IP addressing information
10. Shutdown the VM
11. Connect the NIC
12. Boot VM
13. Advise tech owner for testing
14. Configure Backups
Estimated return to service: Saturday 14th July 2018 @ 8:00am
 
 
V2V Malaga (Major Change Thursday 19th)
Server  Schedule    Size    Notes
PER1CBAR01  6:00pm 18/7/2018    45GB    Linked to PER1SQL02 & 04
PER2ARCGISVD01  6:00pm 18/7/2018    124GB   ArcGIS Virtual Desktop
PER2CDC01   6:00pm 18/7/2018    46GB    Remote Access Gateway Front End
PER2CSH01   6:00pm 18/7/2018    76GB    Redundant Pair Citrix Sessions Hosts (Garreth)
PER2RRAS01  6:00pm 18/7/2018    42GB    Azure RRAS Server
PER2NPS01   6:00pm 18/7/2018    44GB    MS RADIUS Front End (*IP Addressing)

1.  Backups completed
2.  Confirm local login credentials
3.  Note the existing IP addressing information and NIC type
4.  Each server shutdown gracefully
5.  V2V Converter job setup for each server
a.  Destination LUN for each needs to be nominated
6.  Once completed, delete the NIC
7.  Recreate the NIC (not connected)
8.  Boot the VM
9.  Assign IP addressing information
10. Shutdown the VM
11. Connect the NIC
12. Boot VM
13. Advise tech owner for testing
14. Configure Backups
NIC Information 19/7
Estimated return to service: Thursday 18th July 2018 @ 11:00pm

Server  Schedule    Size    Notes
PER1SVN01   6:00pm 25/7/2018    1TB Schneider Code
PER2CDC02   6:00pm 25/7/2018    46GB    Remote Access Gateway Front End
PER2CSH02   6:00pm 25/7/2018    76GB    Redundant Pair Citrix Sessions Hosts (Garreth)
PER1UFS01   6:00pm 25/7/2018    45GB    Canon UniFlow (Derik)
PER2AADS01  8:00am 25/7/2018    85GB    Azure AD Sync

NIC Information 25/7
Estimated return to service: Wednesday 25th July 2018 @ 11:00pm

1.  Backups completed
2.  Confirm local login credentials
3.  Note the existing IP addressing information and NIC type
4.  Each server shutdown gracefully
5.  V2V Converter job setup for each server
a.  Destination LUN for each needs to be nominated
6.  Once completed, delete the NIC
7.  Recreate the NIC (not connected)
8.  Boot the VM
9.  Assign IP addressing information
10. Shutdown the VM
11. Connect the NIC
12. Boot VM
13. Advise tech owner for testing
14. Configure Backups

Decommission - TBA
Server  Notes
PER1AMS01   Azure Migration Server
PER2AWSM01  AWS Migration Server
PER2DC03    DC Decom before Belmont is closed
PER2NPSTST01    ? NPS Test Server AP/GV
PER2NPS02   Azure MFA RADIUS (Test) Powered off – waiting deletion confirmation AP/GV
PER2SRM10   Site Replication
PER2VC01    vSphere Front End
PER2VC02    vSphere Back End
PER2VPSC01  VM Platform Service Controller
PER2VRA01   vSphere Replication


 
NIC Configurations – 19th July 2018
PER1CBAR01 (LUN 113)
Claud.Lin@citicpacificmining.com 0401 207 859
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:03:8a
IP address  172.22.4.112
Subnet Mask 255.255.255.0
Default Gateway 172.22.4.5
DNS Servers 172.22.1.12 & 10

PER2ARCGISVD01 (LUN 118)
Chris.Brown@citicpacificmining.com 0408 950 363
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:59:f4
IP address  172.22.4.55
Subnet Mask 255.255.255.0
Default Gateway 172.22.4.5
DNS Servers 172.22.12 & 10

PER2CDC01 (LUN 106)
Alexander.Potapov@citicpacificmining.com 0481 367 060
Garreth.Holdaway@citicpacificmining.com 0438 610 432
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:5c:2a
IP address  172.22.4.95
Subnet Mask 255.255.255.0
Default Gateway 172.22.4.5
DNS Servers 172.22.1.12 & 11

PER2CSH01 (LUN 106)
Alexander.Potapov@citicpacificmining.com 0481 367 060
Garreth.Holdaway@citicpacificmining.com 0438 610 432
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:5c:2c
IP address  172.22.4.97
Subnet Mask 255.255.255.0
Default Gateway 172.22.4.5
DNS Servers 172.22.1.12 & 11

PER2RRAS01 (LUN 106)
gabriel.van@citicpacificmining.com 0439 943 738
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:5c:1e
IP address  172.22.4.80
Subnet Mask 255.255.255.0
Default Gateway 172.22.4.5
DNS Servers 172.22.1.12 & 11

PER2NPS01 (LUN 106)
yunfei.li@citicpacificmining.com 0439 088 494
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:7f:91
IP address  172.22.1.21
Subnet Mask 255.255.255.0
Default Gateway 172.22.1.5
DNS Servers 172.22.1.12 & 11

 
NIC Configurations 25th July 2018
PER1SVN01
Jonathan.churchman@citicpacificmining.com Mobile #
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:6b:ee
IP address  172.22.1.129
Subnet Mask 255.255.255.0
Default Gateway 172.22.1.5
DNS Servers 172.22.1.12 & 11

PER2CDC02
Alexander.Potapov@citicpacificmining.com 0481 367 060
Garreth.Holdaway@citicpacificmining.com 0438 610 432
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:5c:2b
IP address  172.22.4.96
Subnet Mask 255.255.255.0
Default Gateway 172.22.4.5
DNS Servers 172.22.1.12 & 11

PER2CSH02
Alexander.Potapov@citicpacificmining.com 0481 367 060
Garreth.Holdaway@citicpacificmining.com 0438 610 432
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:5c:2d
IP address  172.22.4.98
Subnet Mask 255.255.255.0
Default Gateway 172.22.4.5
DNS Servers 172.22.1.12 & 11

PER1UFS01
Canon
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:25:0d
IP address  172.22.1.35
Subnet Mask 255.255.255.0
Default Gateway 172.22.1.5
DNS Servers 172.22.1.12 & 11

PER2AADS01
Alexander.Potapov@citicpacificmining.com 0481 367 060
Network Adapter VMXNET 3 (Prod_Servers_20)
MAC Address 00:50:56:a5:7e:22
IP address  172.22.1.145
Subnet Mask 255.255.255.0
Default Gateway 172.22.1.5
DNS Servers 172.22.1.12 & 11

"""


rex = re.compile(r'(\w+\.\w+@\w+\.\w+\s+[\d\s]+)')


entries = []
for entry in rex.findall(data):
    entries.append(entry)

entries = list(set(entries))
rex2 = re.compile(r'(\w+\.\w+@\w+\.\w+)\s+([\d\s]+)')
for entry in entries:
    m = rex2.search(entry)
    if m:
        email = m.group(1).strip()
        phone = m.group(2).strip()
        name = " ".join(s for s in email.split('@')[0].split('.'))
        print(name, '|', phone, email, sep="|")

