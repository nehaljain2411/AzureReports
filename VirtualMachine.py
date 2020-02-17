import os
import traceback
import xlwt
import datetime
import boto3
from azure.common.credentials import ServicePrincipalCredentials
from azure.mgmt.resource import ResourceManagementClient
from azure.mgmt.network import NetworkManagementClient
from azure.mgmt.compute import ComputeManagementClient
from azure.mgmt.compute.models import DiskCreateOption
from msrestazure.azure_exceptions import CloudError
from azure.mgmt.resource import SubscriptionClient
from botocore.exceptions import ClientError
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
today_date=datetime.date.today()
AWS_PROFILE='myprofile1'
email_from = 'nehal.jain@duffandphelps.com'
email_to = ['nehal.jain@duffandphelps.com']
boto3.setup_default_session(profile_name=AWS_PROFILE,region_name='us-east-1')
sts = boto3.client('sts')

wb = xlwt.Workbook(encoding="utf-8")
sheet3=wb.add_sheet("sheet3")

vms = []

sheet3.write(0,0,"VmName")
sheet3.write(0,1,"Subscription")
sheet3.write(0,2,"Resource Group")
sheet3.write(0,3,"Location")
sheet3.write(0,4,"Provisioning State")
sheet3.write(0,5,"OS Name")
sheet3.write(0,6,"OS Version")
sheet3.write(0,7,"Private IP")
sheet3.write(0,8,"Public IP")
sheet3.write(0,9,"VM Size")
sheet3.write(0,10,"Disk Name")
sheet3.write(0,11,"Disk Size (GB)")
sheet3.write(0,12,"Admin Username")
sheet3.write(0,13,"ApplicationName")
sheet3.write(0,14,"BusinessOwner")
sheet3.write(0,15,"Environment")
sheet3.write(0,16,"ServiceLine")
sheet3.write(0,17,"TechnologyOwner")

def get_credentials():    
    subscription_id = "3a130925-d7da-48a2-b023-80148f77c31d"
    credentials = ServicePrincipalCredentials(
        client_id="8c33058c-0ca8-49c4-ba75-8207eec88153",
        secret="b6j=-XrH8XX4UtFuKjt+3Wlsci:tEsDp",
        tenant="781802be-916f-42df-a204-78a2b3144934",
    )
    return credentials
def run_example():
    subscription_id = "3a130925-d7da-48a2-b023-80148f77c31d"
    
    credentials = get_credentials()
    resource_client = ResourceManagementClient(credentials, subscription_id)
    compute_client = ComputeManagementClient(credentials, subscription_id)
    network_client = NetworkManagementClient(credentials, subscription_id)
    resource_client = ResourceManagementClient(credentials, subscription_id)    
    body_text =["\r\r\n Resource Groups \r\r\n "]
    body_text2 =["Please find the attached list of Virtual Machines in Subscription and Resource Groups  \r\r\n Subscriptions \r\r\n "]
    try:
        k=0
        subscriptionClient = SubscriptionClient(credentials)
        for subscription in subscriptionClient.subscriptions.list():
            vmcount1 = []
            sub_id = subscription.subscription_id
            resource_client1 = ResourceManagementClient(credentials, sub_id)
            compute_client1 = ComputeManagementClient(credentials, sub_id)
            network_client1 = NetworkManagementClient(credentials, sub_id)
            groups = resource_client1.resource_groups.list()
            for group in groups:
                vmcount = []
                for vm in compute_client1.virtual_machines.list(group.name):
                    ips =[]
                    pub =[]
                    ni_reference = vm.network_profile.network_interfaces[0]
                    ni_reference = ni_reference.id.split('/')
                    ni_group = ni_reference[4]
                    ni_name = ni_reference[8]

                    net_interface = network_client1.network_interfaces.get(ni_group, ni_name)
                    try:
                        ip_reference = net_interface.ip_configurations[0].public_ip_address
                        ip_reference = ip_reference.id.split('/')
                        ip_group = ip_reference[4]
                        ip_name = ip_reference[8]
        
                        public_ip = network_client1.public_ip_addresses.get(ip_group, ip_name)
                        sheet3.write(k+1,8,public_ip.ip_address)
                    except: pass
                    try:
                        private_ip = network_client1.network_interfaces.get(ni_group, ni_name).ip_configurations[0].private_ip_address
                        sheet3.write(k+1,7,private_ip)
                    except: pass
                    try:
                        vm_state = compute_client1.virtual_machines.instance_view(resource_group_name=group.name, vm_name=vm.name)
                        if vm_state.statuses[1].code == 'PowerState/running':
                            sheet3.write(k+1,4,"Running")
                        else: sheet3.write(k+1,4,"Stopped")
                    except: pass
                    try: sheet3.write(k+1,5,vm_state.os_name)    
                    except: pass
                    try: sheet3.write(k+1,6,vm_state.os_version)    
                    except: pass
                    vmcount1.append(vm.name)
                    vmcount.append(vm.name)
                    vms.append(vm.name)
                    sheet3.write(k+1,0,vm.name)
                    sheet3.write(k+1,1,subscription.display_name)
                    sheet3.write(k+1,2,group.name)
                    sheet3.write(k+1,3,vm.location)
                    sheet3.write(k+1,10,vm.storage_profile.os_disk.name)
                    sheet3.write(k+1,11,vm.storage_profile.os_disk.disk_size_gb)
                    sheet3.write(k+1,9,vm.hardware_profile.vm_size)
                    try: sheet3.write(k+1,12,vm.os_profile.admin_username)
                    except: pass
                    try:
                            for tag,tag1 in vm.tags.items():
                                try:
                                    if tag == 'BusinessOwner': busname = tag1
                                except: busname = "UNKNOWN"
                                try:
                                    if tag == 'ApplicationName': appname = tag1
                                except: appname = "UNKNOWN"
                                try:
                                    if tag == 'Environment': envname = tag1
                                except: envname = "UNKNOWN"
                                try:
                                    if tag == 'ServiceLine': servname = tag1
                                except: servname = "UNKNOWN"
                                try:
                                    if tag == 'TechnologyOwner': techname = tag1
                                except: techname = "UNKNOWN"
                            sheet3.write(k+1,13,appname)
                            sheet3.write(k+1,14,busname)
                            sheet3.write(k+1,15,envname)
                            sheet3.write(k+1,16,servname)
                            sheet3.write(k+1,17,techname)
                    except: pass
                    k =k+1
                if len(vmcount) > 0:
                    body_text.append("Total VMs Count in Resource Group "+ group.name +": "+str(len(vmcount)))
            if len(vmcount1) > 0:
                body_text2.append("Total VMs Count in Subscription "+ subscription.display_name +": "+str(len(vmcount1)))
        print("Total VMs Count : "+str(len(vms)))
        body_text1= "\r\r\n ".join(body_text)
        body_text3= "\r\r\n ".join(body_text2)
        body_text4 = body_text3 + body_text1 + "\r\r\n Total Number of Virtual Machines : "+str(len(vms))
        #print(body_text1)
        day = today_date.strftime("%B_%Y")
        file_name = 'Azure_VirtualMachine'+day+'.xls'
        wb.save(file_name)
    except CloudError:
        print('A VM operation failed:\n{}'.format(traceback.format_exc()))

    try:
        prim_assume_itbn = sts.assume_role(
        RoleArn='arn:aws:iam::104436734642:role/3-Prd-Analyst-Access',
        RoleSessionName='ITBN',
        DurationSeconds=1800,
        )

        Prim_ITBN_RoleAccessKeyId = prim_assume_itbn["Credentials"]["AccessKeyId"]
        Prim_ITBN_RoleSecretAccessKey = prim_assume_itbn["Credentials"]["SecretAccessKey"]
        Prim_ITBN_RoleSessionToken = prim_assume_itbn["Credentials"]["SessionToken"]

        
        CHARSET = "utf-8"
        #Sending Email for Unused Resources
        msg = MIMEMultipart()
##        body_text = (
##            "Attached herewith is the latest Azure Virtual Machine Inventory \r\r\n"
##                      "Total Number of Virtual Machines: " + str(len(vms)) + " \r\r\n"
##                     )
        #html = str(sys.argv[3])
        msg['Subject'] = "Azure Virtual Machine Inventory List"
        msg['From'] = email_from
        msg['To'] = ', '.join(email_to)
        body = MIMEText(body_text4.encode(CHARSET), 'plain', CHARSET)
        msg.attach(body)
        # What a recipient sees if they don't use an email reader
        msg.preamble = 'Multipart message.\n'                                                                 
        part = MIMEApplication(open(file_name, "rb").read())
        part.add_header('Content-Disposition', 'attachment', filename=file_name)
        part.add_header('Content-Type', 'application/vnd.ms-excel; charset=UTF-8')
        msg.attach(part)
        # Create a new SES resource and specify a region.
        ses = boto3.client('ses',
            aws_access_key_id=Prim_ITBN_RoleAccessKeyId,
            aws_secret_access_key=Prim_ITBN_RoleSecretAccessKey,
            aws_session_token=Prim_ITBN_RoleSessionToken,
            region_name='us-east-1',
            verify=False)
        
        # Try to send the email.
        #Provide the contents of the email.
        response=ses.send_raw_email(
            Source=email_from,
            Destinations=email_to,
            RawMessage={
                'Data': msg.as_string(),
            }
        )      
        # Display an error if something goes wrong.	
    except ClientError as e:
        print(e.response['Error']['Message'])
    else:
        print("Email sent! Message ID:"),
        print(response['MessageId'])

if __name__ == "__main__":
    run_example()

      
