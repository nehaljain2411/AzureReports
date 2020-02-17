import os
import traceback
import xlwt
import boto3
import datetime
from azure.common.credentials import ServicePrincipalCredentials
from azure.mgmt.resource import ResourceManagementClient
from azure.mgmt.network import NetworkManagementClient
from azure.mgmt.compute import ComputeManagementClient
from azure.mgmt.compute.models import DiskCreateOption
from msrestazure.azure_exceptions import CloudError
from azure.mgmt.resource import SubscriptionClient
from azure.mgmt.web import WebSiteManagementClient
from botocore.exceptions import ClientError
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
today_date=datetime.date.today()
AWS_PROFILE='myprofile1'
email_from = 'IEG@duffandphelps.com'
email_to = ['IEG@duffandphelps.com']
boto3.setup_default_session(profile_name=AWS_PROFILE,region_name='us-east-1')
sts = boto3.client('sts')
wb = xlwt.Workbook(encoding="utf-8")
sheet2=wb.add_sheet("Subscription")

svms = []

sheet2.write(0,0,"Subscription ID")
sheet2.write(0,1,"Subscription Name")
sheet2.write(0,2,"Authorization Source")

sheet3=wb.add_sheet("Resource Group")

vms = []

sheet3.write(0,0,"Resouce Group Name")
sheet3.write(0,1,"Subscription")
sheet3.write(0,2,"Location")
sheet3.write(0,3,"Managed By")
sheet3.write(0,4,"Provisioning State")
sheet3.write(0,5,"ApplicationName")
sheet3.write(0,6,"BusinessOwner")
sheet3.write(0,7,"Environment")
sheet3.write(0,8,"ServiceLine")
sheet3.write(0,9,"TechnologyOwner")

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
    web_client = WebSiteManagementClient(credentials, subscription_id)    
    
    try:
        k=0
        m=0
        subscriptionClient = SubscriptionClient(credentials)
        for subscription in subscriptionClient.subscriptions.list():
            sub_id = subscription.subscription_id
            svms.append(sub_id)
            sheet2.write(m+1,0,sub_id)
            sheet2.write(m+1,1,subscription.display_name)
            sheet2.write(m+1,2,subscription.authorization_source)
            m =m+1
        print("Total VMs Count : "+str(len(svms)))
        for subscription in subscriptionClient.subscriptions.list():
            sub_id = subscription.subscription_id
            resource_client1 = ResourceManagementClient(credentials, sub_id)
            web_client1 = WebSiteManagementClient(credentials, sub_id)
            groups = resource_client1.resource_groups.list()
            for vm in groups:
                vms.append(vm.name)
                sheet3.write(k+1,0,vm.name)
                sheet3.write(k+1,1,subscription.display_name)
                sheet3.write(k+1,2,vm.location)
                sheet3.write(k+1,3,vm.managed_by)
                sheet3.write(k+1,4,vm.properties.provisioning_state)
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
                        sheet3.write(k+1,5,appname)
                        sheet3.write(k+1,6,busname)
                        sheet3.write(k+1,7,envname)
                        sheet3.write(k+1,8,servname)
                        sheet3.write(k+1,9,techname)
                except: pass
                k =k+1
        print("Total VMs Count : "+str(len(vms)))
        day = today_date.strftime("%B_%Y")
        file_name = 'Azure_Subscription_ResourceGroup_'+day+'.xls'
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
        body_text = (
            "Attached herewith is the latest Azure Subscription and Resource Groups Inventory \r\r\n"
                      "Total Number of Subscriptions: " + str(len(svms)) + " \r\r\n"
                      "Total Number of Resource Groups: " + str(len(vms)) + " \r\r\n"
                     )
        #html = str(sys.argv[3])
        msg['Subject'] = "Azure Subscription and Resource Group Inventory List"
        msg['From'] = email_from
        msg['To'] = ', '.join(email_to)
        body = MIMEText(body_text.encode(CHARSET), 'plain', CHARSET)
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
