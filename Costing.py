import os
import traceback
import xlwt
import datetime
import boto3
from azure.common.credentials import ServicePrincipalCredentials
from azure.mgmt.resource import ResourceManagementClient
from azure.mgmt.billing import BillingManagementClient
from azure.mgmt.consumption import ConsumptionManagementClient
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
email_from = 'IEG@duffandphelps.com'
email_to = ['IEG@duffandphelps.com']
boto3.setup_default_session(profile_name=AWS_PROFILE,region_name='us-east-1')
sts = boto3.client('sts')

wb = xlwt.Workbook(encoding="utf-8")
sheet1=wb.add_sheet("Costing")
vms = []

##sheet3=wb.add_sheet("sheet3")
##sheet3.write(0,0,"VmName")
##sheet3.write(0,1,"Subscription")
##sheet3.write(0,2,"Resource Group")
##sheet3.write(0,3,"Location")
##sheet3.write(0,4,"Provisioning State")
##sheet3.write(0,5,"OS Name")
##sheet3.write(0,6,"OS Version")
##sheet3.write(0,7,"Private IP")
##sheet3.write(0,8,"Public IP")
##sheet3.write(0,9,"VM Size")
##sheet3.write(0,10,"Disk Name")
##sheet3.write(0,11,"Disk Size (GB)")
##sheet3.write(0,12,"Admin Username")
##sheet3.write(0,13,"ApplicationName")
##sheet3.write(0,14,"BusinessOwner")
##sheet3.write(0,15,"Environment")
##sheet3.write(0,16,"ServiceLine")
##sheet3.write(0,17,"TechnologyOwner")

sheet1.write(0,0,"Subscription Name")
sheet1.write(0,1,"Subscription ID")
sheet1.write(0,2,"Billing Period")
sheet1.write(0,3,"Total Cost")
sheet1.write(0,4,"Compute")
sheet1.write(0,5,"Storage")
sheet1.write(0,6,"KeyVault")
sheet1.write(0,7,"Network")
sheet1.write(0,8,"SQL")
sheet1.write(0,9,"WEB")
sheet1.write(0,10,"ContainerRegistry")
sheet1.write(0,11,"ContainerInstance")
sheet1.write(0,12,"Cache")
sheet1.write(0,13,"Search")
sheet1.write(0,14,"Eventhub")
sheet1.write(0,15,"DocumentDB")
sheet1.write(0,16,"Logic")
sheet1.write(0,17,"AnalysisServices")
sheet1.write(0,18,"Servicebus")
sheet1.write(0,19,"ActiveDirectory")
sheet1.write(0,20,"DataFactory")
sheet1.write(0,21,"Insights")
sheet1.write(0,22,"DBforPostgreSQL")
sheet1.write(0,23,"ClassicCompute")
sheet1.write(0,24,"SignalRService")
sheet1.write(0,25,"PowerBIDedicated")
sheet1.write(0,26,"Automation")
sheet1.write(0,27,"Operational Insight")
sheet1.write(0,28,"RecoveryService")
sheet1.write(0,29,"Unassigned")

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
    body_text =["Please find the attached list of Costing in Subscriptions   \r\r\n Subscriptions "]
    credentials = get_credentials()
    resource_client = ResourceManagementClient(credentials, subscription_id)
    compute_client = ComputeManagementClient(credentials, subscription_id)
    network_client = NetworkManagementClient(credentials, subscription_id)
    resource_client = ResourceManagementClient(credentials, subscription_id)  
    con_client = ConsumptionManagementClient(credentials, subscription_id, base_url=None)
    billing_client = BillingManagementClient(credentials, subscription_id, base_url=None)
    k=0
    m=0
    subscriptionClient = SubscriptionClient(credentials)
    for subscription in subscriptionClient.subscriptions.list():
        cost =0
        compute =0
        storage =0
        keyvault =0
        network =0
        sql =0
        contR =0
        contI =0
        web =0
        event =0
        unass =0
        doc =0
        logic=0
        search=0
        servbus=0
        analy=0
        cache=0
        AD=0
        DF=0
        insight=0
        PSQL=0
        Ccomp=0
        SigR=0
        power=0
        aut =0
        rec =0
        oper=0
        sub_id = subscription.subscription_id
        con_client1 = ConsumptionManagementClient(credentials, sub_id, base_url=None)
        for vm in con_client1.usage_details.list():
            cost = cost + vm.pretax_cost
            if vm.consumed_service == 'Microsoft.Compute' : compute = compute + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.Storage' : storage = storage + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.KeyVault' : keyvault = keyvault + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.Network' : network = network + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.Sql' : sql = sql + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.Web' : web = web + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.ContainerRegistry' : contR = contR + vm.pretax_cost
            elif vm.consumed_service == 'microsoft.eventhub' : event = event + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.ContainerInstance' : contI = contI + vm.pretax_cost
            elif vm.consumed_service == 'Unassigned' : unass= unass + vm.pretax_cost
            elif vm.consumed_service == 'microsoft.documentdb' : doc = doc + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.Logic' : logic = logic + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.Search' : search = search + vm.pretax_cost
            elif vm.consumed_service == 'microsoft.servicebus' : servbus = servbus + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.AnalysisServices' : analy = analy + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.Cache' : cache = cache + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.AzureActiveDirectory' : AD = AD + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.DataFactory' :  DF = DF + vm.pretax_cost
            elif vm.consumed_service == 'microsoft.insights' : insight = insight + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.DBforPostgreSQL' : PSQL = PSQL + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.ClassicCompute' : Ccomp = Ccomp + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.SignalRService' : SigR = SigR + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.PowerBIDedicated' : power = power + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.Automation' : aut = aut + vm.pretax_cost
            elif vm.consumed_service == 'Microsoft.RecoveryServices' : rec = rec + vm.pretax_cost
            elif vm.consumed_service == 'microsoft.operationalinsights' : oper = oper + vm.pretax_cost
            else: pass
##            sheet3.write(k+1,0,vm.name)
##            sheet3.write(k+1,1,vm.subscription_name)
##            #sheet3.write(k+1,2,group.name)
##            #sheet3.write(k+1,3,vm.location)
##            sheet3.write(k+1,4,vm.usage_start.date().strftime("%m/%d/%Y"))
##            sheet3.write(k+1,5,vm.usage_end.date().strftime("%m/%d/%Y"))
##            sheet3.write(k+1,6,vm.instance_name)
##            sheet3.write(k+1,7,vm.instance_id)
##            sheet3.write(k+1,8,vm.instance_location)
##            sheet3.write(k+1,9,vm.usage_quantity)
##            sheet3.write(k+1,10,vm.currency)
##            sheet3.write(k+1,11,vm.pretax_cost)
##            sheet3.write(k+1,12,vm.account_name)
##            sheet3.write(k+1,13,vm.department_name)
##            sheet3.write(k+1,14,vm.product)
##            sheet3.write(k+1,15,vm.consumed_service)
##            sheet3.write(k+1,16,vm.cost_center)
##            k= k+1
        #print(subscription.display_name , cost)
        body_text.append("Total Cost of Subscription "+ subscription.display_name +": $"+str(float("{:.2f}".format(cost))))
        sheet1.write(m+1,0,subscription.display_name)
        sheet1.write(m+1,1,subscription.subscription_id)
        sheet1.write(m+1,2,today_date.strftime("%B_%Y"))
        sheet1.write(m+1,3,cost)
        sheet1.write(m+1,4,compute)
        sheet1.write(m+1,5,storage)
        sheet1.write(m+1,6,keyvault)
        sheet1.write(m+1,7,network)
        sheet1.write(m+1,8,sql)
        sheet1.write(m+1,9,web)
        sheet1.write(m+1,10,contR)
        sheet1.write(m+1,11,contI)
        sheet1.write(m+1,12,cache)
        sheet1.write(m+1,13,search)
        sheet1.write(m+1,14,event)
        sheet1.write(m+1,15,doc)
        sheet1.write(m+1,16,logic)
        sheet1.write(m+1,17,analy)
        sheet1.write(m+1,18,servbus)
        sheet1.write(m+1,19,AD)
        sheet1.write(m+1,20,DF)
        sheet1.write(m+1,21,insight)
        sheet1.write(m+1,22,PSQL)
        sheet1.write(m+1,23,Ccomp)
        sheet1.write(m+1,24,SigR)
        sheet1.write(m+1,25,power)
        sheet1.write(m+1,26,aut)
        sheet1.write(m+1,28,rec)
        sheet1.write(m+1,27,oper)
        sheet1.write(m+1,29,unass)      
        m=m+1
    #for subscription in subscriptionClient.subscriptions.list():
    day = today_date.strftime("%B_%Y")
    file_name = 'Azure_Costing_'+day+'.xls'
    wb.save(file_name)
    body_text1= "\r\r\n ".join(body_text)
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
        msg['Subject'] = "Azure Costing per Subscription List"
        msg['From'] = email_from
        msg['To'] = ', '.join(email_to)
        body = MIMEText(body_text1.encode(CHARSET), 'plain', CHARSET)
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

