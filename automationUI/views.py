# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from django.shortcuts import render
from django.shortcuts import redirect
import os,ctypes,time,datetime
import xlwt,xlrd,subprocess
from subprocess import Popen
from django.conf import settings
from django.http import HttpResponse
from django.http import request
from django.http import HttpRequest
from django.contrib import messages
import win32com,win32com.client
from xlutils.copy import copy
from win32com.client import Dispatch
from timeit import default_timer as timer
from django.contrib.staticfiles.templatetags import staticfiles
from django.template.response import TemplateResponse
from django.views.static import serve
import json
from .models import *
from django.contrib.auth.models import UserManager
from .tasks import *
from django.views.decorators.cache import cache_control
import logging
logging.basicConfig(filename='automationui.log',level=logging.DEBUG,format='%(asctime)s:%(funcName)s:%(message)s')

import pdb # pdb.set_trace()

# Create your views here.
global now, totaltimeexpected
WorkBookName = os.path.join(settings.STATIC_ROOT, 'data/RMData.xls')
RMWorkBookName = os.path.join(settings.STATIC_ROOT, 'data/RMData.xls')
ProceedTosubmitWorkBook = os.path.join(settings.STATIC_ROOT, 'data/ProceedToSubmit.xls')
asyncupdateworkbook = os.path.join(settings.STATIC_ROOT, 'data/asyncupdateworkbook.xls')
#approvaltype script list text file
ScriptListFile = "C:\\Users\\mnairchand\\Project\\vatephase2\\src\\static\\data\\ScriptListFile.txt"

AutomationPortal_QCWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AutomationPortal_Phase2_ADODB.vbs')
EmailTriggerScriptWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AP_AutoEmailTrigger.vbs')
RMenvironment = ""
Email = ""
#The main function which has got the entire workflow related to handling request
#the below cache_control does not allow the back arrow action to trigger. User will be thrown warning message
#@cache_control(no_cache=True,must_revalidate=True,no_store=True)
def index(request):
    global RMSanity, clientipaddress, datetimestamp, t0, t1, currenttime, clientipaddressTostop
    global RMProductTypeContext,ExecuteCounter,approvaltype
    ExecuteCounter = 0
######################################Execute Section Start#######################################################
    if (request.method =='POST') and (request.POST.get('Execute')=="Execute"):
        #################################obtaining client ip address############################################
        for i in request.META:
            if i == "REMOTE_ADDR":
                clientipaddress = request.META[i]
                break
        i = None
        #################################obtaining datetimestamp############################################
        timetostr = str(datetime.datetime.now())
        timewithnomillisec = timetostr.split(".")
        datetimestamp = timewithnomillisec[0].replace(":", "")
        ###################################storing the data obtained#########################################
        RMSanity = request.POST.get('RMSanity')
        Email = request.POST.get('email')
        approvaltype = request.POST.get('approvaltype')

        if RMSanity == "RMSanity":
            RMenvironment = request.POST.get('RMenvironment')
            RMcontentdict = readATE_ProceedToSubmitFlag()
            proceedtosubmitflag = RMcontentdict['proceedtosubmitflag']
            ExecuteCounter = int(RMcontentdict['ExecuteCounter'])
            initialrequesttime = RMcontentdict['RequestTime']
            totalwaittimeexpected = 2 * 60 * 60  # hard coded for timebeing
            t1val = time.time()
            t0 = time.time()  # to get the time in which the request submitted for else condition purpose

            # this condition is added to check if the second request from the user has come after a span of total wait time
            #########################################################################################
            if initialrequesttime!= "":
                tnow = t1val - float(initialrequesttime)
                if tnow > totalwaittimeexpected:
                    writeATE_ProceedToSubmitFlag(True, "")
                    RMcontentdict = readATE_ProceedToSubmitFlag()
                    proceedtosubmitflag = RMcontentdict['proceedtosubmitflag']
            #########################################################################################
            # this condition is added to write the requested data , invoke the vbscript to execute the requested script
            # trigger email on completion and render the page based on the info available
            if proceedtosubmitflag == u'True':
                ##########################################################################################
                DeleteTableRecords("automationUI_ate_rmdata")
                writeATE_RMData(request)  # ----------------write to newly created workbook----------------#
                pdb.set_trace()
                ExecuteCounter = ExecuteCounter + 1
                writeATE_ProceedToSubmitFlag(False, t0, clientrequestip=clientipaddress, RMenvironment=RMenvironment,Email=Email, ExecuteCounter=ExecuteCounter)
                ##########################################################################################
                #uncomment these lines during execution
                vbscriptcall.delay()  # task
                #emailclock.delay()  # task
                #checkandrestart_erroredoutscript.delay()  # task
                QCExecutionStatusUpdate.delay()  # task
                #########################################################################################
                pdb.set_trace()
                context = contextRMdataupdateAndemailtrigger()  # in this case email trigger scenario will not occur
                messages.info(request,"Execution is triggered and watchlist recipient will receive email on completion.")
                return render(request, 'automationUI/indexafterupdate.html', context)
                #########################################################################################
            else:
                # this condition is executed if the user submits the execute request before the wait time
                # it also triggers the task which will track on how long the user would need to wait
                #########################################################################################
                messages.info(request,"As machines are busy in executing request from IP: " + clientipaddress + " , please submit after sometime")
                # messages.info(request, "You would be directed to ATE page once the machines are available for taking up your request")
                exlcontentdict = readATE_ProceedToSubmitFlag()
                proceedtosubmitflag = exlcontentdict['proceedtosubmitflag']
                logging.debug('proceedtosubmitflagvalue is {0}'.format(proceedtosubmitflag))
                if proceedtosubmitflag == False:
                    waittimetask.delay(initialrequesttime)  # task
                else:
                    return render(request, 'automationUI/index.html')

                return render(request, 'automationUI/sorry.html')
                #########################################################################################
######################################Execute Section End#########################################################
######################################Refresh Section Start#######################################################
    if (request.method =='POST') and (request.POST.get('Refresh') == "Click Me For Test Status Update"):
        #context = contextRMdataupdateAndemailtrigger()
        pdb.set_trace()
        context = contextRMdataupdateAndemailtrigger()
        proceedtosubmitflag = context['proceedtosubmitflag']
        if not proceedtosubmitflag == True:
            messages.info(request,"Thank you for checking the status.")
            messages.info(request,"The specified machines are actively working to complete your script. Check back after few minutes again.")
            return render(request, 'automationUI/indexafterupdate.html', context)
        else:
            messages.info(request, "No further status check possible, please go to Home page to proceed with new request")
            return render(request, 'automationUI/indexforstop.html', context)
######################################Refresh Section End#########################################################
######################################Stop Section Start##########################################################
    if (request.method =='POST') and (request.POST.get('Stop') == "Stop"):
        context = contextRMdataupdateAndemailtrigger()
        contentval = readATE_ProceedToSubmitFlag()
        proceedtosubmitflag = contentval['proceedtosubmitflag']
        for val in contentval:
            if val=="clientrequestip":
                clientipaddress = contentval[val]

        logging.debug('clientipaddress is {0}'.format(clientipaddress))

        for i in request.META:
            if i == "REMOTE_ADDR":
                clientipaddressTostop = request.META[i]
                break
        i=None
        logging.debug('clientipaddressTostop is {0}'.format(clientipaddressTostop))

        if clientipaddress == clientipaddressTostop:
            if not proceedtosubmitflag == True:
                for i in context:
                    if context[i] == "Running":
                        context[i] = "Stopped"

                messages.info(request,"System would take approximately 10 minutes to stop the execution")
                messages.info(request,"Email would be triggered with the executed script status.")
                messages.info(request,"Please click on Home tab to submit next request if any")
                #killProcess()
                killProcess.delay()  #task
                #writeproceedtosubmitflag(True,"")
                return render(request, 'automationUI/indexforstop.html', context)
            else:
                messages.info(request, "As execution is already completed,Stop action will not be performed.")
                return render(request, 'automationUI/indexafterupdate.html', context)
######################################Stop Section End############################################################
######################################Send Report Section Start###################################################
    if (request.method == 'POST') and (request.POST.get('SendReport') == "Send Report"):
        iscompletedflag = False
        #########################################################################################
        # concatenating the details that has to be displayed on post
        contexttopost = readATE_ProceedToSubmitFlag()
        contextRMData = readRMDataStatusAndScriptName()
        context = {}
        for i in [contexttopost, contextRMData]:
            context.update(i)
            ########################################################################################
        for item in contextRMData:
            print (item)
            print (contextRMData[item])
            if "RMStatus" in item:
                if contextRMData[item].upper() == "FAILED" or contextRMData[item].upper() == "PASSED":
                    iscompletedflag = True
                    break
                else:
                    iscompletedflag = False

        exlcontentdict = readATE_ProceedToSubmitFlag()
        proceedtosubmitflag = exlcontentdict['proceedtosubmitflag']
        requesttime = exlcontentdict['requesttime']

        if iscompletedflag == True and proceedtosubmitflag==False and requesttime!="":
            EmailTrigger()  # trigger auto email
            messages.info(request, "Thank You for submitting the request. You would be receiving Email shortly")
            return render(request, 'automationUI/indexafterupdate.html', context)

        elif proceedtosubmitflag==True:
            messages.info(request, "Email is already sent for this execution to the WatchList recipient.")
            return render(request, 'automationUI/indexafterupdate.html', context)

        else:
            messages.error(request,"Email would be triggered only if a script gets Passed or Failed")
            messages.error(request,"Please retry if any or all of the script status shown as Passed or Failed")
            return render(request, 'automationUI/indexafterupdate.html', context)
######################################Send Report Section End#####################################################
    else:
        #connecting to QC to retrieve the script details
        ApvlTypeDict_ScriptNames = AprvlTypeRetrieveScriptNameAndWritejsFile()
        return render(request, 'automationUI/index.html',{'ApvlTypeDict_ScriptNames':ApvlTypeDict_ScriptNames})
###################################################################################################################
#The purpose of this function is used to read all the scripts under the Quote Approval Types
#And it also stores these values in a dictionary which is returned as part of function call
def AprvlTypeRetrieveScriptNameAndWritejsFile():
    gbQCTestPlanFolderName = "[QualityCenter]Subject\Automation\ECO\QuotesApprovalTypes"
    gbTestType = "QUICKTEST_TEST"
    Dict_ScriptNames = AP_GetAllScriptNameFromALMTestPlan_AndStoreInDictionary(gbQCTestPlanFolderName, gbTestType)
    jsfileformat = open(ScriptListFile, 'w')
    jsfileformat.write("Scriptlist={")
    for i in Dict_ScriptNames:
        print (Dict_ScriptNames[i])
        jsfileformat.write("\"" + i + "\"" + ":" + "\"" + Dict_ScriptNames[i] + "\",")
    jsfileformat.write("};")
    jsfileformat.close()
    return Dict_ScriptNames
##################################################################################################################
#The purpose of this function is used to update the first record related to Proceed to submit flag
def writeATE_ProceedToSubmitFlag(flag,starttime,clientrequestip="",RMenvironment="",Email="", ExecuteCounter=0):
    proceedtosubmitflag = flag
    RequestTime = starttime
    ClientRequestIP = clientrequestip
    RMenvironment = RMenvironment
    Email = Email
    ExecuteCounter = ExecuteCounter
    #update the specific record
    writeProceedFlag = ATE_ProceedToSubmitFlag.objects.get(rowid=1)
    writeProceedFlag.proceedtosubmitflag=proceedtosubmitflag
    writeProceedFlag.RequestTime = RequestTime
    writeProceedFlag.ClientRequestIP = ClientRequestIP
    writeProceedFlag.RMenvironment = RMenvironment
    writeProceedFlag.Email = Email
    writeProceedFlag.ExecuteCounter = ExecuteCounter
    writeProceedFlag.save()
##################################################################################################################
#The purpose of this function is used to read the contents related to proceed to submit flag
def readATE_ProceedToSubmitFlag():
    global proceedtosubmitflag,RequestTime,ClientRequestIP,RMenvironment,Email,ExecuteCounter
    ATE_ProceedToSubmitFlag_Records = ATE_ProceedToSubmitFlag.objects.all()
    for record in ATE_ProceedToSubmitFlag_Records:
        proceedtosubmitflag = record.proceedtosubmitflag
        RequestTime = record.RequestTime
        ClientRequestIP = record.ClientRequestIP
        RMenvironment = record.RMenvironment
        Email = record.Email
        ExecuteCounter = record.ExecuteCounter

    contextFlag = {'proceedtosubmitflag':proceedtosubmitflag,'RequestTime':RequestTime,'ClientRequestIP':ClientRequestIP,
                   'RMenvironment':RMenvironment,'Email':Email,'ExecuteCounter':ExecuteCounter}
    return contextFlag
##################################################################################################################
#The purpose of this function is used to connect QC
def ConnectToQC():
    import pythoncom
    pythoncom.CoInitialize()
    global gbQCConnection
    qcServer = "http://sv2wnecoqc01:8080/qcbin"
    qcUser = "automationqateam"
    qcPassword = "Welcome2"
    qcDomain = "DEFAULT"
    qcProject = "Auto_Test_Engine"
    gbQCConnection = win32com.client.Dispatch("TDApiOle80.TDConnection.1")
    gbQCConnection.InitConnectionEx(qcServer)
    gbQCConnection.Login(qcUser, qcPassword)
    gbQCConnection.Connect(qcDomain, qcProject)
    # pdb.set_trace()
    if gbQCConnection.Connected == True:
        print ("QCConnectionPass")
        return "QCConnectionPass"
    else:
        print ("QCConnectionFail")
        return "QCConnectionFail"
##################################################################################################################
#The purpose of this function is used to read all the script names from QC and store it in Dictionary
def AP_GetAllScriptNameFromALMTestPlan_AndStoreInDictionary(gbQCTestPlanFolderName, gbTestType):
    global QCTreeManager, TestNode, TestFact, TestsList, tests, Dict, Dict_ScriptNames, k, arrScriptName
    # Conenct To QC
    sQCConnection = ConnectToQC()
    QCTreeManager = gbQCConnection.TreeManager
    TestNode = QCTreeManager.NodeByPath(gbQCTestPlanFolderName)
    TestFact = TestNode.TestFactory
    TestsList = TestFact.NewList("")
    Dict_ScriptNames = {}
    k = 0
    for tests in TestsList:
        if tests.Field("TS_TYPE") == gbTestType and tests.Field("TS_Status") == "Ready":
            ScriptName = 'ScriptName_' + str(k)
            Dict_ScriptNames[ScriptName] = tests.Name
            k = k + 1
    return Dict_ScriptNames
##################################################################################################################
#The purpose of this function is used to kill UFT and clear cache in the server machine. It is triggered via batch file
def KillUFTandClearCacheOnServer():
   Popen("KillUFTandClearCacheOnServer.bat", cwd=r"C:\\Users\\automationqateam\\Project\\venv2\\src")
   #stdout, stderr = p.communicate()
   return True
##################################################################################################################
#The purpose of this function is used to write Quote Approvals Data in to the Data Table
def writeATE_QuoteApprovals(request,aprvltypecounter):
    for rec in range(1, int(aprvltypecounter) + 1):
        TestPlanFolderPath = "[QualityCenter]Subject\Automation\ECO\QuotesApprovalTypes"
        ApvlTypescript = request.POST.get('ApvlTypescript' + str(rec))
        TestSetFolderPath = "Root\ECO Automation\Product-Execution\AutomationPortal"
        ApvlTypemachine = request.POST.get('ApvlTypemachine' + str(rec))
        ApvlTypeEnvironment = request.POST.get('ApvlTypeEnvironment' + str(rec))
        Email = request.POST.get('email')
        ApvlTypeRecord = ATE_QuoteApprovals(TestPlanFolderPath=TestPlanFolderPath, TestScriptNameToExecute=ApvlTypescript,
                                      TestSetFolderPath=TestSetFolderPath, SanityName=approvaltype,
                                      RemoteMachineName=ApvlTypemachine, Environment=ApvlTypeEnvironment,
                                      Email=Email)
        ApvlTypeRecord.save()  # table is inserted with new record
##################################################################################################################
#The purpose of this function is used to write RM Data in to Data Table
def writeATE_RMData(request):
    TestPlanFolderPathColo = "[QualityCenter]Subject\Automation\ECO\Products\AutomationPortal\COLO"
    TestPlanFolderPathNetwork = "[QualityCenter]Subject\Automation\ECO\Products\AutomationPortal\Interconnection"
    TestPlanFolderPathService = "[QualityCenter]Subject\Automation\ECO\Products\AutomationPortal\Services"
    TestPlanFolderPathOther = "[QualityCenter]Subject\Automation\ECO\Products\AutomationPortal\Others"
    TestSetFolderPath = "Root\ECO Automation\Product-Execution\AutomationPortal"
    for rm in range(1,5):
        RowID = rm
        RMenvironment = request.POST.get('RMenvironment')
        RMscriptname = request.POST.get('RMscriptname' + str(rm))
        RMmachine = request.POST.get('RMmachine' + str(rm))
        RMProductType = request.POST.get('RMProductType' + str(rm))
        Email = request.POST.get('email')

        if RMProductType == u"Colo":
            TestPlanFolderPath = TestPlanFolderPathColo
        elif RMProductType == u"Network":
            TestPlanFolderPath = TestPlanFolderPathNetwork
        elif RMProductType == u"Service":
            TestPlanFolderPath = TestPlanFolderPathService
        else:
            TestPlanFolderPath = TestPlanFolderPathOther

        RMRecord = ATE_RMData(RowID = RowID,TestPlanFolderPath=TestPlanFolderPath,TestScriptNameToExecute=RMscriptname,
                              TestSetFolderPath = TestSetFolderPath,SanityName = RMSanity,RemoteMachineName =RMmachine,
                              Environment = RMenvironment,Email=Email)
        RMRecord.save()
##################################################################################################################
#The purpose of this function is sed to read RM Data script and status from the data table and store it in a dictionary
def readRMDataStatusAndScriptName():
    RMDictRecordFullset = {}
    readRMDataStatus_Records = ATE_RMData.objects.all()
    counter = 1
    for record in readRMDataStatus_Records:
        RMScriptName = record.TestScriptNameToExecute
        RMStatus = record.RunningStatus
        RMDictRecord = {'RMScriptName'+str(counter):RMScriptName,'RMStatus'+str(counter):RMStatus}
        RMDictRecordFullset.update(RMDictRecord)
        counter = counter + 1
    return RMDictRecordFullset
##################################################################################################################
##################################################################################################################
#The purpose of this function is sed to read RM Data script and status from the data table and store it in a dictionary
def readRMDataStatus():
    RMDictRecordFullset = {}
    readRMDataStatus_Records = ATE_RMData.objects.all()
    counter = 1
    for record in readRMDataStatus_Records:
        RMStatus = record.RunningStatus
        RMDictRecord = {'RMStatus'+str(counter):RMStatus}
        RMDictRecordFullset.update(RMDictRecord)
        counter = counter + 1
    return RMDictRecordFullset
##################################################################################################################
#The purpose of this function is used to trigger email. This function in turn calls a task
#The purpose of the task is to kill UFT, clear cache on server and Email trigger.
#The same task with function is called even when the user wants to send email intermittently  or end of execution
def EmailTrigger():
    KillUFTandClearCacheOnServerAndEmailTriggerForIntermittent.delay() #task
##################################################################################################################
#The purpose of this function is to read the updated RM data and concatenate it with the client requested details
#It also triggers email if all the row in the RM Data is either PASSED or FAILED
def contextRMdataupdateAndemailtrigger():
    iscompletedflag = False
    contexttopost = readATE_ProceedToSubmitFlag()
    contextRMData = readRMDataStatus()
    context = {}
    for i in [contexttopost, contextRMData]:
        context.update(i)
        ########################################################################################
    for item in contextRMData:
        if (contextRMData[item].upper() == "PASSED" or contextRMData[item].upper() == "FAILED"):
            iscompletedflag = True
            break
        else:
            iscompletedflag = False

    if iscompletedflag == True:
        EmailTrigger()  # trigger auto email

    return context
##################################################################################################################
#The purpose of this function is used to delete all records from the table
def DeleteTableRecords(TableName):
    objDBConnection = win32com.client.Dispatch("ADODB.Connection")
    objDBConnection.Open(
        "Provider=PostgreSQL OLE DB Provider;Data Source=localhost;location=ate;User ID=postgres;password=postgres")
    objDBRecordset = objDBConnection.Execute("Delete FROM " + "\"" + TableName + "\"")

##################################################################################################################