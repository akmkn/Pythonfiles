from __future__ import absolute_import, unicode_literals
from celery import Celery
from celery import current_app
from celery import Task
from celery.task import task
#from celery.decorators import task
import time,xlrd,xlwt,subprocess,os
from xlutils.copy import copy
from subprocess import Popen
from django.conf import settings
from django.shortcuts import HttpResponse
import win32com,win32com.client
from win32com.client import Dispatch
from .views import *
from .models import *
import os

#asyncupdateworkbook = 'C:/Users/automationqateam/Project/venv2/static_cdn/data/asyncupdateworkbook.xls'
#MachineNameToKill = 'C:/Users/automationqateam/Project/venv2/src/KillTaskOnSpecificMachine_MachineName.txt'

asyncupdateworkbook = 'C:/Users/mnairchand/Project/venv2/static_cdn/data/asyncupdateworkbook.xls'
MachineNameToKill = 'C:/Users/mnairchand/Project/venv2/src/KillTaskOnSpecificMachine_MachineName.txt'

ProceedTosubmitWorkBook = os.path.join(settings.STATIC_ROOT, 'data/ProceedToSubmit.xls')
WorkBookName = os.path.join(settings.STATIC_ROOT, 'data/RMData.xls')
#AutomationPortal_QCWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AutomationPortal_Phase1_ADODB.vbs')
AutomationPortal_QCWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AutomationPortal_Phase2_ADODB.vbs')
AutomationPortal_UpdateQCExecutionStatusInExcel = os.path.join(settings.STATIC_ROOT, 'vbs/AutomationPortal_Phase2_UpdateQCExecutionStatusInDB.vbs')
EmailTriggerScriptWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AP_AutoEmailTrigger.vbs')
#task created to trigger the vbscript machine
@task()
def vbscriptcall():
    #wkbook = xlrd.open_workbook(WorkBookName,encoding_override="cp1252")
    #sheetname = wkbook.sheet_by_index(0)
    #rowcount = int(sheetname.nrows)
    #rowcount = ATE_RMData.objects.count() #total number of records in database
    #for i in range(rowcount):
    subprocess.Popen(["wscript.exe",AutomationPortal_QCWorkflow],stdout=subprocess.PIPE)
    #         if i==1:
    #             time.sleep(120)
    #         else:
    #             time.sleep(60)

    #os.system("wscript.exe C:\\Users\\automationqateam\\Project\\venv2\\static_cdn\\vbs\\AP_ClientProcessedRecordsHistory.vbs")
    os.system("wscript.exe C:\\Users\\mnairchand\\Project\\venv2\\static_cdn\\vbs\\AP_ClientProcessedRecordsHistory.vbs")
    return True

# latest change to handle UFT error - start - 20 july
# UFT Error Handling function to re-start the script on errored out remote machine
# This function task will continously check (after every 5 mins if proceed to submit flag is flase) in RM data file,
# if UFT is not able to run the script then that machine will be refreshed and script will be re-started
@task()
def checkandrestart_erroredoutscript():
    dictcontent = readATE_ProceedToSubmitFlag()
    proceedtosubmitflag = dictcontent['proceedtosubmitflag']
    while proceedtosubmitflag == False:
        waitfortriggercheck = 5 * 60  # hard coded for timebeing 5 min
        # totaltimeexpected = 30  # hard coded for timebeing
        curnt = 0.0
        tt0 = time.time()
        while curnt < waitfortriggercheck:
            tt1 = time.time()
            curnt = tt1 - tt0
        # asyncupdatestatus()
        Restart_ErroredOut_Script_OnAssignedMachine()
        dictcontent = readATE_ProceedToSubmitFlag()
        proceedtosubmitflag = dictcontent['proceedtosubmitflag']
    return True

# write to file
def Restart_ErroredOut_Script_OnAssignedMachine():
    triggerflag = False
    machinename = ""
    readRMDataStatus_Records = ATE_RMData.objects.all()
    for record in readRMDataStatus_Records:
        RMTrigger = record.Trigger
        if RMTrigger.upper() == "ERROR":
            RMMachineName = record.RemoteMachineName
            triggerflag = True
            break
    if  triggerflag == True:
        fileformachine = open(MachineNameToKill, 'w')
        fileformachine.write(machinename)
        fileformachine.close()
        #os.system("C:\\Users\\automationqateam\\Project\\venv2\\src\\KillTaskOnSpecificMachine.bat")
        os.system("C:\\Users\\mnairchand\\Project\\venv2\\src\\KillTaskOnSpecificMachine.bat")
        subprocess.Popen(["wscript.exe",AutomationPortal_QCWorkflow],stdout=subprocess.PIPE)
        time.sleep(120)
    return True

# regualry check for status from QC after every 7 mins and update the QC status into RM Data file
# if any script is of No Run then make trigger as Error and re-start the script on specific machine.
#AutomationPortal_UpdateQCExecutionStatusInExcel.vbs
@task()
def QCExecutionStatusUpdate():
    dictcontent = readATE_ProceedToSubmitFlag()
    proceedtosubmitflag = dictcontent['proceedtosubmitflag']
    while proceedtosubmitflag == False:
        waitforqcstatuscheck = 5 * 60  # 5 minutes time interval
        checktimenow = 0.0
        chk0 = time.time()
        while checktimenow < waitforqcstatuscheck:
            chk1 = time.time()
            checktimenow = chk1 - chk0
        subprocess.Popen(["wscript.exe", AutomationPortal_UpdateQCExecutionStatusInExcel], stdout=subprocess.PIPE)
        dictcontent = readATE_ProceedToSubmitFlag()
        proceedtosubmitflag = dictcontent['proceedtosubmitflag']
    return True

# latest change to handle UFT error - end

#task created to trigger email
#This in turn invokes the function which would kill the UFT and Clear cache in the server machine
@task()
def emailclock():
    iscompletedflag = False
    allstoppedflag = False
    #while iscompletedflag == False and allstoppedflag==False:
    while iscompletedflag == False and allstoppedflag == False:
        time.sleep(300)
        flagdict = ATE_RMStatusCheck()
        iscompletedflag = flagdict['iscompletedflag']
        allstoppedflag = flagdict['allstoppedflag']

    if iscompletedflag == True and allstoppedflag==False :
        #os.system("C:\\Users\\automationqateam\\Project\\venv2\\src\\KillProcessByName.bat")
        os.system("C:\\Users\\mnairchand\\Project\\venv2\\src\\KillProcessByName.bat")
        clearRemoteTempCacheChrome()
        KillUFTandClearCacheOnServer()
        subprocess.Popen(["wscript.exe", EmailTriggerScriptWorkflow], stdout=subprocess.PIPE)
        time.sleep(200)
        writeATE_ProceedToSubmitFlag(True, "")
    return True

#A wait time task which is triggered only when the user clicks on the Execute button before the completion of initial request
#this function will run to the maximum of 2 hours before accepting the next request
@task()
def waittimetask(t0):
    totaltimeexpected = 2 * 60 * 60  # hard coded for timebeing
    #totaltimeexpected = 30  # hard coded for timebeing
    now = 0.0
    while now < totaltimeexpected:
        t1 = time.time()
        now = t1 - t0
    #asyncupdatestatus()
    writeATE_ProceedToSubmitFlag(True,"")
    return True

#The task will invoke two batch process. One will kill the UFT and chrome in all the virtual machines
#second one will clear the Chrome cache from all the Remote machines
@task()
def killProcess():
   #os.system("C:\\Users\\automationqateam\\Project\\venv2\\src\\KillProcessByName.bat")
   os.system("C:\\Users\\mnairchand\\Project\\venv2\\src\\KillProcessByName.bat")
   #Popen("KillProcessByName.bat", cwd=r"C:\\Users\\automationqateam\\Project\\venv2\\src")
   #stdout, stderr = p.communicate()
   clearRemoteTempCacheChrome()
   writeATE_ProceedToSubmitFlag(True, "")

   return True

#This task is to send intermittent email and before launching UFT, kill UFT if already opened on server.
@task()
def KillUFTandClearCacheOnServerAndEmailTriggerForIntermittent():
   #os.system("C:\\Users\\automationqateam\\Project\\venv2\\src\\KillUFTandClearCacheOnServer.bat")
   os.system("C:\\Users\\mnairchand\\Project\\venv2\\src\\KillUFTandClearCacheOnServer.bat")
   subprocess.Popen(["wscript.exe", EmailTriggerScriptWorkflow], stdout=subprocess.PIPE)
   return True

# This function is for final email sending
#function which would kill the UFT and Clear cache in the server machine
def KillUFTandClearCacheOnServer():
   #os.system("C:\\Users\\automationqateam\\Project\\venv2\\src\\KillUFTandClearCacheOnServer.bat")
   os.system("C:\\Users\\mnairchand\\Project\\venv2\\src\\KillUFTandClearCacheOnServer.bat")
   #Popen("KillUFTandClearCacheOnServer.bat", cwd=r"C:\\Users\\automationqateam\\Project\\venv2\\src")
   #stdout, stderr = p.communicate()
   return True

#The purpose of this function is used to retrieve the status value for RM Data
def ATE_RMStatusCheck():
    stopcounter = 0
    allstoppedflag = False
    iscompletedflag = False
    RMDictRecordFullset = {}
    readRMDataStatus_Records = ATE_RMData.objects.all()
    counter = 1
    for record in readRMDataStatus_Records:
        RMStatus = record.RunningStatus
        RMDictRecord = {'RMStatus'+str(counter):RMStatus}
        RMDictRecordFullset.update(RMDictRecord)
        counter = counter + 1

    for RM in RMDictRecordFullset:
        if (RMDictRecordFullset[RM].upper() != "PASSED" and RMDictRecordFullset[RM].upper() != "FAILED" and RMDictRecordFullset[RM].upper() != "STOPPED"):
            iscompletedflag = False
            break
        else:
            iscompletedflag = True

        if (RMDictRecordFullset[RM].upper() == "STOPPED"):
            stopcounter = stopcounter+1

    if stopcounter == len(RMDictRecordFullset):
        allstoppedflag = True
    else:
        allstoppedflag = False

    flagdict = {'allstoppedflag':allstoppedflag,'iscompletedflag':iscompletedflag}
    return flagdict

#function used to clear the Chrome cache from all the Remote machines
def clearRemoteTempCacheChrome():
   #os.system("C:\\Users\\automationqateam\\Project\\venv2\\src\\RunBatchToClearRemoteCache.bat")
   os.system("C:\\Users\\mnairchand\\Project\\venv2\\src\\RunBatchToClearRemoteCache.bat")
   #Popen("RunBatchToClearRemoteCache.bat", cwd=r"C:\\Users\\automationqateam\\Project\\venv2\\src")
   #stdout, stderr = p.communicate()
   return True

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

