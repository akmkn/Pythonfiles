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

# Create your views here.

global WorkBookName,AutomationPortal_QCWorkflow,EmailTriggerScriptPath,RMWorkBookName,ProceedTosubmitWorkBook
global now, totaltimeexpected, RMenvironment
WorkBookName = os.path.join(settings.STATIC_ROOT, 'data/RMData.xls')
#WorkBookName = '\\sgw-filesvr3\GDC_Team\QA\ECO-Product-Docs\Automation\AutomationPortal\QCWorkFlow\RMData.xls'
RMWorkBookName = os.path.join(settings.STATIC_ROOT, 'data/RMData.xls')
ProceedTosubmitWorkBook = os.path.join(settings.STATIC_ROOT, 'data/ProceedToSubmit.xls')

AutomationPortal_QCWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AutomationPortal_Phase1_ADODB.vbs')
EmailTriggerScriptWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AP_AutoEmailTrigger.vbs')

#The main function which has got the entire workflow related to handling request
def index(request):
    global RMenvironment1, RMscriptname1, RMmachine1, RMexecute1
    global RMenvironment2, RMscriptname2, RMmachine2, RMexecute2
    global RMenvironment3, RMscriptname3, RMmachine3, RMexecute3
    global RMenvironment4, RMscriptname4, RMmachine4, RMexecute4,RMenvironment
    global RMSanity, Email, clientipaddress, datetimestamp, t0, t1, currenttime, clientipaddressTostop

    if (request.method =='POST') and (request.POST.get('Execute')=="Execute"):
        #################################obtaining client ip address############################################
        for i in request.META:
            if i == "REMOTE_ADDR":
                clientipaddress = request.META[i]
                break
        #################################obtaining datetimestamp############################################
        timetostr = str(datetime.datetime.now())
        timewithnomillisec = timetostr.split(".")
        datetimestamp = timewithnomillisec[0].replace(":", "")
        ###################################storing the data obtained#########################################
        RMSanity = request.POST.get('RMSanity')
        Email = request.POST.get('email')
       #ctypes.windll.user32.MessageBoxW(0, RMSanity, "Your Title", 1)
        if RMSanity == "RMSanity":
            RMenvironment = request.POST.get('RMenvironment')
            RMenvironment1 = RMenvironment
            RMscriptname1 = request.POST.get('RMscriptname1')
            RMmachine1 = request.POST.get('RMmachine1')
            RMexecute1 = request.POST.get('RMexecute1')
            if RMexecute1 == "on":
                RMexecute1 = "Yes"

            RMenvironment2 = RMenvironment
            RMscriptname2 = request.POST.get('RMscriptname2')
            RMmachine2 = request.POST.get('RMmachine2')
            RMexecute2 = request.POST.get('RMexecute2')
            if RMexecute2 == "on":
                RMexecute2 = "Yes"

            RMenvironment3 = RMenvironment
            RMscriptname3 = request.POST.get('RMscriptname3')
            RMmachine3 = request.POST.get('RMmachine3')
            RMexecute3 = request.POST.get('RMexecute3')
            if RMexecute3 == "on":
                RMexecute3 = "Yes"

            RMenvironment4 = RMenvironment
            RMscriptname4 = request.POST.get('RMscriptname4')
            RMmachine4 = request.POST.get('RMmachine4')
            RMexecute4 = request.POST.get('RMexecute4')
            if RMexecute4 == "on":
               RMexecute4 = "Yes"

            exlcontentdict = readproceedtosubmitflag()
            proceedtosubmitflag = exlcontentdict['proceedtosubmitflag']
            initialrequesttime = exlcontentdict['requesttime']
            totalwaittimeexpected = 2*60*60  # hard coded for timebeing
            t1val = time.time()
            t0 = time.time()  # to get the time in which the request submitted for else condition purpose
            print t0
            print t1val
            print initialrequesttime

            if initialrequesttime!="":
                tnow = t1val - initialrequesttime
                print tnow
                if tnow > totalwaittimeexpected:
                    writeproceedtosubmitflag(True,"")
                    exlcontentdict = readproceedtosubmitflag()
                    proceedtosubmitflag = exlcontentdict['proceedtosubmitflag']
               # pdb.set_trace()
            if proceedtosubmitflag == True:
                ##########################################################################################
                writeRMData(clientipaddress,datetimestamp)     #----------------write to newly created workbook----------------#
                writeproceedtosubmitflag(False,t0)
                #pdb.set_trace()
                ##########################################################################################
                vbscriptcall()    #----------------vbscript to connect QC and invoke script ---------------#
                #########################################################################################
                context = contextRMdataupdateAndemailtrigger()  #in this case email trigger scenario will not occur

                messages.error(request,"Due to limited machines, we could accept your next request after span of two hours")
                messages.error(request,"Sorry for the inconvenience caused")

                return render(request, 'automationUI/indexafterupdate.html', context)

            else:
                totaltimeexpected = 2*60*60  # hard coded for timebeing
                #totaltimeexpected = 10  # hard coded for timebeing
                #newt0 = t0
                now = 0.0

                while now < totaltimeexpected:
                    t1 = time.time()
                    now = t1 - t0
                    print (now)

                if now >= totaltimeexpected:
                    writeproceedtosubmitflag(True,"")

                return render(request, 'automationUI/sorry.html')

    if (request.method == 'POST') and (request.POST.get('Refresh') == "Click Me For Test Status Update"):
        context = contextRMdataupdateAndemailtrigger()
        messages.info(request,"Thank you for checking the status")
        messages.info(request, "The specified machines are actively working to complete your script. Check back after few mts again")
        return render(request, 'automationUI/indexafterupdate.html', context)

    if (request.method == 'POST') and (request.POST.get('Stop') == "Stop"):
        context = contextRMdataupdateAndemailtrigger()
        for i in request.META:
            if i == "REMOTE_ADDR":
                clientipaddressTostop = request.META[i]
                break
        if clientipaddress==clientipaddressTostop:
            print (clientipaddressTostop)
            print (clientipaddress)
            messages.info(request,"System is about to stop your execution and it would take minimum of 10 minutes to complete the Process")
            messages.info(request,"You would be directed to the main Page to submit the next request")
            killProcess()
            writeproceedtosubmitflag(True,"")
            return render(request, 'automationUI/index.html')
        else:
            print (clientipaddressTostop)
            print (clientipaddress)
            messages.info(request,"As you are not the intended user who initiated the Execute request, system cannot accept the Stop request")
            messages.info(request,"Sorry for the inconvenience caused")
            return render(request, 'automationUI/indexafterupdate.html', context)

    if (request.method == 'POST') and (request.POST.get('SendReport') == "Send Report"):
        iscompletedflag = False
        #########################################################################################
        # concatenating the details that has to be displayed on post
        contexttopost = {'RMenvironment': RMenvironment, 'Email': Email}
        contextRMData = readRMData()
        context = {}
        for i in [contexttopost, contextRMData]:
            context.update(i)
            ########################################################################################
        for item in contextRMData:
            if (contextRMData[item].upper() != "PASSED" or contextRMData[item].upper() != "FAILED"):
                iscompletedflag = False
                break
            else:
                iscompletedflag = True

        if iscompletedflag == True:
            EmailTrigger()  # trigger auto email
            messages.error(request, "Thank You for submitting the request. You would be receiving Email shortly")
            return render(request, 'automationUI/indexafterupdate.html', context)
        else:
            messages.error(request,"Execution is yet to complete to trigger the email")
            messages.error(request,"Please click on the Send Report button once all the script status shown as Passed or Failed")
            return render(request, 'automationUI/indexafterupdate.html', context)

    else:
        return render(request, 'automationUI/index.html')

def readRMData():
    wkbook = xlrd.open_workbook(WorkBookName,'r')
    sheetname = wkbook.sheet_by_index(0)
    RM1statusval = sheetname.cell_value(1, 8) #  Run Time Status value
    RM2statusval = sheetname.cell_value(2, 8)
    RM3statusval = sheetname.cell_value(3, 8)
    RM4statusval = sheetname.cell_value(4, 8)
    context = {'RM1statusval': RM1statusval,'RM2statusval':RM2statusval,'RM3statusval':RM3statusval,'RM4statusval':RM4statusval}
    return context

def vbscriptcall():
    wkbook = xlrd.open_workbook(WorkBookName)
    sheetname = wkbook.sheet_by_index(0)
    #colcount = sheetname.ncols
    rowcount = int(sheetname.nrows)
    print rowcount
    ctypes.windll.user32.MessageBoxW(0,rowcount,"Your Title",1)
    for i in range(rowcount):
        subprocess.Popen(["wscript.exe",AutomationPortal_QCWorkflow],stdout=subprocess.PIPE)
        if i==1:
           time.sleep(90)
        else:
           time.sleep(30)
             #ctypes.windll.user32.MessageBoxW(0,"Middle","Your Title",1)

#Purpose: To write RM Data into the excel sheet either collated one or individual or both
def writeRMData(clientipaddress,datetimestamp):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('RM')

    TestPlanFolderPath = "[QualityCenter]Subject\Automation\ECO\Products\AutomationPortal"
    TestSetFolderPath = "Root\ECO Automation\Product-Execution\AutomationPortal"
    SanityName = "RM"
    ExecutorId = "All"

    ws.write(0, 0, "Execute")
    ws.write(0, 1, "TestPlanFolderPath")
    ws.write(0, 2, "TestScriptNameToExecute")
    ws.write(0, 3, "TestSetFolderPath")
    ws.write(0, 4, "SanityName")
    ws.write(0, 5, "RemoteMachineName")
    ws.write(0, 6, "rt_IsScriptRunning")
    ws.write(0, 7, "rt_IsMachineAvailable")
    ws.write(0, 8, "RunningStatus")
    ws.write(0, 9, "Trigger")
    ws.write(0, 10, "Environment")
    ws.write(0, 11, "IsTestSetCreated")
    ws.write(0, 12, "ALM_NewTestSetName")
    ws.write(0, 13, "MailTo")
    ws.write(0, 14, "ExecutorId")

    if RMSanity == "RMSanity":
        ws.write(1, 0, RMexecute1)
        ws.write(2, 0, RMexecute2)
        ws.write(3, 0, RMexecute3)
        ws.write(4, 0, RMexecute4)

        ws.write(1, 1, TestPlanFolderPath)
        ws.write(2, 1, TestPlanFolderPath)
        ws.write(3, 1, TestPlanFolderPath)
        ws.write(4, 1, TestPlanFolderPath)

        ws.write(1, 2, RMscriptname1)
        ws.write(2, 2, RMscriptname2)
        ws.write(3, 2, RMscriptname3)
        ws.write(4, 2, RMscriptname4)

        ws.write(1, 3, TestSetFolderPath)
        ws.write(2, 3, TestSetFolderPath)
        ws.write(3, 3, TestSetFolderPath)
        ws.write(4, 3, TestSetFolderPath)

        ws.write(1, 4, SanityName)
        ws.write(2, 4, SanityName)
        ws.write(3, 4, SanityName)
        ws.write(4, 4, SanityName)

        ws.write(1, 5, RMmachine1)
        ws.write(2, 5, RMmachine2)
        ws.write(3, 5, RMmachine3)
        ws.write(4, 5, RMmachine4)

        #6 - 9 auto populated

        ws.write(1, 10, RMenvironment1)
        ws.write(2, 10, RMenvironment2)
        ws.write(3, 10, RMenvironment3)
        ws.write(4, 10, RMenvironment4)

        #11 - 12 auto populated
        ws.write(1, 13, Email)
        ws.write(1, 14, ExecutorId)

        #use the below three lines when it is necessary to created individual Data sheet for RM
        #and an consolidated RM Data sheet having all individual data sheet values
#       RMSinglewbname = 'data/RMData_'+clientipaddress+'_'+datetimestamp+'.xls'
#       wb.save(os.path.join(settings.STATIC_ROOT,RMSinglewbname))
#       appendingRMRequestData()  # to append all the Release management data

        #use the below line if its necessary to create just one RM Data sheet not being appended
        wb.save(os.path.join(settings.STATIC_ROOT, 'data/RMData.xls'))

#Purpose: To append all the request data received from the client to an excel named as RMData
def appendingRMRequestData():
    wkbook = xlrd.open_workbook(RMWorkBookName)
    sheetnameforrowcount = wkbook.sheet_by_index(0)
    rowcount = int(sheetnameforrowcount.nrows)  #get the rows count

    #the lines to write to excel sheet
    openbook = xlrd.open_workbook(RMWorkBookName)
    wb = copy(openbook) #xlutil is being used here
    sheetname = wb.get_sheet("data")

    TestPlanFolderPath = "[QualityCenter]Subject\Automation\ECO\Products\AutomationPortal"
    TestSetFolderPath = "Root\ECO Automation\Product-Execution\AutomationPortal"
    SanityName = "ReleaseManagement"
    rt_IsRunning = "No"
    rt_IsStop = "No"
    Email = "ashiskumar@ap.equinix.com"
    ExecutorId = "All"

    sheetname.write(rowcount, 0, RMexecute1)
    sheetname.write(rowcount+1, 0, RMexecute2)
    sheetname.write(rowcount+2, 0, RMexecute3)
    sheetname.write(rowcount+3, 0, RMexecute4)

    sheetname.write(rowcount, 1, TestPlanFolderPath)
    sheetname.write(rowcount + 1, 1, TestPlanFolderPath)
    sheetname.write(rowcount + 2, 1, TestPlanFolderPath)
    sheetname.write(rowcount + 3, 1, TestPlanFolderPath)

    sheetname.write(rowcount, 2, RMscriptname1)
    sheetname.write(rowcount+1, 2, RMscriptname2)
    sheetname.write(rowcount+2, 2, RMscriptname3)
    sheetname.write(rowcount+3, 2, RMscriptname4)

    sheetname.write(rowcount, 3, TestSetFolderPath)
    sheetname.write(rowcount+1, 3, TestSetFolderPath)
    sheetname.write(rowcount+2, 3, TestSetFolderPath)
    sheetname.write(rowcount+3, 3, TestSetFolderPath)

    sheetname.write(rowcount, 4, SanityName)
    sheetname.write(rowcount+1, 4, SanityName)
    sheetname.write(rowcount+2, 4, SanityName)
    sheetname.write(rowcount+3, 4, SanityName)

    sheetname.write(rowcount, 5, RMmachine1)
    sheetname.write(rowcount+1, 5, RMmachine2)
    sheetname.write(rowcount+2, 5, RMmachine3)
    sheetname.write(rowcount+3, 5, RMmachine4)

    sheetname.write(rowcount, 6, rt_IsRunning)
    sheetname.write(rowcount+1, 6, rt_IsRunning)
    sheetname.write(rowcount+2, 6, rt_IsRunning)
    sheetname.write(rowcount+3, 6, rt_IsRunning)

    sheetname.write(rowcount, 7, rt_IsStop)
    sheetname.write(rowcount+1, 7, rt_IsStop)
    sheetname.write(rowcount+2, 7, rt_IsStop)
    sheetname.write(rowcount+3, 7, rt_IsStop)

    sheetname.write(rowcount, 10, RMenvironment1)
    sheetname.write(rowcount+1, 10, RMenvironment2)
    sheetname.write(rowcount+2, 10, RMenvironment3)
    sheetname.write(rowcount+3, 10, RMenvironment4)
    wb.save(RMWorkBookName)

#Purpose: To read from the ProceedtoSubmitWorkbook to determine the request status
def readproceedtosubmitflag():
    wkbook = xlrd.open_workbook(ProceedTosubmitWorkBook)
    sheetname = wkbook.sheet_by_index(0)
    proceedtosubmitflag = sheetname.cell_value(1, 0)
    requesttime = sheetname.cell_value(1, 1)
    excelcontent = {'proceedtosubmitflag':proceedtosubmitflag,'requesttime':requesttime}
    return excelcontent

#Purpose: To write into ProceedtoSubmitWorkbook with flag value
def writeproceedtosubmitflag(flag,starttime):
    openbook = xlrd.open_workbook(ProceedTosubmitWorkBook)
    wb = copy(openbook) #xlutil is being used here
    sheetname = wb.get_sheet("Sheet1")
    sheetname.write(1, 0,flag)
    sheetname.write(1, 1, starttime)
    wb.save(ProceedTosubmitWorkBook)

def killProcess():
   p = Popen("KillProcessByName.bat", cwd=r"C:\\Users\\automationqateam\\Project\\venv2\\src")
   stdout, stderr = p.communicate()

def EmailTrigger():
    subprocess.Popen(["wscript.exe", EmailTriggerScriptWorkflow], stdout=subprocess.PIPE) #asynchronous
    #os.system("wscript.exe C:\\Users\\automationqateam\\Project\\venv2\\static_cdn\\vbs\\AP_AutoEmailTrigger.vbs")

#purpose: To read the updated RM data result and concatenate the client request and then return to the calling function
#It also triggers the Email if all the rows in the RM Data is success
def contextRMdataupdateAndemailtrigger():
    iscompletedflag = False
    #########################################################################################
    # concatenating the details that has to be displayed on post
    contexttopost = {'RMenvironment': RMenvironment, 'Email': Email}
    contextRMData = readRMData()
    context = {}
    for i in [contexttopost, contextRMData]:
        context.update(i)
        ########################################################################################
    for item in contextRMData:
        if (contextRMData[item].upper() != "PASSED" or contextRMData[item].upper() != "FAILED"):
            iscompletedflag = False
            break
        else:
            iscompletedflag = True

    if iscompletedflag == True:
        EmailTrigger()  # trigger auto email

    return context