# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from django.http import HttpResponse
from django.shortcuts import render
import win32com,win32com.client
import os,ctypes,time
import xlwt,xlrd,subprocess
from xlutils.copy import copy
from win32com.client import Dispatch
from django.conf import settings
# Create your views here.

global WorkBookName,AutomationPortal_QCWorkflow,EmailTriggerScriptPath
WorkBookName = os.path.join(settings.STATIC_ROOT, 'data/RMData.xls')

AutomationPortal_QCWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AutomationPortal_QCWorkflow.vbs')
EmailTriggerScriptWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AP_AutoEmailTrigger.vbs')

#def EmailTrigger():
    #subprocess.Popen(["wscript.exe", EmailTriggerScriptWorkflow], stdout=subprocess.PIPE)
    #os.system("wscript.exe C:\\Users\\automationqateam\\Project\\venv2\\static_cdn\\vbs\\AP_AutoEmailTrigger.vbs")

def index(request):
    global RMenvironment1, RMscriptname1, RMmachine1, RMexecute1
    global RMenvironment2, RMscriptname2, RMmachine2, RMexecute2
    global RMenvironment3, RMscriptname3, RMmachine3, RMexecute3
    global RMenvironment4, RMscriptname4, RMmachine4, RMexecute4
    global RMSanity,Email

    if (request.method =='POST') and (request.POST.get('Execute')=="Execute"):
           RMSanity = request.POST.get('RMSanity')
           Email = request.POST.get('email')
           #ctypes.windll.user32.MessageBoxW(0, RMSanity, "Your Title", 1)
           if RMSanity == "RMSanity":
               #ctypes.windll.user32.MessageBoxW(0, Email, "Your Title", 1)
               RMenvironment = request.POST.get('RMenvironment')
               #RMenvironment1 = request.POST.get('RMenvironment1')
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


               #writeRMData()     #----------------write to newly created workbook----------------#
               #vbscriptcall()    #----------------vbscript to connect QC and invoke script ---------------#
               contexttopost = {'RMenvironment':RMenvironment}
               contextRMData = readRMData()
               context = {}
               for i in [contexttopost,contextRMData]:
                   context.update(i)

               return render(request, 'automationUI/indexafterupdate.html', context)



               #EmailTrigger()

    else:
        return render(request, 'automationUI/index.html')

def readRMData():
    wkbook = xlrd.open_workbook(WorkBookName,'r')
    sheetname = wkbook.sheet_by_index(0)
    RM1statusval = sheetname.cell_value(1, 8)
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
    #for i in range(rowcount-1):
    for i in range(4):
             subprocess.Popen(["wscript.exe",AutomationPortal_QCWorkflow],stdout=subprocess.PIPE)
             if i==1:
                 time.sleep(90)
             else:
                 time.sleep(10)
             #ctypes.windll.user32.MessageBoxW(0,"Middle","Your Title",1)

def writeRMData():
    wb = xlwt.Workbook()
    ws = wb.add_sheet('data')

    TestPlanFolderPath = "[QualityCenter]Subject\Automation\ECO\Products\AutomationPortal"
    TestSetFolderPath = "Root\ECO Automation\Product-Execution\AutomationPortal"
    SanityName = "ReleaseManagement"
    rt_IsRunning = "No"
    rt_IsStop = "No"
    Email = "ashiskumar@ap.equinix.com"
    ExecutorId = "All"

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

        ws.write(1, 6, rt_IsRunning)
        ws.write(2, 6, rt_IsRunning)
        ws.write(3, 6, rt_IsRunning)
        ws.write(4, 6, rt_IsRunning)

        ws.write(1, 7, rt_IsStop)
        ws.write(2, 7, rt_IsStop)
        ws.write(3, 7, rt_IsStop)
        ws.write(4, 7, rt_IsStop)

        ws.write(1, 10, RMenvironment1)
        ws.write(2, 10, RMenvironment2)
        ws.write(3, 10, RMenvironment3)
        ws.write(4, 10, RMenvironment4)

        ws.write(0, 11, "MailTo")
        ws.write(1, 11, Email)
        ws.write(0, 12, "ExecutorId")
        ws.write(1, 12, ExecutorId)

        wb.save(os.path.join(settings.STATIC_ROOT, 'data/RMData.xls'))










#def appendingRMRequestData():
#    #the lines to get the number of rows in the excel sheet
#    wkbook = xlrd.open_workbook(WorkBookName)
#    sheetnameforrowcount = wkbook.sheet_by_index(0)
#    #colcount = sheetname.ncols
#    rowcount = int(sheetnameforrowcount.nrows)

#    #the lines to write to excel sheet
#    openbook = xlrd.open_workbook(WorkBookName)
#    wb = copy(openbook) #xlutil is being used here
#    sheetname = wb.get_sheet("AutomationPortal")
#
#    TestPlanFolderPath = "[QualityCenter]Subject\Automation\ECO\Products\AutomationPortal"
#    TestSetFolderPath = "Root\ECO Automation\Product-Execution\AutomationPortal"
#    SanityName = "ReleaseManagement"
#    rt_IsRunning = "No"
#    rt_IsStop = "No"
#    Email = "ashiskumar@ap.equinix.com"
#    ExecutorId = "All"
#    if RMSanity == "RMSanity":
#        sheetname.write(rowcount, 0, RMexecute1)
#        sheetname.write(rowcount+1, 0, RMexecute2)
#        sheetname.write(rowcount+2, 0, RMexecute3)
#        sheetname.write(rowcount+3, 0, RMexecute4)#
#
#        sheetname.write(rowcount, 1, TestPlanFolderPath)
#        sheetname.write(rowcount + 1, 1, TestPlanFolderPath)
#        sheetname.write(rowcount + 2, 1, TestPlanFolderPath)
#        sheetname.write(rowcount + 3, 1, TestPlanFolderPath)
#
#        sheetname.write(rowcount, 2, RMscriptname1)
#        sheetname.write(rowcount+1, 2, RMscriptname2)
#        sheetname.write(rowcount+2, 2, RMscriptname3)
#        sheetname.write(rowcount+3, 2, RMscriptname4)#
#
#        sheetname.write(rowcount, 3, TestSetFolderPath)
#        sheetname.write(rowcount+1, 3, TestSetFolderPath)
#        sheetname.write(rowcount+2, 3, TestSetFolderPath)
#        sheetname.write(rowcount+3, 3, TestSetFolderPath)
#
#        sheetname.write(rowcount, 4, SanityName)
#        sheetname.write(rowcount+1, 4, SanityName)
#        sheetname.write(rowcount+2, 4, SanityName)
#        sheetname.write(rowcount+3, 4, SanityName)
#
#        sheetname.write(rowcount, 5, RMmachine1)
#        sheetname.write(rowcount+1, 5, RMmachine2)
#        sheetname.write(rowcount+2, 5, RMmachine3)
#        sheetname.write(rowcount+3, 5, RMmachine4)
#
#        sheetname.write(rowcount, 6, rt_IsRunning)
#        sheetname.write(rowcount+1, 6, rt_IsRunning)
#        sheetname.write(rowcount+2, 6, rt_IsRunning)
#        sheetname.write(rowcount+3, 6, rt_IsRunning)
#
#        sheetname.write(rowcount, 7, rt_IsStop)
#        sheetname.write(rowcount+1, 7, rt_IsStop)
#        sheetname.write(rowcount+2, 7, rt_IsStop)
#        sheetname.write(rowcount+3, 7, rt_IsStop)
#
#        sheetname.write(rowcount, 10, RMenvironment1)
#        sheetname.write(rowcount+1, 10, RMenvironment2)
#        sheetname.write(rowcount+2, 10, RMenvironment3)
#        sheetname.write(rowcount+3, 10, RMenvironment4)
#    wb.save(WorkBookName)
#    #wb.save("\\\\sgw-filesvr3\\GDC\\GDC_Team\\QA\\ECO-Product-Docs\\Automation\\AutomationPortal\\QCWorkFlow\\AutomationPortal_Data.xls")




# os.system("wscript.exe \\\\sgw-filesvr3\\GDC\\GDC_Team\\QA\\ECO-Product-Docs\\Automation\\AutomationPortal\\QCWorkFlow\\AutomationPortal_QCWorkflow.vbs")
# subprocess.call("wscript.exe \\\\sgw-filesvr3\\GDC\\GDC_Team\\QA\\ECO-Product-Docs\\Automation\\AutomationPortal\\QCWorkFlow\\AutomationPortal_QCWorkflow.vbs",shell=False)
# subprocess.Popen(["wscript.exe","\\\\sgw-filesvr3\\GDC\\GDC_Team\\QA\\ECO-Product-Docs\\Automation\\AutomationPortal\\QCWorkFlow\\AutomationPortal_QCWorkflow.vbs"],stdout=subprocess.PIPE)
# os.system("wscript.exe \\\\sgw-filesvr3\\GDC\\GDC_Team\\QA\\ECO-Product-Docs\\Automation\\AutomationPortal\\QCWorkFlow\\Test.vbs")
# subprocess.call("wscript.exe","C:\\Windows\\SysWOW64\\AutomationPortal_QCWorkflow.vbs",shell=False)


#WorkBookName = "\\\\sgw-filesvr3\\GDC\\GDC_Team\\QA\\ECO-Product-Docs\\Automation\\AutomationPortal\\QCWorkFlow\\AutomationPortal_Data.xls"
#WorkBookName = os.path.join(settings.STATIC_ROOT, 'data/AutomationPortal_Data.xls')

#AutomationPortal_QCWorkflow = "\\\\sgw-filesvr3\\GDC\\GDC_Team\\QA\\ECO-Product-Docs\\Automation\\AutomationPortal\\QCWorkFlow\\AutomationPortal_QCWorkflow.vbs"


#               from timeit import default_timer as timer
#               NotCompletedFlag = False
#               totaltimeexpected = len(contextRMData)*60*60
#                now = timer()
#               while now < totaltimeexpected:
#                   contexttopost = {'RMenvironment': RMenvironment}
#                   contextRMData = readRMData()
#                   context = {}
#
#                   for i in [contexttopost, contextRMData]:
#                       context.update(i)
#
#                   for status in contextRMData:
#                       if contextRMData[status]=='Not Completed':
#                           NotCompletedFlag = True
#                           break
#                   if NotCompletedFlag == False:
#                       break
#                   now = timer()
#                   return render(request, 'automationUI/indexafterupdate.html', context)