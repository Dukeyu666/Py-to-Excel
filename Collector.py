from logging import exception
import winreg,os,shutil,uptime
import socket
from datetime import datetime,timedelta
from dateutil import tz
from win32com.client import Dispatch
from configparser import ConfigParser


config=ConfigParser(interpolation=None)

# Boot time
LastBootTime=uptime.boottime().strftime("%Y-%m-%d %H:%M:%S")
#Uptime=float("%.2f"%(uptime.uptime()/60/60/24))

# Defrag part

try:
    scheduler = Dispatch('Schedule.Service')
    scheduler.Connect()
    folders = scheduler.GetFolder('\\Microsoft\\Windows\\Defrag')
    tasks=folders.GetTasks(1)
    lastruntime=tasks[0].lastruntime.strftime("%Y-%m-%d %H:%M:%S")
    # print(tasks[0].lastruntime.strftime("%Y-%m-%d %H:%M:%S"))
except Exception as e:
    lastruntime = " "
    print(e.__str__()+": Last Defrag time")

# AV part
registry=winreg.ConnectRegistry(None,winreg.HKEY_LOCAL_MACHINE)
try:
    key=winreg.OpenKey(registry,"SOFTWARE\\WOW6432Node\\TrendMicro\\PC-cillinNTCorp\\CurrentVersion\\UpdateInfo")
    LastAVupdate=datetime.fromtimestamp(winreg.QueryValueEx(key,"P.48020000")[0]).strftime("%Y-%m-%d %H:%M:%S")
except Exception as e:
    LastAVupdate = " "
    print(e.__str__()+": Last AV update")
try:
    key=winreg.OpenKey(registry,"SOFTWARE\\WOW6432Node\\TrendMicro\\PC-cillinNTCorp\\CurrentVersion\\ScanOperationLog")
    LastScanTime=datetime.strptime(str(winreg.QueryValueEx(key,"StartDate")[0]),"%Y%m%d").strftime("%Y-%m-%d %H:%M:%S")
except Exception as e:
    LastScanTime = " "
    print(e.__str__()+": Last Scan time")
#windows update part
try:
    COMobject_AutoUpdate = Dispatch('Microsoft.Update.AutoUpdate')
    LastWindowsUpdate = COMobject_AutoUpdate.Results.LastInstallationSuccessDate
    tzlocal = tz.gettz('Asia\Taipei')
    LastWindowsUpdate = datetime.strftime(datetime.astimezone(LastWindowsUpdate,tzlocal),"%Y-%m-%d %H:%M:%S")
    # key=winreg.OpenKey(registry,"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install")
    # LastWindowsUpdate=(datetime.strptime(winreg.QueryValueEx(key,"LastSuccessTime")[0],"%Y-%m-%d %H:%M:%S")+timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
except Exception as e :
    LastWindowsUpdate = " "
    print(e.__str__()+": Last Windows Update")

winreg.CloseKey(registry)
winreg.CloseKey(key)

#
#config["Uptime(days)"]={"Uptime(days)":Uptime} 
config["LastBootTime"]={"LastBootTime":LastBootTime}
config["LastDefragTime"]={"lastdefragtime":lastruntime}
config["LastAVUpdate"]={"LastAVUpdate":LastAVupdate}
config["LastScanTime"]={"LastScanTime":LastScanTime}
config["LastWindowsUpdate"]={"LastWindowsUpdate":LastWindowsUpdate}

drive_letter=["C:","D:","E:"]
# Tools part

for drive in drive_letter:
    if os.path.exists(drive+r"\Tools\ATF"):
        information_parser = Dispatch("Scripting.FileSystemObject")
        ATF = information_parser.GetFileVersion(drive+r"\Tools\ATF\ATF.exe")
        config["ATF"]={"ATF":ATF}
    if os.path.exists(drive+r"\FSDASH\diags\Dummy_Diags_Version.txt"):
        with open(file=drive+r"\FSDASH\diags\Dummy_Diags_Version.txt",mode="r") as f:
            LAPS = f.read()
        config["LAPS"]={"LAPS":LAPS.strip()}
    if os.path.exists(drive+r"\FSDASH\WDT\DPSVERSION.txt"):
        with open(file=drive+r"\FSDASH\WDT\DPSVERSION.txt",mode="r") as f:
            DPS = f.read()
        config["DPS"]={"DPS":DPS}    
    if os.path.exists(drive+r"\FSDASH\WDT\UBER.ini"):
        uberini=ConfigParser()
        uberini.read(drive+r"\FSDASH\WDT\UBER.ini")
        WDT = uberini.get("ServerToolsetVersion","MainToolSet")
        config["WDT"]={"WDT":WDT}
    if os.path.exists(drive+r"\fsdash\WDT\MENU\UnlockAll_Dev\ExtractAssembly\amd64"):
        information_parser = Dispatch("Scripting.FileSystemObject")
        ExtractAssembly = information_parser.GetFileVersion(drive+r"\fsdash\WDT\MENU\UnlockAll_Dev\ExtractAssembly\amd64\ExtractAssembly.exe")
        config["ExtractAssembly"]={"ExtractAssembly":ExtractAssembly}
    if os.path.exists(drive+r"\Tools\PCR\PCRMANAGER\PcrMgrNet.exe"):
        information_parser = Dispatch("Scripting.FileSystemObject")
        PCR = information_parser.GetFileVersion(drive+r"\Tools\PCR\PCRMANAGER\PcrMgrNet.exe")
        config["PCR"]={"PCR":PCR}
    if os.path.exists(drive+r"\PO_IN\POConfig.exe"):
        information_parser = Dispatch("Scripting.FileSystemObject")
        PoConfig = information_parser.GetFileVersion(drive+r"\PO_IN\POConfig.exe")
        config["PoConfig"]={"PoConfig":PoConfig}

# Disks part

config.add_section("Disks Free Rate")
for drive in drive_letter :
    if os.path.exists(drive) :
        FreeRate="%.2f" % \
        (( ((shutil.disk_usage(drive).free) / 2**30) / ((shutil.disk_usage(drive).total) / 2**30) )*100)
        
        config["Disks Free Rate"].update({drive.replace(":",""):FreeRate})
        

with open(socket.gethostname()+".ini",mode="w") as f:
    config.write(f)