'==========================================================================
'
' VBScript Source File -- SFC_Deploy.vbs
'
' NAME: SFC IE8 upgrade
'
' AUTHOR: Johnny Lai
' DATE  : 2017/03/17
' Ver :	v2.0
' COMMENT: 
'
'               #####################################################
'               #                                                   #
'               #                       _oo0oo_                     #
'               #                      o8888888o                    #
'               #                      88" . "88                    #
'               #                      (| -_- |)                    #
'               #                      0\  =  /0                    #
'               #                    ___/`---'\___                  #
'               #                  .' \\|     |# '.                 #
'               #                 / \\|||  :  |||# \                #
'               #                / _||||| -:- |||||- \              #
'               #               |   | \\\  -  #/ |   |              #
'               #               | \_|  ''\---/''  |_/ |             #
'               #               \  .-\__  '-'  ___/-. /             #
'               #             ___'. .'  /--.--\  `. .'___           #
'               #          ."" '<  `.___\_<|>_/___.' >' "".         #
'               #         | | :  `- \`.;`\ _ /`;.`/ - ` : | |       #
'               #         \  \ `_.   \_ __\ /__ _/   .-` /  /       #
'               #     =====`-.____`.___ \_____/___.-`___.-'=====    #
'               #                       `=---='                     #
'               #     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   #
'               #                                                   #
'               #               佛祖保佑         永無BUG            #
'               #                                                   #
'               #####################################################
'==========================================================================
On Error Resume Next



Dim objNetwork
Dim userName
Dim FSO
strPath = "C:\SFC_D\"
strLogPath = "C:\SFCS\deploy_log\"
strLogFile = strLogPath & "AutoDeploy.txt"


'statusflag = 0 copy file to c:\sfc_d from web
'statusflag = 1 resolved
'statusflag = 2 resolved
statusflag = 0

Set logfso = CreateObject("Scripting.FileSystemObject")

'Removed log file if exist, We want keep one version
If logfso.FileExists(strLogFile) Then 
	logfso.DeleteFile strLogFile
End If 



Set objWMIService = GetObject("winmgmts:")
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='c:'")
intDiskSpace = int(objLogicalDisk.FreeSpace/1048576)

'Cerate Log folder
If logfso.FolderExists(strLogPath) = 0 Then
	Set shl = CreateObject("WScript.Shell")
	Call shl.Run("%COMSPEC% /c mkdir """&strLogPath&"""",0,true)
	set shl = nothing
end if

'init log 
Set txtStream = logfso.OpenTextFile(strLogFile, 8, True)
txtStream.WriteLine(NOW() & "  program is running...") 




'Download file from web function
Function Download(strUrl,strPath,strFile)
	Set xHttp = CreateObject("MSXML2.ServerXMLHTTP")
	xHttp.Open "GET",strUrl,0
	xHttp.Send()
	
	Set bStrm= CreateObject("ADODB.Stream")
	
	with bStrm
		.type = 1 '//binary
		.open
		.write xHttp.responseBody
		strSave = strPath&strFile
		.savetofile strSave, 2 '//overwrite
	end with
End Function

'Check software install status function
Function SoftCheck(strTxt)
	Const HKEY_LOCAL_MACHINE = &H80000002
	strComputer = "."
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
		strComputer & "\root\default:StdRegProv")
	RegKey= strTxt
	strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

	instNet = 0
	For Each subkey In arrSubKeys
		instNet  = InStr(subkey, strTxt)  
		if instNet <> 0 then exit for
	Next
'	
	SoftCheck = instNet
End Function


Function win7_SoftCheck(strTxt,strKeyPath)
	Const HKEY_LOCAL_MACHINE = &H80000002
	strComputer = "."
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
		strComputer & "\root\default:StdRegProv")
	RegKey= strTxt
	'strKeyPath = "SOFTWARE\Microsoft\Windows\.NETFramework"
	oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

	instNet = 0
	For Each subkey In arrSubKeys
		instNet  = InStr(subkey, strTxt)  
		if instNet <> 0 then exit for
	Next
'	
	win7_SoftCheck = instNet
End Function

'Check file downloaded or not
Function Sleeping(fg,strItem)
	
	Select Case fg
		Case 0
			'down load patch
			intLoop_number = 0
			intLoop_Count = 0			

			
			Set ckfso = CreateObject("Scripting.FileSystemObject")
			do while intLoop_number > 0				
				If (ckfso.FileExists(strItem)) Then 				
					exit do
				else
					'sleep for 5 seconds
					WScript.Sleep(5000)
				
					intLoop_Count = intLoop_Count + 1
					if (intLoop_Count => 20) then 
						Sleeping = 1
						MsgBox("file" & strItem & "couldn't download in 100 seconds") 'download file time out
						exit do
					end if
				end if
								
			Loop 
			Set ckfso = nothing
			
		Case 1
			'install patch section resolved
			
		Case 2
			'remvoe file section resolved
End Select

 
 
End Function


Function GetOsBits()
   Set shell = CreateObject("WScript.Shell")
   If shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%") = "AMD64" Then
      GetOsBits = 1
   Else
      GetOsBits = 0
   End If
End Function



if intDiskSpace < 1000 then
	txtStream.WriteLine(NOW() & "  program is terminated becasue Disk Space is not enough 1Gb") 
	wscript.quit
end if

'get sleep
Set objRandom = CreateObject( "System.Random" )
intRanNumber = objRandom.Next_2( 0, 24 ) 
intRanNumber  = (intRanNumber*1000*60)*5 '1000 means 1 second, x5 means every number of 5 min in 2hr
'wscript.echo intRanNumber/1000 'debug to check how many minutes will be sleep
txtStream.WriteLine(NOW() & "  program will be activate after " &  intRanNumber/1000 &" Seconds")
WScript.Sleep intRanNumber


'================================== Main code ===============================================

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")


For Each objOperatingSystem in colOperatingSystems
	  'Check OS ver, the program only install on XP platform
      If instr(objOperatingSystem.Caption, "Windows XP") Then
			
            'do stuff if this is Windows XP

			'initial software install status
			instNETFramwork = 0
			instKB = 0
			instIE8 = 0
			
			'check .Netframwork 4 whether installed
			strTxt = "Microsoft .NET Framework 4"
			instNETFramwork = SoftCheck(strTxt)
			
			txtStream.WriteLine(NOW() & "  Initial Check All package install status before donwload process")
			if instNETFramwork = 0 then
				txtStream.WriteLine(NOW() & "  Check .netframword 4 install status : non install")
			else 
				txtStream.WriteLine(NOW() & "  Check .netframword 4 install status : installed already")
			end if
			
			'check KB whether installed
			strTxt = "KB2468871"
			instKB = SoftCheck(strTxt)
			if instKB = 0 then
				txtStream.WriteLine(NOW() & "  Check KB2468871 install status : non install")
			else 
				txtStream.WriteLine(NOW() & "  Check KB2468871 install status : installed already")
			end if
			
			'Check IE8 whether installed
			strTxt = "ie8"
			instIE8 = SoftCheck(strTxt)
			if instIE8 = 0 then
				txtStream.WriteLine(NOW() & "  Check instIE8 install status : non install") 
			else 
				txtStream.WriteLine(NOW() & "  Check instIE8 install status : installed already") 
			end if		
			
			txtStream.WriteLine(NOW() & "  End of Check All package install status before donwload process")
			'For debug
			'wscript.echo instNETFramwork
			'wscript.echo instKB
			'wscript.echo instIE8
			
			'Donwload update 
			if (instIE8) = 0 or  (instKB = 0) or (instNETFramwork = 0) then
			'wscript.echo instIE8
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set objNetwork = CreateObject("WScript.Network")
				userName = objNetwork.userName
				
				'Donwload file from web to C:\SFC_D
				If fso.FolderExists(strPath) Then

					strFile = "dotNetFx40_Full_x86_x64.exe"		
					
					If (fso.FileExists(strPath&strFile) = FALSE) Then
					
						Download "http://172.17.101.89/SFCS_Portal/Publish/SoftwareDownload?DirName=dotNetFx40_Full_x86_x64.exe&FileName=dotNetFx40_Full_x86_x64.exe",strPath,strFile
						call Sleeping(statusflag,strPath&strFile)
						txtStream.WriteLine(NOW() & "  " &  strFile &"  downloaded")
					end If
					
					strFile = "NDP40-KB2468871-v2-x86.exe"
					If (fso.FileExists(strPath&strFile) = FALSE) Then
						Download "http://172.17.101.89/SFCS_Portal/Publish/SoftwareDownload?DirName=NDP40-KB2468871-v2-x86.exe&FileName=NDP40-KB2468871-v2-x86.exe",strPath,strFile
						call Sleeping(statusflag,strPath&strFile)
						txtStream.WriteLine(NOW() & "  " &  strFile &"  downloaded")
					end If
					
					strFile = "IE8-WindowsXP-x86-CHT.exe"
					if instr(objOperatingSystem.Caption, "Windows XP") then
						If (fso.FileExists(strPath&strFile) = FALSE) Then
							Download "http://172.17.101.89/SFCS_Portal/Publish/SoftwareDownload?DirName=IE8-WindowsXP-x86-CHT.exe&FileName=IE8-WindowsXP-x86-CHT.exe",strPath,strFile
							call Sleeping(statusflag,strPath&strFile)
							txtStream.WriteLine(NOW() & "  " &  strFile &"  downloaded")
						end If	
					end if
					statusflag = statusflag + 1
				else
					fso.CreateFolder(strPath)
					txtStream.WriteLine(NOW() &"  "& strPath &" folder created")
					
					strFile = "dotNetFx40_Full_x86_x64.exe"
					Download "http://172.17.101.89/SFCS_Portal/Publish/SoftwareDownload?DirName=dotNetFx40_Full_x86_x64.exe&FileName=dotNetFx40_Full_x86_x64.exe",strPath,strFile
					call Sleeping(statusflag,strPath&strFile)
					txtStream.WriteLine(NOW() & "  " & strFile &"  downloaded")
					
					strFile = "NDP40-KB2468871-v2-x86.exe"
					Download "http://172.17.101.89/SFCS_Portal/Publish/SoftwareDownload?DirName=NDP40-KB2468871-v2-x86.exe&FileName=NDP40-KB2468871-v2-x86.exe",strPath,strFile
					call Sleeping(statusflag,strPath&strFile)
					txtStream.WriteLine(NOW() & "  " &  strFile &"  downloaded")
					
					
					If instr(objOperatingSystem.Caption, "Windows XP") then
						strFile = "IE8-WindowsXP-x86-CHT.exe"
						Download "http://172.17.101.89/SFCS_Portal/Publish/SoftwareDownload?DirName=IE8-WindowsXP-x86-CHT.exe&FileName=IE8-WindowsXP-x86-CHT.exe",strPath,strFile
						call Sleeping(statusflag,strPath&strFile)
						txtStream.WriteLine(NOW() & "  " &  strFile &"  downloaded")
					End if
					statusflag = statusflag + 1
				End If 'end of download patch
			end if 'end of IE8 install 

				Set WshShell = CreateObject("WScript.Shell")
				

				'install .Netfromwork
				if instNETFramwork = 0 then 
					strFile = "dotNetFx40_Full_x86_x64.exe" 
					strCommand = strPath&strFile& " /q /norestart "
					
					txtStream.WriteLine(NOW() & "  " &  strFile &" is going to install")
					call WshShell.Run (strCommand,1,true)
			
					'recheck Microsoft .NET Framework 4 install status
					strTxt = "Microsoft .NET Framework 4"
					instNETFramwork = SoftCheck(strTxt)				
					if instNETFramwork = 0 then
						txtStream.WriteLine(NOW() & "  Check .netframword 4 installed status : install failed")
					else 
						txtStream.WriteLine(NOW() & "  Check .netframword 4 installed status : install succeed")
					end if
					
					'txtStream.WriteLine(NOW() &  "  " & strFile &" install process is done")
				end if ' end 'install .Netfromwork
				
				'isntall KB2468871
				if instKB = 0 then 
					strFile = "NDP40-KB2468871-v2-x86.exe"
					strCommand = strPath&strFile& " /q /norestart "
					txtStream.WriteLine(NOW() & "  " &  strFile &" is going to install")
					call WshShell.Run (strCommand,1,true)
					
					'recheck KB2468871 install status
					strTxt = "KB2468871"
					instKB = SoftCheck(strTxt)
					if instKB = 0 then
						txtStream.WriteLine(NOW() & "  Check KB2468871 install status : install failed")
					else 
						txtStream.WriteLine(NOW() & "  Check KB2468871 install status : install succeed")
					end if						
					
					'txtStream.WriteLine(NOW() & "  " &  strFile &" install process is done")

				end if ' end of install KB2468871

				'isntall IE8
				if instr(objOperatingSystem.Caption, "Windows XP") then
					if instIE8 = 0 then
						strFile = "IE8-WindowsXP-x86-CHT.exe"
						strCommand = strPath&strFile& "  /quiet /update-no /norestart "

						txtStream.WriteLine(NOW() & "  " &  strFile &" is going to install")
						call WshShell.Run (strCommand,1,true)			
					
				
 
						'reCheck IE8  installed status
						strTxt = "ie8"
						instIE8 = SoftCheck(strTxt)
						if instIE8 = 0 then
							txtStream.WriteLine(NOW() & "  Check instIE8 install status : install failed") 
						else 
							txtStream.WriteLine(NOW() & "  Check instIE8 install status : install succeed")
							wscript.echo "IE8 完成安裝, 請手動重開機!! 謝謝您~~"
						end if	
					'txtStream.WriteLine(NOW() & "  " &  strFile &" install process is done")
					End if 
				End if ' end of install IE8 
				
				
				set WshShell=Nothing
				
				
				'If fso.FolderExists(strPath) Then  
				'	set deletefolder = fso.GetFolder(strPath)
				'	deletefolder.Delete(True) 
			
				'end if
				
				Set fso = Nothing
				
			

			'===============================================
			'remove install files
			txtStream.WriteLine(NOW() & "  Deleting source folder and files : " & strPath)
			Set fso = CreateObject("Scripting.FileSystemObject")
			If fso.FolderExists(strPath) Then  
				set deletefolder = fso.GetFolder(strPath)
				deletefolder.Delete(True) 
			end if
			txtStream.WriteLine(NOW() & "  source folder and files " & strPath & " are Deleted")
		elseif instr(objOperatingSystem.Caption, "Windows 7") Then
'windows 7 Here =======================================================================
			'initial software install status
			instNETFramwork = 0
			instKB = 0
			'init Windows 7 x32 or x64 -- 0 = x32 ; 1 = 64
			W7ver = 0
			Dim strTxt,strKeyPath

			'Check Win7 x32 or x64 -- 0 = x32 ; 1 = 64
			W7ver = GetOsBits
			
			
			if W7ver = 1 then
				'check .Netframwork 4 whether installed
				strTxt = "v4"
				strKeyPath = "SOFTWARE\Wow6432Node\Microsoft\.NETFramework"
				'check Win7 X64 first
				instNETFramwork = win7_SoftCheck(strTxt,strKeyPath)	
			else
				strTxt = "v4"
				strKeyPath = "SOFTWARE\Microsoft\.NETFramework"
				instNETFramwork = win7_SoftCheck(strTxt,strKeyPath)	
			end if			
			
			'wscript.echo instNETFramwork
			if instNETFramwork = 0 then
				txtStream.WriteLine(NOW() & "  Check .netframword 4 install status : non install")
			else 
				txtStream.WriteLine(NOW() & "  Check .netframword 4 install status : installed already")
			end if
			'wscript.echo instNETFramwork
			
			if W7ver = 1 then
				'check KB2468871 whether installed
				strTxt = "KB2468871"
				strKeyPath = "SOFTWARE\Wow6432Node\Microsoft\Updates\Microsoft .NET Framework 4 Client Profile"
				instKB = win7_SoftCheck(strTxt,strKeyPath)			
			else
				strTxt = "KB2468871"
				strKeyPath = "SOFTWARE\Microsoft\Updates\Microsoft .NET Framework 4 Client Profile"
				instKB = win7_SoftCheck(strTxt,strKeyPath)	
			end if
				
			
			if instKB = 0 then
				txtStream.WriteLine(NOW() & "  Check KB2468871 install status : non install")
			else 
				txtStream.WriteLine(NOW() & "  Check KB2468871 install status : installed already")
			end if			
			'wscript.echo instKB
			
			if (instKB = 0) or (instNETFramwork = 0) then
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set objNetwork = CreateObject("WScript.Network")
				userName = objNetwork.userName
				
				'Create C:\SFC_D folder 
				If fso.FolderExists(strPath) Then
					'debug
				else
					fso.CreateFolder(strPath)
					txtStream.WriteLine(NOW() &"  "& strPath &" folder created")
				end if
				
				strFile = "dotNetFx40_Full_x86_x64.exe"							
				If (fso.FileExists(strPath&strFile) = FALSE) Then					
					Download "http://172.17.101.89/SFCS_Portal/Publish/SoftwareDownload?DirName=dotNetFx40_Full_x86_x64.exe&FileName=dotNetFx40_Full_x86_x64.exe",strPath,strFile
					call Sleeping(statusflag,strPath&strFile)
					txtStream.WriteLine(NOW() & "  " &  strFile &"  downloaded")
				end If
					
				if W7ver = 0 then
					strFile = "NDP40-KB2468871-v2-x86.exe"
				else
					strFile = "NDP40-KB2468871-v2-x64.exe"
				end if
				

				If (fso.FileExists(strPath&strFile) = FALSE) Then				    
					Download "http://172.17.101.89/SFCS_Portal/Publish/SoftwareDownload?DirName="&strFile&"&FileName="&strFile,strPath,strFile
					call Sleeping(statusflag,strPath&strFile)
					txtStream.WriteLine(NOW() & "  " &  strFile &"  downloaded")
				end If
				
				Set WshShell = CreateObject("WScript.Shell")
				
				if instNETFramwork = 0 then
					strFile = "dotNetFx40_Full_x86_x64.exe"
				    
					strCommand = strPath&strFile& " /q /norestart "

					txtStream.WriteLine(NOW() & "  " &  strFile &" is going to install")
					call WshShell.Run (strCommand,1,true)
			
					'recheck Microsoft .NET Framework 4 install status
					if W7ver = 1 then
						'check .Netframwork 4 whether installed
						strTxt = "v4"
						strKeyPath = "SOFTWARE\Wow6432Node\Microsoft\.NETFramework"
						'check Win7 X64 first
						instNETFramwork = win7_SoftCheck(strTxt,strKeyPath)	
					else
						strTxt = "v4"
						strKeyPath = "SOFTWARE\Microsoft\.NETFramework"
						instNETFramwork = win7_SoftCheck(strTxt,strKeyPath)	
					end if	
				
					
					if instNETFramwork = 0 then
						txtStream.WriteLine(NOW() & "  Check .netframword 4 installed status : install failed")
					else 
						txtStream.WriteLine(NOW() & "  Check .netframword 4 installed status : install succeed")
					end if
				 end if
				 
				 if instKB = 0 then
					if W7ver = 0 then
						strFile = "NDP40-KB2468871-v2-x86.exe"
					else
						strFile = "NDP40-KB2468871-v2-x64.exe"
					end if
					
				    
					strCommand = strPath&strFile& "  /q /update-no /norestart "
					
					
					txtStream.WriteLine(NOW() & "  " &  strFile &" is going to install")
					call WshShell.Run (strCommand,1,true)
					
					're-check KB2468871 whether installed
					if W7ver = 1 then
						
						strTxt = "KB2468871"
						strKeyPath = "SOFTWARE\Wow6432Node\Microsoft\Updates\Microsoft .NET Framework 4 Client Profile"
						instKB = win7_SoftCheck(strTxt,strKeyPath)						
					else
						strTxt = "KB2468871"
						strKeyPath = "SOFTWARE\Microsoft\Updates\Microsoft .NET Framework 4 Client Profile"
						instKB = win7_SoftCheck(strTxt,strKeyPath)	
					end if
				

					
					if instKB = 0 then
						txtStream.WriteLine(NOW() & "  Check "& strFile &" installed status : install failed")
					else 
						txtStream.WriteLine(NOW() & "  Check " & strFile &" installed status : install succeed")
						wscript.echo "系統更新完成安裝, 請手動重開機!! 謝謝您~~"
					end if					
				 end if
				
				
				'================================End of Windows 7 main program ============================================
				'delete C:\SFC_D folder
				If fso.FolderExists(strPath) Then  
					set deletefolder = fso.GetFolder(strPath)
					deletefolder.Delete(True) 
				end if
				txtStream.WriteLine(NOW() & "  source folder and files " & strPath & " are Deleted")
				
				set fso = nothing
				set WshShell = nothing
				set objNetwork = nothing
				
			end if

			
			
			
			
		else
			txtStream.WriteLine(NOW() & "  OS is not XP or Windows 7, The install is not necessary.")
		End If  ' end of program 
Next
txtStream.WriteLine(NOW() & "  program is terminated") 
Set txtStream = nothing
set logfso = nothing