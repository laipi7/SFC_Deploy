'==========================================================================
'
' VBScript Source File -- SFC_IE8_Deploy.vbs
'
' NAME: SFC IE8 upgrade
'
' AUTHOR: Johnny Lai
' DATE  : 2017/01/17
'
' COMMENT: 
'
'==========================================================================
On Error Resume Next



Dim objNetwork
Dim userName
Dim FSO
strPath = "C:\SFC_D\"
strLogPath = "C:\SFCS\deploy_log\"
strLogFile = strLogPath & "IE8_Install_Log.txt"


'statusflag = 0 copy file to c:\sfc_d from web
'statusflag = 1 resolved
'statusflag = 2 resolved
statusflag = 0

Set logfso = CreateObject("Scripting.FileSystemObject")



'Cerate Log folder
If logfso.FolderExists(strLogPath) = 0 Then
	Set shl = CreateObject("WScript.Shell")
	Call shl.Run("%COMSPEC% /c mkdir """&strLogPath&"""",0,true)
	set shl = nothing
end if

'
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
			
			if instNETFramwork = 0 then
				txtStream.WriteLine(NOW() & "  Check .netframword 4 installed status : non install")
			else 
				txtStream.WriteLine(NOW() & "  Check .netframword 4 installed status : installed already")
			end if
			
			'check KB whether installed
			strTxt = "KB2468871"
			instKB = SoftCheck(strTxt)
			if instKB = 0 then
				txtStream.WriteLine(NOW() & "  Check KB2468871 installed status : non install")
			else 
				txtStream.WriteLine(NOW() & "  Check KB2468871 installed status : installed already")
			end if
			
			'Check IE8 whether installed
			strTxt = "ie8"
			instIE8 = SoftCheck(strTxt)
			if instIE8 = 0 then
				txtStream.WriteLine(NOW() & "  Check instIE8 installed status : non install, program is going to install process") 
			else 
				txtStream.WriteLine(NOW() & "  Check instIE8 installed status : installed already, program is going to end") 
			end if			
			'For debug
			'wscript.echo instNETFramwork
			'wscript.echo instKB
			'wscript.echo instIE8
			
			'Donwload update 
			if instIE8 = 0 then
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
					If (fso.FileExists(strPath&strFile) = FALSE) Then
						Download "http://172.17.101.89/SFCS_Portal/Publish/SoftwareDownload?DirName=IE8-WindowsXP-x86-CHT.exe&FileName=IE8-WindowsXP-x86-CHT.exe",strPath,strFile
						call Sleeping(statusflag,strPath&strFile)
						txtStream.WriteLine(NOW() & "  " &  strFile &"  downloaded")
					end If	
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
					
					strFile = "IE8-WindowsXP-x86-CHT.exe"
					Download "http://172.17.101.89/SFCS_Portal/Publish/SoftwareDownload?DirName=IE8-WindowsXP-x86-CHT.exe&FileName=IE8-WindowsXP-x86-CHT.exe",strPath,strFile
					call Sleeping(statusflag,strPath&strFile)
					txtStream.WriteLine(NOW() & "  " &  strFile &"  downloaded")
					
					statusflag = statusflag + 1
				End If 'end of download patch


				Set WshShell = CreateObject("WScript.Shell")
				

				'install .Netfromwork
				if instNETFramwork = 0 then 
					strFile = "dotNetFx40_Full_x86_x64.exe" 
					strCommand = strPath&strFile& " /q /norestart "
					
					txtStream.WriteLine(NOW() & "  " &  strFile &" is going to install")
					call WshShell.Run (strCommand,1,true)
					txtStream.WriteLine(NOW() &  "  " & strFile &" install process is finished")
				end if
				
				'isntall KB
				if instKB = 0 then 
					strFile = "NDP40-KB2468871-v2-x86.exe"
					strCommand = strPath&strFile& " /q /norestart "
					txtStream.WriteLine(NOW() & "  " &  strFile &" is going to install")
					call WshShell.Run (strCommand,1,true)
					txtStream.WriteLine(NOW() & "  " &  strFile &" install process is finished")
				end if

				'isntall IE8
				strFile = "IE8-WindowsXP-x86-CHT.exe"
				strCommand = strPath&strFile& "  /quiet /update-no /norestart "
				
				txtStream.WriteLine(NOW() & "  " &  strFile &" is going to install")
				call WshShell.Run (strCommand,1,true)			
				txtStream.WriteLine(NOW() & "  " &  strFile &" install process is finished")	
				
				set WshShell=Nothing
				
				
				'If fso.FolderExists(strPath) Then  
				'	set deletefolder = fso.GetFolder(strPath)
				'	deletefolder.Delete(True) 
			
				'end if
				
				Set fso = Nothing
				
			else	
				'IE8 Installed
				txtStream.WriteLine(NOW() & "  The install is not necessary.")
			end if 'end of IE8 install 

			'remove install files
			txtStream.WriteLine(NOW() & "  Deleting source folder and files : " & strPath)
			Set fso = CreateObject("Scripting.FileSystemObject")
			If fso.FolderExists(strPath) Then  
				set deletefolder = fso.GetFolder(strPath)
				deletefolder.Delete(True) 
			end if
			txtStream.WriteLine(NOW() & "  source folder and files " & strPath & " are Deleted")
		else
			txtStream.WriteLine(NOW() & "  OS is not XP, The install is not necessary.")
		End If  ' end of XP 
		
		

Next
txtStream.WriteLine(NOW() & "  program is terminated") 
Set txtStream = nothing
set logfso = nothing