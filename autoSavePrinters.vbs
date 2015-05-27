'Tanner Cornejo
'3/18/2015
'Finds all installed network printers and writes a script that saves to the user personal drive (default S:\).  
'Also creates a text file that lists the printers to be installed.  It is referenced by the script so they must be in the same location.
'When run, that script will reinstall the printers.
'Ideally this would be set up as a logon script.
'For use by Health Delivery Inc.

'todo: Validate printer name before installing. What happens when name is invalid?
'Notes:  Most commented code is for saving the printers to a text file, and changes the script to pull printer names from that file. 
'        It is obsolete and unused at the moment.

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshNetwork = CreateObject("WScript.Network")  
strName = WshNetwork.UserName
strComputer = WshNetwork.ComputerName
Dim strPrinter
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colInstalledPrinters =  objWMIService.ExecQuery("Select * from Win32_Printer Where Local = FALSE")
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8


'Change to wherever you want the install script and printer list to be saved.
Const userMappedDrive = "S:\"
fileOut = userMappedDrive & "All Network Printers.txt"
   
If Not colInstalledPrinters.Count = 0 Then
	'''''x=msgbox("Writing Script.  " & colInstalledPrinters.Count & " printer is installed.",0,"Write Attempting")
	'This section creates a script that tries to read the printer lists and adds them
	scrptName = strComputer & " networkPrinters.vbs"
	strTemp = userMappedDrive & scrptName
	
	'Delete old version, replace with new
	If objFSO.FileExists(strTemp) Then
		objFSO.DeleteFile(strTemp)
	End If
	Set objScrptFile = objFSO.CreateTextFile(strTemp, ForAppending, True)

	'strTemp = "Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer")"
	'objScriptFile.WriteLine(strTemp)
	
	'Write the script line by line.
	strTemp = "Set objFSO = CreateObject(""Scripting.FileSystemObject"")"
	objScrptFile.WriteLine(strTemp)
	strTemp = "Set WshNetwork = CreateObject(""WScript.Network"")"
	objScrptFile.WriteLine(strTemp)
	If objFSO.FileExists(fileOut) Then
		If objFSO.getFile(fileOut).size <> 0 Then
			arrLines = Split(objFSO.OpenTextFile(fileOut).ReadAll(), vbCrlf)
		End If
	Else
		Set objPrinterFile = objFSO.CreateTextFile(fileOut)
		ReDim arrLines(1)
		objPrinterFile.close
	End If
	
	'Add each installed printer to the list.
	For Each objPrinter in colInstalledPrinters  
		strPrinter = objPrinter.Name
		strTemp = "WshNetwork.AddWindowsPrinterConnection """ & strPrinter & """"
		objScrptFile.WriteLine(strTemp)
		write = True
		On Error Resume Next
		intUpper = Ubound(arrLines)
		If Err = 0 Then
			'Compares each installed printer with each line in the printer list.
			For Each line in arrLines
				strPrinter = objPrinter.Name
				'If it is already in the list, do not write it to the list.  This prevents duplicates.
				If line = strPrinter Then
					write = False
					strPrinter = line
				End If
			Next
			'Otherwise, add it to the list of printers to install.
			If write = True Then
				Set objPrinterFile = objFSO.OpenTextFile(fileOut, ForAppending)
				objPrinterFile.WriteLine(strPrinter)
				objPrinterFile.Close
			End If
		Else
			Set objPrinterFile = objFSO.OpenTextFile(fileOut, ForAppending)
			objPrinterFile.WriteLine(strPrinter)
			objPrinterFile.Close
			Err.Clear
		End If
	Next  
	strTemp = "x=msgbox(""Printer Installation Complete. "",0, ""Operation Complete"")"
	objScrptFile.WriteLine(strTemp)
	objScrptFile.Close
End If

'After run you will be left with a .vbs file that will reinstall the printers that you had installed when the script was run, 
'and a .txt file with a list of the printers to install.  This list is referenced by the .vbs file.
