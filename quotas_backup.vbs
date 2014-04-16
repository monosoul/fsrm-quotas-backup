Const ForReading = 1
Const ForWriting = 2
Const bWaitOnReturn = True

Set oShell = CreateObject("WScript.Shell")
SysDrive=oShell.ExpandEnvironmentStrings("%SystemDrive%")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives
CurrentDirectory = objFSO.GetAbsolutePathName(".")
quotasexist=0

'Getting quotas list
oShell.run "cmd /c ""dirquota q l > " & CurrentDirectory & "\all_quotas.txt""",0,bWaitOnReturn

'If quotas are exist then templates would be exported
Set objFileIn = objFSO.OpenTextFile(CurrentDirectory & "\all_quotas.txt", ForReading)
Do Until objFileIn.AtEndOfStream
	tmpline = objFileIn.ReadLine
	If (Len(tmpline) > 0) Then
		If (InStr(tmpline, "Quota Path:") = 1) Then quota_path = LTrim(Right(tmpline, Len(tmpline) - Len("Quota Path:")))
		If quota_path <> "" Then
			quotasexist=1
		End If
	End If
Loop

objFileIn.Close

If objFSO.FileExists(SysDrive & "\Windows\System32\dirquota.exe") And (quotasexist = 1) Then
	'Exporting templates
	oShell.run "dirquota template export /file:" & CurrentDirectory & "\quota_templates.xml",0,bWaitOnReturn
Else
	Set objFileQtemplates = objFSO.OpenTextFile(CurrentDirectory & "\quota_templates.xml", ForWriting, True)
	objFileQtemplates.Close
End If


' ==============================================================================================
'|																								|
'|							Generating batch file for quotas creation							|
'|																								|
' ==============================================================================================


Set objFileIn = objFSO.OpenTextFile(CurrentDirectory & "\all_quotas.txt", ForReading)
Set objFileOut = objFSO.OpenTextFile(CurrentDirectory & "\quotas_create.bat", ForWriting, True)
Set objFileOut1 = objFSO.OpenTextFile(CurrentDirectory & "\folders_copy.bat", ForWriting, True)
Set filesize = objFSO.GetFile(CurrentDirectory & "\quota_templates.xml")
objFileOut.Write("@echo off" & vbCrLf)
objFileOut.Write("chcp 1251" & vbCrLf)
objFileOut1.Write("@echo off" & vbCrLf)
objFileOut1.Write("chcp 1251" & vbCrLf)
Set objREx = CreateObject("VBScript.RegExp")
objREx.Global = True   
objREx.IgnoreCase = True
objREx.Pattern = ".*\\"
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True   
objRegEx.IgnoreCase = True
objRegEx.Pattern = "\((.*?)\)"

'Importing templates if they were exported previously
If filesize.size <> 0 Then
	objFileOut.Write("dirquota template import /file:" & CurrentDirectory & "\quota_templates.xml /Overwrite" & vbCrLf)
End If

Do Until objFileIn.AtEndOfStream
	ProcessQuota()
Loop

objFileOut.Close
objFileOut1.Close
objFileIn.Close

Sub ProcessQuota()
	
	quota_path = ""
	quota_status = ""
	quota_limit = ""
	quota_type = ""
	source_template = ""
	
	Do Until objFileIn.AtEndOfStream Or ((Len(quota_path) > 0) And (Len(quota_status) > 0) And (Len(quota_limit) > 0))
		tmpline = objFileIn.ReadLine
		If (Len(tmpline) > 0) Then
			If (InStr(tmpline, "Quota Path:") = 1) Then quota_path = LTrim(Right(tmpline, Len(tmpline) - Len("Quota Path:")))
			If (InStr(tmpline, "Source Template:") = 1) Then source_template = LTrim(Right(tmpline, Len(tmpline) - Len("Source Template:")))
			If (InStr(tmpline, "Quota Status:") = 1) Then quota_status = LTrim(Right(tmpline, Len(tmpline) - Len("Quota Status:")))
			If (InStr(tmpline, "Limit:") = 1) Then quota_limit = LTrim(Right(tmpline, Len(tmpline) - Len("Limit:")))
		End If
	Loop
	
	If quota_path <> "" Then
	
	'Removing "(Does not match template)" from template's name (if there is one)
	
	If (InStr(source_template, "(Does not match template)") <> 0) Then
		source_template = RTrim(Left(source_template, Len(source_template) - Len("(Does not match template)")))
	End If
	
	'Getting quota type from the limit string
	Set matches = objRegEx.Execute(quota_limit)
	count = matches.count
	For i = 0 To count - 1
		quota_type=matches(i).submatches(0)
		quota_limit = replace(left(quota_limit, Len(quota_limit) - Len(" (" & quota_type & ")"))," ","")
	Next
	
	'Converting units to a one grade lower to get rid of separator
	delimpos = 0
	If (UCase(Right(quota_limit, 2)) = "GB") Then
		quota_limit = CStr(CLng(CDbl(Left(quota_limit, Len(quota_limit) - 2)) * 1024)) & "MB"
	ElseIf (UCase(Right(quota_limit, 2)) = "MB") Then
		quota_limit = CStr(CLng(CDbl(Left(quota_limit, Len(quota_limit) - 2)) * 1024)) & "KB"
	ElseIf (UCase(Right(quota_limit, 2)) = "KB") Then
		quota_limit = CStr(CLng(Left(quota_limit, Len(quota_limit) - 2))) & "KB"
	End If
	
	If source_template = "None" Then
		objFileOut.Write("dirquota quota add /Overwrite /Path:""" & quota_path & """ /Limit:" & quota_limit & " /Type:" & quota_type & " /Status:" & quota_status & "" & vbCrLf)
	Else 'If quota was based off template then use this template again during quota creation
		objFileOut.Write("dirquota quota add /Overwrite /Path:""" & quota_path & """ /Limit:" & quota_limit & " /Type:" & quota_type & " /Status:" & quota_status & " /SourceTemplate:""" & source_template & """" & vbCrLf)
	End If
	
	'Generating folders_copy.bat
	objFileOut1.Write("xcopy /y /k /e /z """ & quota_path & "\*"" ""%1\" & objREx.Replace(quota_path,"") & "\*""" & vbCrLf)
	
	End If
End Sub
