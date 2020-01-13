
' ------------------------------------------------------------------
' (c) Copyright 2020 Tuomo Kulomaa <tuomo@smartpx.fi>
' ------------------------------------------------------------------
'
'                         SPX Videotool
'
' Video conversion tool for automatic processing of source folders.
' The process runs periodically and scans source folder(s) for video
' files(s) and processes those according to rules in config.ini.
'
' This application utilizes HTML Application capabilities of Windows
' operating system https://en.wikipedia.org/wiki/HTML_Application
' and ffmpeg (https://www.ffmpeg.org/) as the video processor. Other
' processors can be configured in task's configuration. For instance
' Imagemagick (http://imagemagick.org) or other commandline utilities.
'
' You can easily find ffmpeg commands using a search engine, try:
' https://www.google.com/search?q=ffmpeg+scale+video+example
'
' -------------------------------------------------------------------
'      ***     USE AT YOUR OWN RISK. SEE LICENSE BELOW.     ***
' -------------------------------------------------------------------
'
'	                    -- MIT LICENSE --
'	      Copyright 2020 Tuomo Kulomaa <tuomo@smartpx.fi>
'
'	Permission is hereby granted, free of charge, to any person
'	obtaining a copy of this software and associated documentation
'	files (the "Software"), to deal in the Software without
'	restriction, including without limitation the rights to use,
'	copy, modify, merge, publish, distribute, sublicense, and/or
'	sell copies of the Software, and to permit persons to whom
'	the Software is furnished to do so, subject to the following
'	conditions:
'
'	The above copyright notice and this permission notice shall
'	be included in all copies or substantial portions of the Software.
'
'	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY
'	KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
'	WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'	PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS
'	OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR
'	OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'	OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
'	SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'
' ---------------------------------------------------------------------------
'
'
' VBScript for SPX Videotool app using HTA UI.
' Functions sorted alphabetically.
' 
' Change history:
' 26.06.2019 TK		Original version
' 01.08.2019 TK		Refactored subfolder structure
' 31.10.2019 TK		Original version
' 31.10.2019 TK		Moved util functions to separate vbs, cleaned source code
' 27.11.2019 TK 	DEV (with hard coded paths) AVC intra 100
' 20.12.2019 TK 	Refacrored to the ini-file logic.
' 20.12.2019 TK		Consolidated all functions to this file from others.
' 13.01.2020 TK		Changed template strings to {{double mustache}}.


' GLOBAL VARIABLES
Dim pbTimerID
Dim pbHTML 
Dim pbHeight
Dim pbWidth
Dim pbBorder
Dim pbUnloadedColor
Dim pbLoadedColor
Dim pbStartTime
Dim SourceFolder1
Dim fileExtens_1
Dim fromFile
Dim DBLQUO
DBLQUO = Chr(34)
Dim CurFolder
Dim ProcCount
Dim FailCount
FailCount=0
ProcCount=0
Dim TargetFolder
Dim ProcessFoldr
Dim FailedFolder
Dim SOURCEFILE
Dim TARGETFILE
Dim EXTENSION
Dim ROOTFOLDER
Dim IniFile


Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set WshShell = CreateObject("WScript.Shell") 
strHtmlLocnVal = document.location.href 
strThisHTA = Replace(Right(strHtmlLocnVal, Len(strHtmlLocnVal) - 8), "/", "\") 
strThisHTA = UnEscape(strThisHTA) 
Set objThisFile = objFSO.GetFile(strThisHTA) 
objParentDir = objThisFile.ParentFolder 
Set objFolder = objFSO.GetFolder(objParentDir) 
CurFolder = objFolder.ShortPath
ROOTFOLDER = CurFolder
LoggingFolder = CurFolder & "\log\"
IniFile = CurFolder & "\config.ini"


' --------------------------------------------------------------------
' >>> Program starts when the page is loaded with Window_OnLoad() <<<<
' --------------------------------------------------------------------


Sub CancelAction
	On Error Resume Next
	Self.Close
End Sub

	

sub deleteIfExist(file)
	Set filesys = CreateObject("Scripting.FileSystemObject")
	If filesys.FileExists(file) Then 
		call SPXLog("Deleting existing file: '" & file & "'...")
		filesys.DeleteFile file
	end if
end sub



Sub DoActions
	x = addToConsole("",true)
	msg = "Completed " & ProcCount & ", failed " & FailCount & ". Status " & SPXTime("HHMMSS") & ":<BR>"
	msg = "Status " & SPXTime("HHMMSS") & "<BR>"
	x = addToConsole(msg,true)
	call ChangeText("completed",ProcCount)
	call ChangeText("failed",FailCount)
	ExecuteTasks(IniFile)
	pbStartTime = Now
	StartInterval
End Sub



sub ExecuteFFMPG(CONVERSIONTASK, SourceFolder, TargetFolder, FileExtensns)
	FilesForProcessingArr = Array() 'Empty array
	FilesForProcessingArr = getFilesForProcessingArr(SourceFolder, fileExtensns)
	msg = "<SPAN CLASS='spxBLUE'>" & CONVERSIONTASK & "</SPAN> -xfiles " & UBound(FilesForProcessingArr) & "/" & FilesForProcessingArr(UBound(FilesForProcessingArr)) & "."
	call addToConsole(msg,false)
	For x = 0 To Ubound(FilesForProcessingArr)-1
		file = FilesForProcessingArr(x)
		call SPX_CONVERSION(CONVERSIONTASK, file, TargetFolder)
	Next
end sub


sub ExecuteConversionTask(strSection, str_DESCRI, sourceFile, str_FILEXT, TargetFolder, str_EXECUT, str_PARAMS)
	call SPXLog("Executing Conversion Task - sourceFile: " & sourceFile & ", targetFolder: " & TargetFolder & ")")
	Set filesys = CreateObject("Scripting.FileSystemObject")
	If Not filesys.FileExists(sourceFile) Then 
		call SPXLog("FAILED: File [" & sourceFile & "] does not exist. Exiting routine.")
		FailCount = FailCount + 1
		call ChangeText("failed",FailCount)
		exit sub
	end if
	FilePath = filesys.GetParentFolderName(sourceFile)		' C:/folder/
	FileBase = filesys.GetBaseName(sourceFile)				' video
	FileName = filesys.GetFileName(sourceFile)				' video.avi
	FileExte = filesys.GetExtensionName(sourceFile)			' avi

	'Custom global vars for template mappings ============
	SOURCEFILE = sourceFile
	TARGETFILE = SuffixBackslash(TargetFolder) & FileBase
	EXTENSION = FileExte
	ROOTFOLDER = SuffixBackslash(CurFolder)
	'=====================================================

	ProcessFoldr = FilePath & "\PROCESSED\"
	FailedFolder = FilePath & "\FAILED\"
	call SPX_CreateFolder(ProcessFoldr)
	call SPX_CreateFolder(FailedFolder)
	call SPX_CreateFolder(PopulateTemplateString(TargetFolder))
	RenamedF = ProcessFoldr & FileName	
	FailedFi = FailedFolder & FileName

	CONVERSION_COMMAND = str_EXECUT & " " & str_PARAMS
	Log_File = GetLogfileRef()
	call addToConsole("- Executing task <SPAN CLASS='spxBLUE'>" & strSection & "</span><BR>- " & str_DESCRI, false) 
	SYSTEM_CMD= "cmd.exe /c echo CONVERTING " & SOURCEFILE & " & " & PopulateTemplateString(CONVERSION_COMMAND)  & " 2>> " & Log_File
	call SPXLog("Executing: " & SYSTEM_CMD)
	intReturn = 0

	' // Showtime!
	Set objShell = CreateObject("WScript.Shell")
	intReturn = objShell.Run(SYSTEM_CMD, 1, true)
	
	If intReturn <> 0 Then 
		call addToConsole("- <SPAN CLASS='spxERR'><B>ERROR</B></SPAN> " & intRetunr & " while processing " & SOURCEFILE & ". See log for details."  ,false)
		deleteIfExist(FailedFi)
		filesys.MoveFile SOURCEFILE, FailedFi
		call SPXLog("FAILED: Conversion retured error code: " & intReturn & ". File moved to " & FailedFi & ".")
		FailCount = FailCount + 1
	Else
		ProcCount = ProcCount + 1
		call addToConsole("- <B>DONE.</B><BR> ",false)
		call SPXLog("Moving file [" & SOURCEFILE & "] -> [" & RenamedF & "]")
		deleteIfExist(RenamedF)				
		filesys.MoveFile SOURCEFILE, RenamedF
		call SPXLog("COMPLETED: Original file moved to " & RenamedF & ".")
	End If		
	call ChangeText("completed",ProcCount)
	call ChangeText("failed",FailCount)
end sub


sub ExecuteIniSection(strSection, str_DESCRI, str_SOURCE, str_FILEXT, str_TARGET, str_EXECUT, str_PARAMS)
	FilesForProcessingArr = Array()
	FilesForProcessingArr = getFilesForProcessingArr(PopulateTemplateString(str_SOURCE), str_FILEXT)
	msg = "<SPAN CLASS='spxBLUE'>" & RightPad(strSection,20) & "</span> " & UBound(FilesForProcessingArr) & " hits / " & FilesForProcessingArr(UBound(FilesForProcessingArr)) & " files"
	call addToConsole(msg,false)
	For x = 0 To Ubound(FilesForProcessingArr)-1
		sourceFile = FilesForProcessingArr(x)
		call ExecuteConversionTask(strSection, str_DESCRI, sourceFile, str_FILEXT, str_TARGET, str_EXECUT, str_PARAMS)
	Next
end sub


Function ExecuteTasks( IniFile )
    Dim objFSO, objIniFile
    Dim strFilePath, strLine, strSection
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )
    strFilePath = Trim( IniFile )
    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, 1, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )
            If InStr(strLine,"[")>=0 AND InStr(strLine,"]")>0 And strLine <> "[CONFIG]" AND Right(strLIne,1)="]" Then
				strSection = strLine
				strSection = Replace(strSection,"[","")
				strSection = Replace(strSection,"]","")
				str_DESCRI = ReadIniValue(IniFile, strSection, "Task_Description")
				str_SOURCE = ReadIniValue(IniFile, strSection, "Source_Directory")
				str_FILEXT = ReadIniValue(IniFile, strSection, "Source_FileExtns")
				str_TARGET = ReadIniValue(IniFile, strSection, "Target_Directory")
				str_EXECUT = ReadIniValue(IniFile, strSection, "Converter_Progrm")
				str_PARAMS = ReadIniValue(IniFile, strSection, "Converter_Params")
				call ExecuteIniSection(strSection, str_DESCRI, str_SOURCE, str_FILEXT, str_TARGET, str_EXECUT, str_PARAMS)
            End If
        Loop
        objIniFile.Close
    Else
		msg = "ERROR! Config file '" & IniFile & "' doesn't exists. Cannot continue, please create the file. See Docs in the commands drop down and search for help."
		call addToConsole(msg,true)
        Exit function
    End If
End function


function FindString(LongTxtString,StartDelim,EndDelim)
	' will return part ("green") of string, such as "red green blue"
	' when usign "red" and "blue" as delims.
	FirstDelimPos = InStr(LongTxtString,StartDelim)+1+(Len(StartDelim)-1)
	RemainingStrg = Mid(LongTxtString,FirstDelimPos)
	SecondDelimPs = InStr(RemainingStrg,EndDelim)-1
	ResultingText = Mid(RemainingStrg,1,SecondDelimPs)
	FindString = ResultingText
end function


function GetLogfileRef()
	GetLogfileRef = LoggingFolder & "SPXLog_" & SPXTime("FILEDATE") & ".txt"
end function


function getFilesForProcessingArr(SourceFolder,fileExtensns)
	' Search for folder with desired fileformats and return filelist as array.
	' Note LAST index is for amount of all files in the source folder.
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(SourceFolder)
	Set colFiles = objFolder.Files
	fileExtnsArr = Split(fileExtensns, ",")
	fileList = Array()
	for each file in colFiles
		for each ext in fileExtnsArr
			if UCase(objFSO.getextensionname(file.path)) = UCase(Trim(ext)) then
				ReDim Preserve fileList(UBound(fileList) + 1)
				fileList(UBound(fileList)) = file.path			
			end if 
		next 
	next
	ReDim Preserve fileList(UBound(fileList) + 1)
	fileList(UBound(fileList)) = colFiles.Count
	getFilesForProcessingArr=fileList
end function 


Function LZ(ByVal Number, ByVal Places)
  Dim Zeros
  Zeros = String(CInt(Places), "0")
  LZ = Right(Zeros & CStr(Number), Places)
End Function


Function RightPad(str, len)
	Fill = "......................................"
	RightPad = Left(str & " " & Fill , len)
end function



function PopulateTemplateString(MyString)
	' (c) 2019 tuomo@smartpx.fi ---------------------------------------------------------------
	' Utility function to populate a template string with variable values. See this example:
	' 	BRAND = "AUDI"
	' 	MODEL = "eTron"
	' 	STR = "My current car is {{BRAND}} and it's model is {{MODEL}}."
	' 	Result- -> "My current car is AUDI and it's model is eTron."
	' --------------------------------------------
	StaDelim = "{{"
	EndDelim = "}}"
	do while InStr(MyString, StaDelim)>0
		FrstStaPos = InStr(MyString,StaDelim)-1
		FrstEndPos = InStr(MyString,EndDelim)+1+(Len(EndDelim)-1)
		FirstPart = Mid(MyString, 1, FrstStaPos)
		Remainder = Mid(MyString, FrstEndPos)
		MiddlPart = eval(FindString(MyString,StaDelim,EndDelim))
		MyString = FirstPart & MiddlPart & Remainder
	loop
	PopulateTemplateString = MyString
end function


Sub ProgressTick
	pbWaitTime=CInt(ReadIniValue(IniFile, "CONFIG", "Process_Interval"))
	BtnCaption=document.getElementById( "stateButton" ).value
	if BtnCaption = "►" then
		exit sub
	end if
	pbSecsPassed = DateDiff("s",pbStartTime,Now)
	pbMinsToGo =  Int((pbWaitTime - pbSecsPassed) / 60)
	pbSecsToGo = Int((pbWaitTime - pbSecsPassed) - (pbMinsToGo * 60))
	document.getElementById( "timeleft" ).innerText = pbSecsToGo
	if pbSecsToGo =< 0 then
		StopInterval
		DoActions
	end if
End Sub


Function ReadIniValue( myFilePath, mySection, myKey )
    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )
    ReadIniValue = ""
    strFilePath  = Trim( myFilePath )
    strSection   = Trim( mySection )
    strKey      = Trim( myKey )
    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, 1, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )
                Do While Left( strLine, 1 ) <> "["
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIniValue = Trim( Mid( strLine, intEqualPos + 1 ) )
                            If ReadIniValue = "" Then
                                ReadIniValue = " "
                            End If
                            Exit Do
                        End If
                    End If
                    If objIniFile.AtEndOfStream Then Exit Do
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
		x = msgbox("Config file '" & myFilePath & "' doesn't exists. Cannot continue, please re-install or create the file.", 16, "ERROR")
        Exit function
    End If
End Function



Sub SPXLog(TextToLog)
	Dim objFSO, myLogFile, LogStamp
	if (Right(LoggingFolder,1)<>"\") Then
		LogPath=LoggingFolder&"\"
	else
		LogPath=LoggingFolder
	end if
	call SPX_CreateFolder(LogPath)
	LogStamp=SPXTime("DDMMYY") & vbTab & SPXTime("HRMISEMS") & vbTab & TextToLog
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FolderExists(LogPath) Then
		Dim shl
		Set shl = CreateObject("WScript.Shell")
		Call shl.Run("%COMSPEC% /c mkdir " & LogPath,0,true)
	End If
	Const ForAppending = 8
	Set objTextFile = objFSO.OpenTextFile (GetLogfileRef(), ForAppending, True)
	objTextFile.WriteLine(LogStamp)
	objTextFile.Close
end sub


function SPX_CreateFolder(folderName)
	BCKSLH = Chr(92)
	Set objShell = CreateObject("Wscript.Shell")
	WaitCompletion = True
	HiddenWindow = 0
	objShell.Run "cmd /c mkdir " & Replace(folderName, "/", BCKSLH), HiddenWindow, WaitCompletion
end function


Sub SPXRun(File)
	Dim shell
	Set shell = CreateObject("WScript.Shell")
    shell.Run Chr(34) & File & Chr(34), 1, false
    Set shell = Nothing
End Sub

Function SPXTime(spxTimeFormat)
	Dim CurrTime, Elapsed, MilliSecs
	CurrTime = Now()
	Elapsed = Timer()
    YEA = LZ(Year(CurrTime),	4)
	MON = LZ(Month(CurrTime),	2)
	DEI = LZ(Day(CurrTime),		2)
	HRS = LZ(Hour(CurrTime),   	2)
	MIN = LZ(Minute(CurrTime), 	2)
	SEC = LZ(Second(CurrTime), 	2)
	MIS = LZ(Int((Elapsed - Int(Elapsed)) * 1000), 3)
	select case UCase(spxTimeFormat)
		case "HHMMSS"
			'SPXTime("HHMMSS") '18:35:55
			val = HRS & ":" & MIN & ":" & SEC 
		case "DDMMYY"
			'SPXTime("DDMMYY") '29.12.2019
			val = DEI & "." & MON & "." & YEA 
		case "FILEDATE"
			'SPXTime("DDMMYY") '29.12.2019
			val = DEI & "-" & MON & "-" & YEA 
		case "HRMISEMS"
			'SPXTime("HRMISEMS") '18:45:55.965
			val = HRS & ":" & MIN & ":" & SEC & "." & MIS 
		case "DATETIME"
			'SPXTime("DATETIME") '31.12.2019 18:45
			val = DEI & "." & MON & "." & YEA & " " & HRS & ":" & MIN & ":" & SEC  
		case else
			val = "Uknown date format"
	end select
	SPXTime = val
End Function


sub StartInterval
	pbTimerID = window.setInterval("ProgressTick", 200)
end sub


Sub StopInterval
	window.clearInterval(PBTimerID)
End Sub


function SuffixBackslash(pathAsString)
	select case Right(pathAsString,1)
		case "/"
			SuffixBackslash = Left(pathAsString, Len(pathAsString)-1) & "\"
		case "\"
			SuffixBackslash = pathAsString
		case else
			SuffixBackslash = pathAsString & "\"
	end select
end function


sub Talk(message)
	Dim sapi
	Set sapi=CreateObject("sapi.spvoice")
	sapi.Speak message
end sub


sub ToolSelected(optionValue)
	select case UCase(optionValue)
		case "DOC"
			call SPXRun("https://bitbucket.org/TuomoKu/spx-videotool")
		case "INI"
			call SPXRun(SuffixBackslash(ROOTFOLDER) & "config.ini")
		case "LOG"
			call SPXRun(GetLogfileRef)
		case "SPX"
			call SPXRun("http://www.smartpx.fi/")
		case else
			call addToConsole("<BR><SPAN CLASS='spxERR'>ERROR: Unknown menu call '" & optionValue & "'.</span>", false) 
	end select
end sub



Sub Window_OnLoad
	window.resizeTo 500,400		
	call SPXLog("--- Program started ---")
	pbStartTime = Now
	ProgressTick
	StartInterval
	DoActions 'remove this to avoid "scan on start"
	call ChangeText("started",SPXTime("DATETIME"))
End Sub


