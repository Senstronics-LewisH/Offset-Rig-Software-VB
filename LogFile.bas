Attribute VB_Name = "LogFile"
Option Explicit

#If VB7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Const CompileVersion As String = "v1.0.38"
Public UpdateNetworkPath As String
Public DevMode As Boolean

Public PSU_Visa_ID As String
Public Temp_Cal_Offset As Double
Public Relay_Delay As Long

Public BoardTypePath As String
Public UnionListPath As String
Public ConnectorTypePath As String
Public ColourListPath As String
Public CableListPath As String
Public CableUsagePath As String
Public RigTypePath As String
Public ResultsPath As String
Public LocalResultsPath As String
Public WorkOrderPath As String
Public LabelPrintInputPath As String
Public LabelPrintBatchPath As String
Public SoundCompletePath As String
Public SoundFailedPath As String


Public Type CurrentDrawOverride
    ProductRange As String
    LowerLimit As Double
    UpperLimit As Double
End Type

Public CurrentDrawOverrides() As CurrentDrawOverride
Public NumberOfCurrentDrawOverrides As Integer

Public UnionCode() As String
Public ProgramNumber() As Integer
Public OringCode() As String
Public OringColour() As String
Public NumberOfOrings As Integer
Public CableCode() As String
Public CableNumber() As String
Public CableID() As String
Public CableUsage() As Long
Public NumberOfCables As Integer
Public NumberOfUnions As Integer
Public NumberOfCableTypes As Integer
Public OffsetType As Integer
Public NumberOfConnectors As Integer
Public ConnectorCode() As String
Public ConnectorName() As String
Public IsFourPin() As Integer
Public BoardTypeList() As String
Public IsBoardSTC() As Integer
Public NumberOfBoardTypes As Integer
Public PinOut() As String
Public Sub AddToHistoryLogCDrive(ByVal LogEntry As String)

    Dim LogFile As Integer
    Dim PathName As String
    Dim FileLine As String
    
    On Error GoTo errhandler
    
    PathName = GetLogFileNameCDrive

    LogFile = FreeFile
    Open PathName For Append As #LogFile

    FileLine = Format$(Now, "ttttt") & vbTab & LogEntry
    
    Print #LogFile, FileLine
    
    Close #LogFile
    Exit Sub
    
errhandler:
    Exit Sub
End Sub
Public Function GetLogFileNameCDrive() As String
    
    Dim PathName As String
    Dim CurrentTimeAndDate As Date
    
    CurrentTimeAndDate = Now
    
    PathName = CreateCsfDirectoryStructureCDrive(CurrentTimeAndDate)
    
    GetLogFileNameCDrive = PathName & "\" & Format$(CurrentTimeAndDate, "yyyymmdd") & ".log"
    
End Function
Public Function CreateCsfDirectoryStructureCDrive(ByVal CalibrationStartTime As Date) As String
    Dim PathName As String
    Dim WeekNumber As String

    WeekNumber = Format$(CalibrationStartTime, "ww")
    If Len(WeekNumber) < 2 Then
        WeekNumber = "0" + WeekNumber
    End If
    
    On Error Resume Next
 
    PathName = LocalResultsPath
    MkDir PathName
   
    PathName = PathName & "\" & Format$(CalibrationStartTime, "yyyy")
    MkDir PathName
    
    PathName = PathName & "\Week" & WeekNumber
    MkDir PathName
    
    CreateCsfDirectoryStructureCDrive = PathName

End Function
Public Sub AddToHistoryLog(ByVal LogEntry As String)
    Dim LogFile As Integer
    Dim PathName As String
    Dim FileLine As String
    Dim Retries As Integer
    Dim Success As Boolean
    
    PathName = GetLogFileName
    FileLine = Format$(Now, "ttttt") & vbTab & LogEntry
    
    Retries = 0
    Success = False
    
    Do While Retries < 3 And Not Success
        On Error GoTo errhandler_append
        LogFile = FreeFile
        Open PathName For Append As #LogFile
        Print #LogFile, FileLine
        Close #LogFile
        Success = True
        On Error GoTo 0
    Loop
    
    If Not Success Then
        ' Fallback to local CDrive log
        AddToHistoryLogCDrive "[Network Log Failure] " & LogEntry
    End If
    Exit Sub

errhandler_append:
    Retries = Retries + 1
    Sleep 100
    Resume
End Sub
Public Function CreateCsfDirectoryStructure(ByVal CalibrationStartTime As Date) As String
    Dim PathName As String
    Dim WeekNumber As String

    WeekNumber = Format$(CalibrationStartTime, "ww")
    If Len(WeekNumber) < 2 Then
        WeekNumber = "0" + WeekNumber
    End If
    
    On Error Resume Next
 
    PathName = ResultsPath & "Switch Offset"
    MkDir PathName
   
    PathName = PathName & "\" & Format$(CalibrationStartTime, "yyyy")
    MkDir PathName
    
    PathName = PathName & "\Week" & WeekNumber
    MkDir PathName
    
    CreateCsfDirectoryStructure = PathName

End Function
Public Function GetLogFileName() As String
    
    Dim PathName As String
    Dim CurrentTimeAndDate As Date
    
    CurrentTimeAndDate = Now
    
    PathName = CreateCsfDirectoryStructure(CurrentTimeAndDate)
    
    GetLogFileName = PathName & "\" & Format$(CurrentTimeAndDate, "yyyymmdd") & ".log"
    
End Function
Public Function ReadBoardTypeList()

    Dim FileHandle As Integer
    Dim FileName As String
        
    On Error GoTo errhandler
    FileName = BoardTypePath
    FileHandle = FreeFile
    
    Open FileName For Input As #FileHandle

    Dim s As String
    Dim SplitValues() As String
    Dim i As Long
    i = 1
    Do While Not EOF(FileHandle)
        s = InputLine(FileHandle)
        s = Trim$(s)
        If s <> "" And Left$(s, 1) <> "#" And Left$(s, 1) <> ";" Then
            SplitValues = Split(s, ",")
            If UBound(SplitValues) >= 2 Then
                ReDim Preserve BoardTypeList(i)
                BoardTypeList(i) = Trim$(SplitValues(0))
                ReDim Preserve PinOut(i)
                PinOut(i) = Trim$(SplitValues(1))
                ReDim Preserve IsBoardSTC(i)
                IsBoardSTC(i) = Val(Trim$(SplitValues(2)))
                i = i + 1
            End If
        End If
    Loop
    NumberOfBoardTypes = i - 1
    
    Close #FileHandle
    
    Exit Function
    
errhandler:
    If Not DevMode Then MsgBox "Error opening Board Type list" & FileName & " Contact Engineering"
    ReadBoardTypeList = False

End Function
Public Function ReadRetrieveUnionList()
    
    Dim FileHandle As Integer
    Dim FileName As String
        
    On Error GoTo errhandler
    FileName = UnionListPath
    FileHandle = FreeFile
    
    Open FileName For Input As #FileHandle

    Dim s As String
    Dim SplitValues() As String
    Dim i As Long
    i = 1
    Do While Not EOF(FileHandle)
        s = InputLine(FileHandle)
        s = Trim$(s)
        If s <> "" And Left$(s, 1) <> "#" And Left$(s, 1) <> ";" Then
            SplitValues = Split(s, ",")
            If UBound(SplitValues) >= 1 Then
                ReDim Preserve UnionCode(i)
                UnionCode(i) = Trim$(SplitValues(0))
                ReDim Preserve ProgramNumber(i)
                ProgramNumber(i) = Val(Trim$(SplitValues(1)))
                i = i + 1
            End If
        End If
    Loop
    NumberOfUnions = i - 1
    
    Close #FileHandle
    
    Exit Function
    
errhandler:
    If Not DevMode Then MsgBox "Error opening union list" & FileName & " Contact Engineering"
    ReadRetrieveUnionList = False

End Function
Public Function ReadConnectorTypeList()
    
    Dim FileHandle As Integer
    Dim FileName As String
        
    On Error GoTo errhandler
    FileName = ConnectorTypePath
    FileHandle = FreeFile
    
    Open FileName For Input As #FileHandle

    Dim s As String
    Dim SplitValues() As String
    Dim i As Long
    i = 1
    Do While Not EOF(FileHandle)
        s = InputLine(FileHandle)
        s = Trim$(s)
        If s <> "" And Left$(s, 1) <> "#" And Left$(s, 1) <> ";" Then
            SplitValues = Split(s, ",")
            If UBound(SplitValues) >= 2 Then
                ReDim Preserve ConnectorCode(i)
                ConnectorCode(i) = Trim$(SplitValues(0))
                ReDim Preserve ConnectorName(i)
                ConnectorName(i) = Trim$(SplitValues(1))
                ReDim Preserve IsFourPin(i)
                IsFourPin(i) = Val(Trim$(SplitValues(2)))
                i = i + 1
            End If
        End If
    Loop
    NumberOfConnectors = i - 1
    
    Close #FileHandle
    
    Exit Function
    
errhandler:
    If Not DevMode Then MsgBox "Error opening connector code" & FileName & " Contact Engineering"
    ReadConnectorTypeList = False

End Function
Public Function ReadRetrieveColourList()
    
    Dim FileHandle As Integer
    Dim FileName As String
        
    On Error GoTo errhandler
    FileName = ColourListPath
    FileHandle = FreeFile
    
    Open FileName For Input As #FileHandle

    Dim s As String
    Dim SplitValues() As String
    Dim i As Long
    i = 1
    Do While Not EOF(FileHandle)
        s = InputLine(FileHandle)
        s = Trim$(s)
        If s <> "" And Left$(s, 1) <> "#" And Left$(s, 1) <> ";" Then
            SplitValues = Split(s, ",")
            If UBound(SplitValues) >= 1 Then
                ReDim Preserve OringCode(i)
                OringCode(i) = Trim$(SplitValues(0))
                ReDim Preserve OringColour(i)
                OringColour(i) = Trim$(SplitValues(1))
                i = i + 1
            End If
        End If
    Loop
    NumberOfOrings = i - 1
    
    Close #FileHandle
    
    Exit Function
    
errhandler:
    If Not DevMode Then MsgBox "Error opening O-ring Colour list" & FileName & " Contact Engineering"
    ReadRetrieveColourList = False

End Function
Public Function ReadRetrieveCableList()
    
    Dim FileHandle As Integer
    Dim FileName As String
        
    On Error GoTo errhandler
    FileName = CableListPath
    FileHandle = FreeFile
    
    Open FileName For Input As #FileHandle

    Dim s As String
    Dim SplitValues() As String
    Dim i As Long
        
    NumberOfCables = 0
    i = 1
    Do While Not EOF(FileHandle)
        s = InputLine(FileHandle)
        s = Trim$(s)
        If s <> "" And Left$(s, 1) <> "#" And Left$(s, 1) <> ";" Then
            SplitValues = Split(s, ",")
            If UBound(SplitValues) >= 1 Then
                ReDim Preserve CableCode(i)
                CableCode(i) = Trim$(SplitValues(0))
                ReDim Preserve CableNumber(i)
                CableNumber(i) = Trim$(SplitValues(1))
                i = i + 1
            End If
        End If
    Loop
    NumberOfCables = i - 1
    
    Close #FileHandle
    
    Exit Function
    
errhandler:
    If Not DevMode Then MsgBox "Error opening Cable list" & FileName & " Contact Engineering"
    ReadRetrieveCableList = False

End Function
Public Function ReadRetrieveCableUsage()
    
    Dim FileHandle As Integer
    Dim FileName As String
        
    On Error GoTo errhandler
    FileName = CableUsagePath
    FileHandle = FreeFile
    
    Open FileName For Input As #FileHandle

    Dim s As String
    Dim SplitValues() As String
    Dim i As Long
        
    NumberOfCableTypes = 0
    i = 1
    Do While Not EOF(FileHandle)
        s = InputLine(FileHandle)
        s = Trim$(s)
        If s <> "" And Left$(s, 1) <> "#" And Left$(s, 1) <> ";" Then
            SplitValues = Split(s, ",")
            If UBound(SplitValues) >= 1 Then
                ReDim Preserve CableID(i)
                CableID(i) = Trim$(SplitValues(0))
                ReDim Preserve CableUsage(i)
                CableUsage(i) = Val(Trim$(SplitValues(1)))
                i = i + 1
            End If
        End If
    Loop
    NumberOfCableTypes = i - 1
    
    Close #FileHandle
    
    Exit Function
    
errhandler:
    If Not DevMode Then MsgBox "Error opening Cable list" & FileName & " Contact Engineering"
    ReadRetrieveCableUsage = False

End Function
Public Function FindRigType()
    If DevMode Then
        OffsetType = 1
        FindRigType = True
        Exit Function
    End If
    
    Dim FileHandle As Integer
    Dim FileName As String
        
    On Error GoTo errhandler
    FileName = RigTypePath
    FileHandle = FreeFile
    
    Open FileName For Input As #FileHandle

    Dim s As String
    Dim SplitValues() As String
    Dim i As Long
     
    s = InputLine(FileHandle)

'1= Offset with switch and printer
'2= 25 Day Hold
'3= 90 Day Hold

        If s = "1" Then
            OffsetType = 1
        ElseIf s = "2" Then
            OffsetType = 2
        ElseIf s = "3" Then
            OffsetType = 3
        Else
            MsgBox "Do Not Recognose Offset - See Engineering"
        End If
                  
    Close #FileHandle
    
    Exit Function
    
errhandler:
    MsgBox "Error opening Rig Type" & FileName & " Contact Engineering"
    FindRigType = False

End Function
Public Sub FindWorkOrder()

  If MainForm.ScanAllCheck = 0 Then
        WorkOrderFound = False
        ReadLiveWorksOrderFile (WorkOrderPath)
    Else
        WorkOrderFound = True
    End If
    
    If WorkOrderFound = True Then
        
        If MainForm.ScanAllCheck = 0 Then
            MainForm.FirstMCSBarcode = "%0" & MainForm.FirstMCSBarcode
            MainForm.SecondMCSBarcode = "%1" & MainForm.SecondMCSBarcode
            MainForm.ThirdMCSBarcode = "%2" & MainForm.ThirdMCSBarcode
            MainForm.PartNumber = "%P" & MainForm.PartNumber 'LH 11/06/25 - Changes for Milwaukee
        End If
        
        
    End If

End Sub
Public Function InputLine(ByVal FileHandle As Integer) As String
    InputLine = vbNullString
    Dim NewCharacter As String
    
    While (EOF(FileHandle) = False) And (NewCharacter <> vbCr)
        NewCharacter = Input(1, FileHandle)
        If (NewCharacter <> vbCr) And (NewCharacter <> vbLf) And (NewCharacter <> vbTab) Then
            InputLine = InputLine & NewCharacter
        End If
    Wend
End Function
Public Function WriteCableUsageFile(ByVal FileName As String) As Boolean
    Dim FileLine As String
    Dim FileHandle As Integer
    Dim i As Integer
    
    On Error GoTo errhandler
    
    FileHandle = FreeFile
    Open FileName For Output As #FileHandle
    
For i = 1 To NumberOfCableTypes
    
    
    FileLine = CableID(i) & "," & CableUsage(i)
    Print #FileHandle, FileLine
    
Next

    Close #FileHandle

    Exit Function
    
errhandler:

End Function

Public Sub LoadCurrentDrawOverrides()
    Dim FileHandle As Integer
    Dim FileName As String
    Dim fileContent As String
    Dim lines() As String
    Dim SplitValues() As String
    Dim LineIdx As Long
    Dim Count As Long
    Dim s As String
    
    NumberOfCurrentDrawOverrides = 0
    ReDim CurrentDrawOverrides(0)
    
    FileName = App.Path & "\current_draw.txt"
    
    If Dir$(FileName) = "" Then
        Exit Sub
    End If
    
    On Error GoTo errhandler
    FileHandle = FreeFile
    Open FileName For Input As #FileHandle
    
    If LOF(FileHandle) > 0 Then
        fileContent = Input(LOF(FileHandle), FileHandle)
    Else
        fileContent = ""
    End If
    Close #FileHandle
    
    ' Normalize line endings to LF
    fileContent = Replace(fileContent, vbCrLf, vbLf)
    fileContent = Replace(fileContent, vbCr, vbLf)
    
    lines = Split(fileContent, vbLf)
    
    Count = 0
    For LineIdx = 0 To UBound(lines)
        s = Trim$(lines(LineIdx))
        If s <> "" And Left$(s, 1) <> "#" And Left$(s, 1) <> ";" Then
            SplitValues = Split(s, ",")
            If UBound(SplitValues) >= 2 Then
                ReDim Preserve CurrentDrawOverrides(Count)
                CurrentDrawOverrides(Count).ProductRange = Trim$(SplitValues(0))
                CurrentDrawOverrides(Count).LowerLimit = Val(Trim$(SplitValues(1)))
                CurrentDrawOverrides(Count).UpperLimit = Val(Trim$(SplitValues(2)))
                Count = Count + 1
            End If
        End If
    Next LineIdx
    
    NumberOfCurrentDrawOverrides = Count
    Exit Sub
    
errhandler:
    MsgBox "Error reading current_draw.txt: " & Err.Description & ". Using default limits."
End Sub

Public Sub GetCurrentLimits(ByVal ProductRange As String, ByRef OutLowerLimit As Double, ByRef OutUpperLimit As Double, Optional ByVal DefaultLower As Double = 2#, Optional ByVal DefaultUpper As Double = 5.6)
    Dim i As Long
    Dim SearchKey As String
    
    OutLowerLimit = DefaultLower
    OutUpperLimit = DefaultUpper
    
    SearchKey = Trim$(UCase$(ProductRange))
    For i = 0 To NumberOfCurrentDrawOverrides - 1
        If Trim$(UCase$(CurrentDrawOverrides(i).ProductRange)) = SearchKey Then
            OutLowerLimit = CurrentDrawOverrides(i).LowerLimit
            OutUpperLimit = CurrentDrawOverrides(i).UpperLimit
            Exit Sub
        End If
    Next i
End Sub

Public Sub LoadOffsetConfig()
    Dim FileHandle As Integer
    Dim FileName As String
    Dim fileContent As String
    Dim lines() As String
    Dim SplitValues() As String
    Dim LineIdx As Long
    Dim s As String
    Dim key As String
    Dim value As String

    ' Set default values
    PSU_Visa_ID = "9103875"
    Temp_Cal_Offset = -5#
    Relay_Delay = 0
    
    BoardTypePath = "C:\offset setup files\Board Type.txt"
    UnionListPath = "C:\offset setup files\union list.txt"
    ConnectorTypePath = "C:\offset setup files\connector type.txt"
    ColourListPath = "C:\offset setup files\Colour list.txt"
    CableListPath = "C:\offset setup files\cable list.txt"
    CableUsagePath = "C:\offset setup files\Cable Usage.txt"
    RigTypePath = "C:\offset setup files\Rig Type.txt"
    ResultsPath = "\\USVR8\Results\Production\Offset Check Results\"
    LocalResultsPath = "C:\My documents\Offset Check Results\"
    WorkOrderPath = "M:\paulseal.txt"
    LabelPrintInputPath = "M:\system\load\Vborders.txt"
    LabelPrintBatchPath = "C:\liveorders\PrintLabel.bat"
    SoundCompletePath = "C:\offset setup files\complete.wav"
    SoundFailedPath = "C:\offset setup files\failed.wav"
    UpdateNetworkPath = "Q:\SENSTRONICS\CONTROLLED MACHINE SOFTWARE\Offset Rig Software VB"
    DevMode = False

    FileName = App.Path & "\offset_config.txt"

    If Dir$(FileName) = "" Then
        Exit Sub
    End If

    On Error GoTo errhandler
    FileHandle = FreeFile
    Open FileName For Input As #FileHandle

    If LOF(FileHandle) > 0 Then
        fileContent = Input(LOF(FileHandle), FileHandle)
    Else
        fileContent = ""
    End If
    Close #FileHandle

    fileContent = Replace(fileContent, vbCrLf, vbLf)
    fileContent = Replace(fileContent, vbCr, vbLf)

    lines = Split(fileContent, vbLf)

    For LineIdx = 0 To UBound(lines)
        s = Trim$(lines(LineIdx))
        If s <> "" And Left$(s, 1) <> "#" And Left$(s, 1) <> ";" Then
            SplitValues = Split(s, ",")
            If UBound(SplitValues) >= 1 Then
                key = Trim$(LCase$(SplitValues(0)))
                value = Trim$(SplitValues(1))
                If key = "psu_visa_id" Then
                    PSU_Visa_ID = value
                ElseIf key = "temp_cal_offset" Then
                    Temp_Cal_Offset = Val(value)
                ElseIf key = "relay_delay" Then
                    Relay_Delay = Val(value)
                ElseIf key = "board_type_path" Then
                    BoardTypePath = value
                ElseIf key = "union_list_path" Then
                    UnionListPath = value
                ElseIf key = "connector_type_path" Then
                    ConnectorTypePath = value
                ElseIf key = "colour_list_path" Then
                    ColourListPath = value
                ElseIf key = "cable_list_path" Then
                    CableListPath = value
                ElseIf key = "cable_usage_path" Then
                    CableUsagePath = value
                ElseIf key = "rig_type_path" Then
                    RigTypePath = value
                ElseIf key = "results_path" Then
                    ResultsPath = value
                ElseIf key = "local_results_path" Then
                    LocalResultsPath = value
                ElseIf key = "work_order_path" Then
                    WorkOrderPath = value
                ElseIf key = "label_print_input_path" Then
                    LabelPrintInputPath = value
                ElseIf key = "label_print_batch_path" Then
                    LabelPrintBatchPath = value
                ElseIf key = "sound_complete_path" Then
                    SoundCompletePath = value
                ElseIf key = "sound_failed_path" Then
                    SoundFailedPath = value
                ElseIf key = "update_network_path" Then
                    UpdateNetworkPath = value
                ElseIf key = "dev_mode" Then
                    DevMode = (value = "1" Or LCase$(value) = "true")
                End If
            End If
        End If
    Next LineIdx
    Exit Sub

errhandler:
    MsgBox "Error reading offset_config.txt: " & Err.Description & ". Using default configuration."
End Sub

Public Sub CheckForUpdates()
    Dim NetworkExe As String
    Dim LocalExe As String
    Dim FSO As Object
    Dim UpdateExists As Boolean
    
    NetworkExe = UpdateNetworkPath & "\OffsetCheck.exe"
    LocalExe = App.Path & "\" & App.EXEName & ".exe"
    
    On Error GoTo errhandler
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If FSO.FileExists(NetworkExe) Then
        If FSO.GetFile(NetworkExe).DateLastModified > FSO.GetFile(LocalExe).DateLastModified Then
            UpdateExists = True
        End If
    End If
    
    If UpdateExists Then
        MsgBox "A newer version of the software is available on the network." & vbCrLf & _
               "The application will now close to install updates." & vbCrLf & _
               "Please relaunch using the desktop launcher icon.", vbInformation + vbOKOnly, "Update Available"
        Unload MainForm
        End
    End If
    Exit Sub
    
errhandler:
    ' Gracefully skip if network is disconnected
    Exit Sub
End Sub
