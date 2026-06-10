Attribute VB_Name = "OffsetCheck"
Option Explicit

Public ReadyFlag As Boolean
Public ReadyFlag2 As Boolean
Public PartNotCal As Boolean
Public OutputType As String
Public LimitPercent As Double
Public Limit As Double
Public Offset As Double
Public FullScale As Double
Public OutputSpan As Double
Public ZeroPressure As Double
Public ZeroPressureSign As String
Public FSPressure As Double
Public LoadValue As Double
Public LoadType As String
Public BarcodeUnits As String
Public Units As String
Public BarcodeSpan As Double
Public BarcodeOffset As Double
Public BarcodeFullScale As Double
Public BarcodeZeroPressure As Double
Public BarcodeFSPressure As Double
Public PressureSpan As Double
Public Vout2 As Boolean
Public Vout2STC As Boolean
Public LOWER As Double
Public UPPER As Double
Public Vout2Target As Double
Public VOUT2Spana As Double
Public VOUT2SpanB As Double
Public VOUT2Span As Double
Public CurrentTemp As Double
Public Span2 As Double
Public Fullscale2 As Double
Public Offset2 As Double
Public BoardType As String
Public CableType As String
Public CableType1 As String
Public CableUsageUpdate As Integer
Public STCTARGET As String
Public Result As Boolean
Public SupplyCurrent As Double
Public ConnectorType As String
Private Const Vout2errorlimit As Double = 2
Public SensorNumber As Double
Public WorksOrder As String
Public Vout2Only As Boolean
Public SwitchHyst As Boolean
Public CSVLine As Integer
Public CSVOffsetError As Double
Public CSVOffset As Double
Public OffsetError As Double
Public CSVLineRequest As Double
Public OffsetErrorPercent As Double
Public OringMCS As String
Public ORingRequired As Boolean
Public RestrictorRequired As Boolean
Public MCSUnionType As String
Public MCSORingType As String
Public VisionProgram As String
Public PackOnly As Boolean
Public WorkOrderFound As Boolean
Public NumberOfCycle As Integer
Public ProgRequired As Boolean
Public DetailChecked As Boolean
Public VoltageReading As Double
Public LogEntry As String
Public Error2 As Double
Public OffsetOnly As Boolean
Public Offset1Result As Boolean
Public Offset2Result As Boolean
Public CurrentResult As Boolean
Public VoltageResult As Boolean
Public CalResult As Boolean
Public SecondOffResult As Boolean
Public Post25DayTest As Boolean
Public Retests As Boolean
Public ORingResult As Boolean
Public RestrictorResult As Boolean
Public RestrictorWelded As Boolean
Public CorrectUnion As Boolean
Public STCGndResult As Boolean
Public STCVSResult As Boolean
Public STCV1Result As Boolean
Public STCV2Result As Boolean
Public PinOutSwitch As String
Public ScannedCableCode As String
Public IOCard As Boolean
Public Vout1Error As Double
Public TotalBad As Double
Public TotalGood As Double
Public SWPressure As Double
Public PSUSwitchTarget As Double
Public LowerSwitchTarget As Double
Public UpperSwitchTarget As Double
Public SwitchPSUValueError As Double
Public SWError As Double
Public UsePrinter As Boolean
Public Vsupply As Double
Public StartTime As Date
Public EndTime As Date
Public TestTime As Date
Public RecordTime As Date
Public CompletedTime As Date
Public TotalTime As Date

Public Enum SensorStatusEnum
    PASSED = 1
    FailedOffset1 = 2
    FailedOffToCal = 3
    FailedSecondOffset1 = 4
    FailedSTC = 5
    FailedCurrent = 6
    FailedSwitch = 7
    FailedTemp = 8
    FailedOring = 9
    FailedRestPresent = 10
    FailedRestweld = 11
    FailedWrongUnion = 12
    DidNotPassCalibration = 13
End Enum

Public SensorStatus(1 To 2000) As SensorStatusEnum  'increased from 800 by DW
Public Function ReadLiveWorksOrderFile(ByVal FileName As String)

    Dim FileHandle As Integer
    Dim s As String
        
    FileHandle = FreeFile
    Open FileName For Input As #FileHandle
    
    While EOF(FileHandle) = False
        s = WOInputLine(FileHandle)
        ProcessLiveWorkOrdersFileLine s
    Wend

End Function
Public Function WOInputLine(ByVal FileHandle As Integer) As String
    WOInputLine = vbNullString
    Dim NewCharacter As String
    
    While (EOF(FileHandle) = False) And (NewCharacter <> vbCr)
        NewCharacter = Input(1, FileHandle)
        If NewCharacter = vbLf Then
            Exit Function
        Else
            WOInputLine = WOInputLine & NewCharacter
        End If
    Wend
End Function
Private Function ProcessLiveWorkOrdersFileLine(ByVal s As String)

    Dim WO As String
    Dim BC1 As String
    Dim BC2 As String
    Dim BC3 As String
    Dim PN As String
    Dim WorkOrder As String
        
    If s <> "" Then
    
        SplitLiveWorkOrder s, WO, PN, BC1, BC2, BC3
        
        WorkOrder = (Mid$(MainForm.WorksOrderBarcode, 5, 15))
        
        If WorkOrder = WO Then
            WorkOrderFound = True
            MainForm.FirstMCSBarcode = BC1
            MainForm.SecondMCSBarcode = BC2
            MainForm.ThirdMCSBarcode = BC3
            
            'LH 11/06/25 - Changes for Milwaukee
            'Store Part Number
            MainForm.PartNumber = PN
        End If
    End If
End Function
Private Sub SplitLiveWorkOrder(ByVal s As String, ByRef WO As String, ByRef PN As String, ByRef BC1 As String, ByRef BC2 As String, ByRef BC3 As String)
     
   Dim SplitValues() As String
    
SplitValues = Split(s, ",")
        
WO = SplitValues(0)
PN = SplitValues(1)
BC1 = SplitValues(2)
BC2 = SplitValues(3)
BC3 = SplitValues(4)

End Sub
Public Function CheckDetails()

Dim i As Integer
Dim MyValue As String

    WorksOrder = (Mid$(MainForm.WorksOrderBarcode, 5, 15))
    
    MainForm.SensorID.Text = 1
    SensorNumber = 1
    
    If MainForm.WorksOrderBarcode = "" Then
        MsgBox "Works Order Not Entrered"
        Exit Function
    End If
    
    FindWorkOrder
    
    If WorkOrderFound = False Then
        MsgBox "NO WORK ORDER FOUND"
        MainForm.WorksOrderBarcode = ""
        Exit Function
    End If
    
    ReadConnectorTypeList
    ReadRetrieveColourList
    ReadRetrieveCableList
    ReadRetrieveUnionList
    ReadRetrieveCableUsage
    ReadBoardTypeList
    LoadOffsetConfig
    LoadCurrentDrawOverrides
    RestrictorRequired = False
   
    If Mid$(WorksOrder, 1, 2) = "UB" Or Mid$(WorksOrder, 1, 2) = "AM" Or Mid$(WorksOrder, 1, 2) = "ST" Or Mid$(WorksOrder, 1, 2) = "UC" Or Mid$(WorksOrder, 1, 2) = "Z0" Then
        PackOnly = True
    Else
        PackOnly = False
    End If
        
    If PackOnly = True Then
    
        MainForm.ClearDownForOffsetOnly
        MainForm.ClearDownForPackOnly
        MainForm.OffsetTargetDisplay = 0
        
        If Mid$(MainForm.WorksOrderBarcode, 5, 2) = "UB" Or Mid$(MainForm.WorksOrderBarcode, 5, 2) = "AM" Then
            MainForm.LowerLimitDisplay = -12
            MainForm.UpperLimitDisplay = 12
        Else
            MainForm.LowerLimitDisplay = -10
            MainForm.UpperLimitDisplay = 10
        End If
        
        MainForm.CurrentDisplay.Visible = True
        MainForm.CurrentLabel.Visible = True
        MainForm.InsulationDisplay.Visible = True
        MainForm.InsulationLabel.Visible = True
          
        RestrictorRequired = False
            
        MainForm.ConnectorTypeDisplay = "Pack"
        BoardType = "None"
    
        If (Mid$(MainForm.FirstMCSBarcode, 14, 1) = "/") Then
                
            If (Mid$(MainForm.FirstMCSBarcode, 15, 2) = "") Or (Mid$(MainForm.FirstMCSBarcode, 15, 1) = "-") Or (Mid$(MainForm.FirstMCSBarcode, 15, 2) = "--") Or (Mid$(MainForm.FirstMCSBarcode, 29, 2) = "TT") Then
                MainForm.ORingDisplay = "No"     ' Added 15,1 "-" to above line - Lee 24/04/20
                ORingRequired = False
            Else
                MainForm.ORingDisplay = "Yes"
                ORingRequired = True
            End If
        Else
            If (Mid$(MainForm.FirstMCSBarcode, 18, 2) = "") Or (Mid$(MainForm.FirstMCSBarcode, 18, 2) = "--") Then
                MainForm.ORingDisplay = "No"
                ORingRequired = False
            Else
                MainForm.ORingDisplay = "Yes"
                ORingRequired = True
            End If
        End If
        
        MCSUnionType = Mid$(MainForm.FirstMCSBarcode, 10, 3)
        MainForm.UnionCodeDisplay = MCSUnionType
        
        i = 1
    
        For i = 1 To NumberOfUnions + 1
        
            If i = NumberOfUnions Then
                MsgBox "CAN'T FIND UNION TYPE. CONTACT ENGINEERING"
                MainForm.ClearDown
                Exit Function
            Else
                If MCSUnionType = UnionCode(i) Then
                    VisionProgram = ProgramNumber(i)
                    MainForm.ProgramDisplay = VisionProgram
                    i = NumberOfUnions + 1
                End If
            End If
        Next
        
        MainForm.ProgramDisplay = VisionProgram
        ChangeVisionProgram VisionProgram
        
        If Mid$(WorksOrder, 1, 2) = "UC" Then
         
            If (Mid$(MainForm.FirstMCSBarcode, 13, 1) = "0") Then
                MainForm.RestrictorDisplay = "No"
                RestrictorRequired = False
            Else
                MainForm.RestrictorDisplay = "Yes"
                RestrictorRequired = True
            End If
             
        CableType = (Mid$(MainForm.FirstMCSBarcode, 15, 5))
        CableType1 = (Mid$(MainForm.FirstMCSBarcode, 15, 6))    ' Added Lee Scott 16/03/20
        If CableType1 = "AT2XTV" Then CableType = "AT2YT"       ' Added Lee
                  
        i = 1
                
        For i = 1 To NumberOfCables + 1
        
            If i = NumberOfCables + 1 Then
                MsgBox "Connector Code Not In Cable List. See Engineering"
                MainForm.ClearDown
                Exit Function
            End If
              
            If CableType = CableCode(i) Then
                
                MainForm.CableNumberDisplay = CableNumber(i)
            
                MyValue = InputBox("PLEASE SCAN CABLE " & CableNumber(i), "", "", 7000, 5000)
                                
                MainForm.CableNumberDisplay = MyValue
            
                      
                    If Len(MyValue) = 2 Then
                        ScannedCableCode = Left(MyValue, 2)
                    Else
                        ScannedCableCode = Left(MyValue, 1)
                    End If
                
                    If CableNumber(i) <> ScannedCableCode Then
                        MsgBox "CABLE DOES NOT MATCH. CHECK CABLE & TRY AGAIN"
                        MainForm.ClearDown
                        Exit Function
                    Else
                        i = NumberOfCables + 1
                    End If
                End If
            Next
        
            ConnectorType = (Mid$(MainForm.FirstMCSBarcode, 18, 2))
            
            i = 1
           
            For i = 1 To NumberOfConnectors + 1
            
                If i = NumberOfConnectors + 1 Then
                    MainForm.ConnectorTypeDisplay = "Unknown"
                    MsgBox "UNKNOWN CONNECTOR SEE ENGINEERING"
                    MainForm.ClearDown
                    Exit Function
                Else
        
                    If ConnectorType = ConnectorCode(i) Then
                        MainForm.ConnectorTypeDisplay = ConnectorName(i)
                        If OffsetOnly = False Then
                            If IsFourPin(i) = 1 Then
                                Vout2STC = True
                                MainForm.STCVOUT2LABEL.Visible = True
                                MainForm.STCVOUT2Display.Visible = True
                            Else
                                Vout2STC = False
                                MainForm.STCOUTPUTLabel.Visible = False
                                MainForm.STCTargetDisplay.Visible = False
                                MainForm.STCVOUT2LABEL.Visible = False
                                MainForm.STCVOUT2Display.Visible = False
                            End If
                        End If
                        i = NumberOfConnectors + 1
                    End If
                End If
            Next
            
          BoardType = (Mid$(MainForm.FirstMCSBarcode, 15, 2))
        
            If (Mid$(MainForm.FirstMCSBarcode, 20, 1)) = "M" Or (Mid$(MainForm.FirstMCSBarcode, 20, 1)) = "H" Then
                MsgBox " PLEASE STAMP FIT MATING CONNECTOR ON FRONT OF WORKS ORDER"
            End If
            
            
            i = 1
            
            For i = 1 To NumberOfBoardTypes + 1
        
                If i = NumberOfBoardTypes + 1 Then
                    MainForm.BoardTypeDisplay = "Unknown"
                    MsgBox "UNKNOWN BOARDTYPE SEE ENGINEERING"
                    MainForm.ClearDown
                    Exit Function
                Else
                    If BoardType = BoardTypeList(i) Then
                        MainForm.BoardTypeDisplay = BoardTypeList(i)
                        PinOutSwitch = PinOut(i)
                        If IsBoardSTC(i) = 0 Then
                          MainForm.STCTargetDisplay = "OVRFLW"
                        Else
                          MainForm.STCTargetDisplay = "Less Than 2 ohms"
                            If ConnectorType = "XY" Then
                                MainForm.STCTargetDisplay = "OVRFLW"
                            End If
                        End If
                        i = NumberOfBoardTypes + 1
                    End If
                End If
            
            Next
       End If
    End If
    
    If PackOnly = False Then
           
        CableType = (Mid$(MainForm.FirstMCSBarcode, 15, 5))
        CableType1 = (Mid$(MainForm.FirstMCSBarcode, 15, 6))    ' Added Lee Scott 16/03/20
        If CableType1 = "AT2XTV" Then CableType = "AT2YT"       ' Added Lee
                  
        i = 1
                
        For i = 1 To NumberOfCables + 1
        
            If i = NumberOfCables + 1 Then
                MsgBox "Connector Code Not In Cable List. See Engineering"
                MainForm.ClearDown
                Exit Function
            End If
              
            If CableType = CableCode(i) Then
                
                MainForm.CableNumberDisplay = CableNumber(i)
            
                MyValue = InputBox("PLEASE SCAN CABLE " & CableNumber(i), "", "", 7000, 5000)
                                                                
                MainForm.CableNumberDisplay = MyValue
                
                If Len(MyValue) = 2 Then
                    ScannedCableCode = Left(MyValue, 1)
                Else
                    ScannedCableCode = Left(MyValue, 2)
                End If
            
            ScannedCableCode = MyValue
            
                If CableNumber(i) <> ScannedCableCode Then
                    MsgBox "CABLE DOES NOT MATCH. CHECK CABLE & TRY AGAIN"
                    MainForm.ClearDown
                    Exit Function
                Else
                    i = NumberOfCables + 1
                End If
            End If
        Next
                 
        i = 1
        
        For i = 1 To NumberOfCableTypes + 1
        
            If MainForm.CableNumberDisplay = CableID(i) Then
                MainForm.NumberOfUsesDisplay = CableUsage(i)
                i = NumberOfCableTypes + 1
            End If
        Next
            
        MCSUnionType = Mid$(MainForm.FirstMCSBarcode, 10, 3)
        MainForm.UnionCodeDisplay = MCSUnionType
        
        i = 1
    
        For i = 1 To NumberOfUnions + 1
        
            If i = NumberOfUnions Then
                MsgBox "CAN'T FIND UNION TYPE. CONTACT ENGINEERING"
                MainForm.ClearDown
                Exit Function
            Else
                If MCSUnionType = UnionCode(i) Then
                    VisionProgram = ProgramNumber(i)
                    MainForm.ProgramDisplay = VisionProgram
                    i = NumberOfUnions + 1
                End If
            End If
        Next
        
        MainForm.ProgramDisplay = VisionProgram
        ChangeVisionProgram VisionProgram
                  
        ConnectorType = (Mid$(MainForm.FirstMCSBarcode, 18, 2))
        
        i = 1
       
        For i = 1 To NumberOfConnectors + 1
        
            If i = NumberOfConnectors + 1 Then
                MainForm.ConnectorTypeDisplay = "Unknown"
                MsgBox "UNKNOWN CONNECTOR SEE ENGINEERING"
                MainForm.ClearDown
                Exit Function
            Else
    
                If ConnectorType = ConnectorCode(i) Then
                    MainForm.ConnectorTypeDisplay = ConnectorName(i)
                    If OffsetOnly = False Then
                        If IsFourPin(i) = 1 Then
                            Vout2STC = True
                            MainForm.STCVOUT2LABEL.Visible = True
                            MainForm.STCVOUT2Display.Visible = True
                        Else
                            Vout2STC = False
                            MainForm.STCOUTPUTLabel.Visible = False
                            MainForm.STCTargetDisplay.Visible = False
                            MainForm.STCVOUT2LABEL.Visible = False
                            MainForm.STCVOUT2Display.Visible = False
                        End If
                    End If
                    i = NumberOfConnectors + 1
                End If
            End If
        
        Next
    
        If (Mid$(MainForm.ThirdMCSBarcode, 3, 1) = "T") Or (Mid$(MainForm.ThirdMCSBarcode, 3, 1) = "S") Then
            Vout2STC = False
            MainForm.STCOUTPUTLabel.Visible = False
            MainForm.STCTargetDisplay.Visible = False
            MainForm.STCVOUT2LABEL.Visible = False
            MainForm.STCVOUT2Display.Visible = False
            MainForm.VOUT2OUTPUTERRORDISPLAY.Visible = True
        End If
    
        If ConnectorType = "XP" Then
            MsgBox "USE GOLD CABLE"
        End If
             
        If ConnectorType = "XJ" Or ConnectorType = "XR" Then
        
            If Mid$(MainForm.FirstMCSBarcode, 10, 3) = "S05" Then
                Vout2STC = False ' common rail
                MainForm.STCOUTPUTLabel.Visible = False
                MainForm.STCTargetDisplay.Visible = False
                MainForm.STCVOUT2LABEL.Visible = False
                MainForm.STCVOUT2Display.Visible = False
            End If
        End If
        
        If ConnectorType = "XU" Then
           If (Mid$(MainForm.FirstMCSBarcode, 15, 2)) = "DT" Then
                MainForm.ConnectorTypeDisplay = "Large GDS NO STC"
                Vout2STC = False
                MainForm.STCOUTPUTLabel.Visible = False
                MainForm.STCTargetDisplay.Visible = False
                MainForm.STCVOUT2LABEL.Visible = False
                MainForm.STCVOUT2Display.Visible = False
            End If
        End If
    
                
        BoardType = (Mid$(MainForm.FirstMCSBarcode, 15, 2))
        
        If (Mid$(MainForm.FirstMCSBarcode, 20, 1)) = "M" Or (Mid$(MainForm.FirstMCSBarcode, 20, 1)) = "H" Then
            MsgBox " PLEASE STAMP FIT MATING CONNECTOR ON FRONT OF WORKS ORDER"
        End If
        
        
        i = 1
        
        For i = 1 To NumberOfBoardTypes + 1
    
            If i = NumberOfBoardTypes + 1 Then
                MainForm.BoardTypeDisplay = "Unknown"
                MsgBox "UNKNOWN BOARDTYPE SEE ENGINEERING"
                MainForm.ClearDown
                Exit Function
            Else
                If BoardType = BoardTypeList(i) Then
                    MainForm.BoardTypeDisplay = BoardTypeList(i)
                    PinOutSwitch = PinOut(i)
                    If IsBoardSTC(i) = 0 Then
                      MainForm.STCTargetDisplay = "OVRFLW"
                    Else
                      MainForm.STCTargetDisplay = "Less Than 2 ohms"
                        If ConnectorType = "XY" Then
                            MainForm.STCTargetDisplay = "OVRFLW"
                        End If
                    End If
                    i = NumberOfBoardTypes + 1
                End If
            End If
        
        Next
    
        If BoardType = "X5" Then
            Vout2 = False
            Vout2STC = False
            MainForm.VOUT2OUTPUTLABEL.Visible = False
            MainForm.VOUT2OutputDisplay.Visible = False
            MainForm.VOUT2OUTPUTERRORDISPLAY.Visible = False
            MainForm.VOUT2TargetDisplay.Visible = False
            MainForm.VOUT2TempDisplay.Visible = False
            MainForm.VOUT2TEMPLABEL.Visible = False
            MainForm.VOUT2ERRORLIMITLABEL.Visible = False
            MainForm.VOUT2LimitDisplay.Visible = False
            MainForm.VOUT2TARGETLABEL.Visible = False
            MainForm.STCVOUT2LABEL.Visible = False
            MainForm.STCVOUT2Display.Visible = False
        End If
        
        SwitchHyst = False
            
        If BoardType = "XD" Or BoardType = "XE" Then
            Vout2Only = True
            MainForm.VOUT2OUTPUTLABEL.Visible = True
            MainForm.VOUT2OutputDisplay.Visible = True
            MainForm.VOUT2OUTPUTERRORDISPLAY.Visible = True
            MainForm.VOUT1OUTPUTLABEL.Visible = False
            MainForm.VOUT1OutputDisplay.Visible = False
            MainForm.VOUT1OUTPUTERRORDISPLAY.Visible = False
            MainForm.STCVOUT1Label.Visible = False
            MainForm.STCVOUT1Display.Visible = False
            MainForm.STCVOUT2LABEL.Visible = True
            MainForm.STCVOUT2Display.Visible = True
        ElseIf BoardType = "XF" Or BoardType = "XG" Then
            SwitchHyst = True
            MainForm.VOUT2TEMPLABEL.Visible = False
            MainForm.VOUT2TempDisplay.Visible = False
            MainForm.VOUT1OUTPUTERRORDISPLAY.Visible = False
        ElseIf BoardType = "XH" Or BoardType = "XV" Then
            SwitchHyst = True
            MainForm.VOUT2TEMPLABEL.Visible = False
            MainForm.VOUT2TempDisplay.Visible = False
            MainForm.VOUT1OUTPUTERRORDISPLAY.Visible = False
        Else
            Vout2Only = False
            MainForm.VOUT1OUTPUTLABEL.Visible = True
            MainForm.VOUT1OutputDisplay.Visible = True
            MainForm.VOUT1OUTPUTERRORDISPLAY.Visible = True
            
            If PackOnly = True Or OffsetOnly = True Then
                MainForm.STCVOUT1Label.Visible = False
                MainForm.STCVOUT1Display.Visible = False
            Else
                MainForm.STCVOUT1Label.Visible = True
                MainForm.STCVOUT1Display.Visible = True
            End If
        End If

        If (Mid$(MainForm.FirstMCSBarcode, 22, 2)) = "SM" Then
             MsgBox "SMPE Paint Caps Required"
        End If

        If (Mid$(MainForm.FirstMCSBarcode, 18, 2)) = "AN" Then
             MsgBox "M12 Paint Caps Required"
        End If

        LoadValue = (Mid$(MainForm.FirstMCSBarcode, 25, 3))
        If LoadValue = "000" Then
        
            If CheckLoadOn = True Then
                MsgBox "REMOVE RESISTOR CONNECTED"
                MainForm.LoadValueDisplay.Visible = False
                MainForm.LoadTypeDisplay.Visible = False
                MainForm.LoadLabel.Visible = False
            End If
        Else
            MainForm.LoadTypeDisplay.Visible = True
            MainForm.LoadValueDisplay.Visible = True
            MainForm.LoadLabel.Visible = True
            CheckLoad
            MainForm.LoadValueDisplay = LoadValue
            MsgBox "PLACE RESISTOR " & LoadValue & LoadType
        End If
        If (Mid$(MainForm.ThirdMCSBarcode, 1, 3) = "%2T") Or (Mid$(MainForm.ThirdMCSBarcode, 1, 3) = "%2S") Then
            Vout2 = True
            MainForm.VOUT2OUTPUTLABEL.Visible = True
            MainForm.VOUT2OutputDisplay.Visible = True
            MainForm.VOUT2TargetDisplay.Visible = True
            MainForm.VOUT2TempDisplay.Visible = True
            MainForm.VOUT2TEMPLABEL.Visible = True
            MainForm.VOUT2LimitDisplay.Visible = True
            MainForm.VOUT2TARGETLABEL.Visible = True
            MainForm.VOUT2OUTPUTERRORDISPLAY.Visible = True
            MainForm.VOUT2ERRORLIMITLABEL.Visible = True
            
        Else
            Vout2 = False
            MainForm.VOUT2OUTPUTLABEL.Visible = False
            MainForm.VOUT2OutputDisplay.Visible = False
            MainForm.VOUT2OUTPUTERRORDISPLAY.Visible = False
            MainForm.VOUT2TargetDisplay.Visible = False
            MainForm.VOUT2TempDisplay.Visible = False
            MainForm.VOUT2TEMPLABEL.Visible = False
            MainForm.VOUT2ERRORLIMITLABEL.Visible = False
            MainForm.VOUT2LimitDisplay.Visible = False
            MainForm.VOUT2TARGETLABEL.Visible = False
        End If
    
        FINDLIMITS
        
    End If
    
    If OffsetOnly = True Or PackOnly = True Then
        MainForm.OFFSETFROMCALLabel.Visible = False
        MainForm.OffsetFromCalDisplay.Visible = False
        MainForm.OFFSETDIFFDISPLAY.Visible = False
    Else
        MainForm.OFFSETFROMCALLabel.Visible = True
        MainForm.OffsetFromCalDisplay.Visible = True
        MainForm.OFFSETDIFFDISPLAY.Visible = True
    End If
    
    DetailChecked = True
       
    If FindExcelFile = False Then
        If Post25DayTest = True Then
            MsgBox "CAN'T FIND ORIGINAL FILE. CONTACT ENGINEERING"
            MainForm.ClearDown
            Exit Function
        End If
        
        CreateExcel
    Else
        If FindPODInExcelFile = True Then
            MainForm.PODCheck = 1
        End If
        
        FindResults
        
    End If
       
End Function
Private Sub FINDLIMITS()
    Dim Vout1Reading As Double
    Dim Difference As Double

     OutputType = (Mid$(MainForm.SecondMCSBarcode, 3, 1))
    
     BarcodeUnits = (Mid$(MainForm.SecondMCSBarcode, 12, 1))
     BarcodeOffset = (Mid$(MainForm.SecondMCSBarcode, 4, 3))
     If (Mid$(MainForm.SecondMCSBarcode, 8, 1)) = "A" Then
     BarcodeFullScale = 1025
     Else
     BarcodeFullScale = (Mid$(MainForm.SecondMCSBarcode, 8, 3))
     End If
     
     If BarcodeUnits = "B" Then
     BarcodeZeroPressure = (Mid$(MainForm.SecondMCSBarcode, 13, 2))
     BarcodeFSPressure = (Mid$(MainForm.SecondMCSBarcode, 17, 4))
     Units = "bar"
     Else
     BarcodeZeroPressure = (Mid$(MainForm.SecondMCSBarcode, 12, 3))
     BarcodeFSPressure = (Mid$(MainForm.SecondMCSBarcode, 16, 5))
     Units = "psi"
     End If
    
     ZeroPressureSign = (Mid$(MainForm.SecondMCSBarcode, 11, 1))
     
     If OutputType = "A" Then
         Offset = BarcodeOffset / 10
         FullScale = BarcodeFullScale / 10
         Vsupply = 24
         MainForm.CurrentLabel.Visible = True
         MainForm.CurrentDisplay.Visible = True
         
     ElseIf OutputType = "B" Then
         Offset = BarcodeOffset / 100
         FullScale = BarcodeFullScale / 100
         Vsupply = 5
         MainForm.CurrentLabel.Visible = True
         MainForm.CurrentDisplay.Visible = True
         
     ElseIf OutputType = "C" Then
         Offset = BarcodeOffset / 10
         FullScale = BarcodeFullScale / 10
         Vsupply = 24
         
     ElseIf OutputType = "D" Then
         Offset = BarcodeOffset / 100
         FullScale = BarcodeFullScale / 100
         Vsupply = 24
         MainForm.CurrentLabel.Visible = True
         MainForm.CurrentDisplay.Visible = True
     ElseIf OutputType = "E" Then
         Offset = BarcodeOffset / 100
         FullScale = BarcodeFullScale / 100
         Vsupply = 5
         MainForm.CurrentLabel.Visible = True
         MainForm.CurrentDisplay.Visible = True
     End If
                   
                  
        If Vsupply = 5 Then
            SetPSU2 5
            PSUCheck
            Vout1Reading = MeasureDigitalMultimeterVolts
            Difference = 5 - Vout1Reading
            Vsupply = 5 + Difference
            SetPSU2 0
        End If
               
     OutputSpan = FullScale - Offset
         
     If ZeroPressureSign = "-" Then
     PressureSpan = BarcodeFSPressure + BarcodeZeroPressure
     Offset = (OutputSpan / PressureSpan * BarcodeZeroPressure) + Offset
     Else
     PressureSpan = BarcodeFSPressure - BarcodeZeroPressure
     End If
          
    If BarcodeOffset = 136 And BarcodeFullScale = 428 Then
        LOWER = 1.342
        UPPER = 1.378
        MainForm.UpperLimitDisplay = "1.378"
        MainForm.LowerLimitDisplay = "1.342"
        MainForm.OffsetTargetDisplay = Format$(Offset, "0.000")
       
    Else
    
        If Mid$(MainForm.FirstMCSBarcode, 22, 2) = "P2" Then
            Offset = 5.656
            MainForm.OffsetTargetDisplay = Offset
            OutputSpan = 16
        ElseIf Mid$(MainForm.FirstMCSBarcode, 22, 2) = "TG" Then 'added GC for Terex Genie
            Offset = 0.185
            MainForm.OffsetTargetDisplay = Offset
            OutputSpan = 6
        End If
    
    
    
        If Mid$(MainForm.FirstMCSBarcode, 22, 2) = "A1" Then
            MainForm.LimitPercentDisplay = 0.2
        ElseIf Mid$(MainForm.FirstMCSBarcode, 22, 2) = "X2" Then
            MainForm.LimitPercentDisplay = 0.15
        Else
            MainForm.LimitPercentDisplay = 0.5
        End If
        
        Limit = (MainForm.LimitPercentDisplay / 100) * OutputSpan
        LOWER = Offset - Limit
        UPPER = Offset + Limit
        
        If Offset = 0 Then
            LOWER = -0.025
        End If
        
        MainForm.LowerLimitDisplay = Format$(LOWER, "0.000")
        MainForm.UpperLimitDisplay = Format$(UPPER, "0.000")
        MainForm.OffsetTargetDisplay = Format$(Offset, "0.000")
        
    End If

    If (Mid$(MainForm.ThirdMCSBarcode, 3, 1) = "T") Then
    
        Fullscale2 = Mid$(MainForm.ThirdMCSBarcode, 8, 3)
        Offset2 = Mid$(MainForm.ThirdMCSBarcode, 4, 3)
    
        If OutputType = "A" Then
             Offset2 = Offset2 / 10
             Fullscale2 = Fullscale2 / 10
                      
        ElseIf OutputType = "B" Then
             Offset2 = Offset2 / 100
             Fullscale2 = Fullscale2 / 100
            
        ElseIf OutputType = "E" Then
             Offset2 = Offset2 / 100
             Fullscale2 = Fullscale2 / 100
            
        ElseIf OutputType = "D" Then
             Offset2 = Offset2 / 100
             Fullscale2 = Fullscale2 / 100
        End If
     
        Span2 = Fullscale2 - Offset2

'switch for temp meas
        
        SwitchTempMeas
        CurrentTemp = MeasureTemp
        CurrentTemp = CurrentTemp + Temp_Cal_Offset
        If CurrentTemp < 15 Or CurrentTemp > 35 Then
            MainForm.PASSED.Visible = False
            MainForm.FAILED.Visible = True
            MsgBox "Temperature Measurment Error - ABORT"
        Else
            VOUT2Spana = Mid$(MainForm.ThirdMCSBarcode, 11, 4)
            VOUT2SpanB = Mid$(MainForm.ThirdMCSBarcode, 17, 4)
            VOUT2Span = VOUT2SpanB - VOUT2Spana
            Vout2Target = (Span2 / VOUT2Span * (CurrentTemp - VOUT2Spana)) + Offset2
            MainForm.VOUT2TargetDisplay = Format$(Vout2Target, "0.0000")
            MainForm.VOUT2TempDisplay = Format$(CurrentTemp, "0.00" & " oC")
            MainForm.VOUT2LimitDisplay = Vout2errorlimit & " %"
        End If
    End If
    
    If (Mid$(MainForm.ThirdMCSBarcode, 3, 1) = "S") Then
        MainForm.VOUT2LimitDisplay.Visible = False
        MainForm.VOUT2ERRORLIMITLABEL.Visible = False
        
'check switch high or low

        If (Mid$(MainForm.ThirdMCSBarcode, 4, 1) = "0") Then
            MainForm.VOUT2TargetDisplay = "<0.2v"
        Else
            MainForm.VOUT2TargetDisplay = ">4v"
        End If
    End If

        If (Mid$(MainForm.FirstMCSBarcode, 13, 1) = "0") Then
            MainForm.RestrictorDisplay = "No"
            RestrictorRequired = False
        Else
            MainForm.RestrictorDisplay = "Yes"
            RestrictorRequired = True
        End If
        If (Mid$(MainForm.FirstMCSBarcode, 29, 2) = "--") Then
            MainForm.ORingDisplay = "No"
            ORingRequired = False
        Else
            MainForm.ORingDisplay = "Yes"
            ORingRequired = True
        End If
      
End Sub
Public Sub START()

StartTime = Format$(Now, "ttttt")
ReadyFlag2 = False
SensorStatus(MainForm.SensorID) = PASSED
MainForm.EnableStart.Visible = False

'AddToHistoryLog "Start = " & Now

'StartTime = Now

PartNotCal = False
MainForm.EmptyFields
DoEvents

If PackOnly = True Then
    PackTestOnly
ElseIf Vout2Only = True Then
' 3 pin deutsch with switch
    CHECKVOUT2ONLY
ElseIf OffsetOnly = True Then
    CHECKVOUT1
Else
    
    CHECKSTC
    CHECKVOUT1
    CHECKVOUT2
   
    If Post25DayTest = False And Retests = False And SwitchHyst = False And BoardType <> "HY" Then
    
        FindOffset
    
        If PartNotCal = True Then
            MainForm.EmptyFields
            DoEvents
            PASS_FAIL
            Exit Sub
        End If
        
        If SwitchHyst <> True Then
            
            OffsetError = Format$((MainForm.VOUT1OutputDisplay - CSVOffset), "0.0000")
            
            OffsetErrorPercent = Format$(OffsetError / OutputSpan * 100, "0.0000")
           If MainForm.BoardTypeDisplay = "RT" Then
            MainForm.OffsetFromCalDisplay = 0.185   'added GC for Terex Genie
            OffsetError = Format$((MainForm.VOUT1OutputDisplay - 0.185), "0.0000")
            OffsetErrorPercent = Format$(OffsetError / OutputSpan * 100, "0.0000")
                        
           Else
            MainForm.OffsetFromCalDisplay = CSVOffset
           End If
            MainForm.OFFSETDIFFDISPLAY = OffsetErrorPercent
        End If

        If Abs(OffsetErrorPercent) > 0.4 Then 'Limits changed on 30/03/2022 ref ECR22-021
        'If OffsetErrorPercent < -0.2 Or OffsetErrorPercent > 0.3 Then
            MainForm.OFFSETFROMCALPASS.Visible = False
            MainForm.OFFSETFROMCALFAIL.Visible = True
            SensorStatus(MainForm.SensorID) = FailedOffToCal
        Else
            MainForm.OFFSETFROMCALPASS.Visible = True
            MainForm.OFFSETFROMCALFAIL.Visible = False
        End If
    End If

End If

    If SensorStatus(MainForm.SensorID) = PASSED Then
        RecieveVision
    End If

DoEvents

PASS_FAIL

End Sub
Private Sub PackTestOnly()
    
Dim Vout1Reading As Double
Dim VSSTCReading As Double

    SetPSU2 10
    SwitchPackOffsetMeas
    MainForm.CurrentDisplay.Visible = True
    MainForm.CurrentLabel.Visible = True
    MainForm.InsulationDisplay.Visible = True
    MainForm.InsulationLabel.Visible = True
    
    Vout1Reading = MeasureDigitalMultimeterVolts
    Vout1Reading = Vout1Reading * 1000
    MainForm.VOUT1OutputDisplay = Format$(Vout1Reading, "0.000 " & "mv")
    
 
    If Mid$(MainForm.WorksOrderBarcode, 5, 2) = "UB" Or Mid$(MainForm.WorksOrderBarcode, 5, 2) = "AM" Then
        
        If Abs(Vout1Reading) < 12 Then
            MainForm.VOUT1PASS.Visible = True
        Else
            MainForm.VOUT1FAIL.Visible = True
            SensorStatus(MainForm.SensorID) = FailedOffset1
        End If
    Else
    
        If Abs(Vout1Reading) < 10 Then
            MainForm.VOUT1PASS.Visible = True
        Else
            MainForm.VOUT1FAIL.Visible = True
            SensorStatus(MainForm.SensorID) = FailedOffset1
        End If
    
    End If
   
' read current

    VerifySupplyCurrent 1#, 999999#

' check insulation

    SwitchPackInsulation
    'dummy reading
    VSSTCReading = MeasureDigitalMultimeterOhms
    Sleep 100
'take reading
    VSSTCReading = MeasureDigitalMultimeterOhms
        
    If VSSTCReading > 100000000 Then
        MainForm.STCVSDisplay = Format$(VSSTCReading, "0.00" & " Ohms")
        MainForm.InsulationDisplay = "OVRFLW"
        MainForm.InsulationPass.Visible = True
        MainForm.STCVSPASS.Visible = True
        MainForm.STCVSFAIL.Visible = False
        STCVSResult = True
    Else
        MainForm.STCVSDisplay = Format$(VSSTCReading, "0.00" & " Ohms")
        MainForm.InsulationDisplay = MainForm.STCVSDisplay
        MainForm.STCVSPASS.Visible = False
        MainForm.STCVSFAIL.Visible = True
        MainForm.InsulationFail.Visible = True
        SensorStatus(MainForm.SensorID) = FailedSTC
        STCVSResult = False
    End If
    
End Sub
Private Sub CHECKSTC()

    Dim VSSTCReading As Double
    Dim GNDSTCReading As Double
    Dim VOUT1STCReading As Double
    Dim VOUT2STCReading As Double
    
'connect VS to stc
    
    If PinOutSwitch = "T" Then
        SwitchSTCVsMeas
    ElseIf PinOutSwitch = "Q" Then
        SwitchSTCVsMeasQ
    ElseIf PinOutSwitch = "M" Then
        SwitchSTCVsMeasM
    ElseIf PinOutSwitch = "Z" Then
        SwitchSTCVsMeasZ
    ElseIf PinOutSwitch = "A" Then
        SwitchSTCVsMeasA
    ElseIf PinOutSwitch = "W" Then
        SwitchSTCVsMeasW
    ElseIf PinOutSwitch = "L" Then
        SwitchSTCVsMeasL
    ElseIf PinOutSwitch = "N" Then
        SwitchSTCVsMeasN
    ElseIf PinOutSwitch = "F" Then
        SwitchSTCVsMeasF
        If ConnectorType = "C1" Or ConnectorType = "C2" Then 'sumitomo
            SwitchSTCVsMeas
        End If
    ElseIf PinOutSwitch = "C" Then
        SwitchSTCVsMeasC
    ElseIf PinOutSwitch = "6" Then
        SwitchSTCVsMeas6
    ElseIf PinOutSwitch = "U" Then 'ADDED BY DW 17/11/23
        SwitchSTCVsMeasU
    Else
        MsgBox "UNKNOWN BOARDTYPE CONTACT ENGINEERING"
        MainForm.ClearDown
        Exit Sub
    End If
    
          
'dummy reading
    VSSTCReading = MeasureDigitalMultimeterOhms
    Sleep 100
'take reading
    VSSTCReading = MeasureDigitalMultimeterOhms

    If VSSTCReading > 40000000 Then
        MainForm.STCVSDisplay = "OVRFLW"
    Else
        MainForm.STCVSDisplay = Format$(VSSTCReading, "0.00" & " Ohms")
    End If
    
    If VSSTCReading > 40000000 Then
        MainForm.STCVSPASS.Visible = True
        MainForm.STCVSFAIL.Visible = False
        STCVSResult = True
    Else
        MainForm.STCVSPASS.Visible = False
        MainForm.STCVSFAIL.Visible = True
        SensorStatus(MainForm.SensorID) = FailedSTC
        STCVSResult = False
    End If

DoEvents

'connect GND to stc

    If PinOutSwitch = "T" Then
        SwitchSTCGndMeas
    ElseIf PinOutSwitch = "Q" Then
        SwitchSTCGndMeasQ
    ElseIf PinOutSwitch = "M" Then
        SwitchSTCGndMeasM
    ElseIf PinOutSwitch = "Z" Then
        SwitchSTCGndMeasZ
    ElseIf PinOutSwitch = "A" Then
        SwitchSTCGndMeasA
    ElseIf PinOutSwitch = "W" Then
        SwitchSTCGndMeasW
    ElseIf PinOutSwitch = "X" Then
        SwitchSTCGndMeasL
    ElseIf PinOutSwitch = "N" Then
        SwitchSTCGndMeasN
    ElseIf PinOutSwitch = "F" Then
        SwitchSTCGndMeasF
        If ConnectorType = "C1" Or ConnectorType = "C2" Then 'sumitomo
            SwitchSTCGndMeas
        End If
    ElseIf PinOutSwitch = "C" Then
        SwitchSTCGndMeasC
    ElseIf PinOutSwitch = "6" Then
        SwitchSTCGndMeas6
    ElseIf PinOutSwitch = "U" Then 'ADDED BY DW 17/11/23
        SwitchSTCGndMeasU
    End If

    If BoardType = "HY" Then
    Sleep 300
    End If

'take reading
    GNDSTCReading = MeasureDigitalMultimeterOhms

    If GNDSTCReading > 40000000 Then
        MainForm.STCGNDDisplay = "OVRFLW"
    Else
        MainForm.STCGNDDisplay = Format$(GNDSTCReading, "0.00" & " Ohms")
    End If

    If GNDSTCReading > 40000000 Then
        MainForm.STCGNDPASS.Visible = True
        MainForm.STCGNDFAIL.Visible = False
        STCGndResult = True
    Else
        MainForm.STCGNDPASS.Visible = False
        MainForm.STCGNDFAIL.Visible = True
        SensorStatus(MainForm.SensorID) = FailedSTC
        STCGndResult = False
    End If

DoEvents
'connect Vout1 to stc

    If PinOutSwitch = "T" Then
        SwitchSTCVout1Meas
    ElseIf PinOutSwitch = "Q" Then
        SwitchSTCVout1MeasQ
    ElseIf PinOutSwitch = "M" Then
        SwitchSTCVout1MeasM
    ElseIf PinOutSwitch = "Z" Then
        SwitchSTCVout1MeasZ
    ElseIf PinOutSwitch = "A" Then
        SwitchSTCVout1MeasA
    ElseIf PinOutSwitch = "W" Then
        SwitchSTCVout1MeasW
    ElseIf PinOutSwitch = "L" Then
        SwitchSTCVout1MeasL
    ElseIf PinOutSwitch = "N" Then
        SwitchSTCVout1MeasN
    ElseIf PinOutSwitch = "F" Then
        SwitchSTCVout1MeasF
        If ConnectorType = "C1" Or ConnectorType = "C2" Then 'sumitomo
            SwitchSTCVout1Meas
        End If
    ElseIf PinOutSwitch = "C" Then
        SwitchSTCVout1MeasC
    ElseIf PinOutSwitch = "6" Then
        SwitchSTCVout1Meas6
    ElseIf PinOutSwitch = "U" Then 'ADDED BY DW 17/11/23
        SwitchSTCVout1MeasU
    End If
    
'take reading
    VOUT1STCReading = MeasureDigitalMultimeterOhms

    If VOUT1STCReading > 40000000 Then
        MainForm.STCVOUT1Display = "OVRFLW"
    Else
        MainForm.STCVOUT1Display = Format$(VOUT1STCReading, "0.00" & " Ohms")
    End If

    If VOUT1STCReading > 40000000 Then
        MainForm.STCVOUT1PASS.Visible = True
        MainForm.STCVOUT1FAIL.Visible = False
        STCV1Result = True
    Else
        MainForm.STCVOUT1PASS.Visible = False
        MainForm.STCVOUT1FAIL.Visible = True
        SensorStatus(MainForm.SensorID) = FailedSTC
        STCV1Result = False
    End If

DoEvents


If Vout2STC = True Then

'connect Vout2 to stc

    If PinOutSwitch = "T" Then
        SwitchSTCVout2Meas
    ElseIf PinOutSwitch = "Q" Then
        SwitchSTCVout2MeasQ
    ElseIf PinOutSwitch = "M" Then
        SwitchSTCVout2MeasM
    ElseIf PinOutSwitch = "Z" Then
        SwitchSTCVout2MeasZ
    ElseIf PinOutSwitch = "W" Then
        SwitchSTCVout2MeasW
    ElseIf PinOutSwitch = "A" Then
        SwitchSTCVout2MeasA
    ElseIf PinOutSwitch = "L" Then
        SwitchSTCVout2MeasL
    ElseIf PinOutSwitch = "N" Then
        SwitchSTCVout2MeasN
    ElseIf PinOutSwitch = "F" Then
        SwitchSTCVout2MeasF
        If ConnectorType = "C1" Or ConnectorType = "C2" Then 'sumitomo
            SwitchSTCVout2Meas
        End If
    ElseIf PinOutSwitch = "C" Then
        SwitchSTCVout2MeasC
    ElseIf PinOutSwitch = "6" Then
        SwitchSTCVout2Meas6
    End If
    
'take reading
    VOUT2STCReading = MeasureDigitalMultimeterOhms

    If VOUT2STCReading > 40000000 Then
    MainForm.STCVOUT2Display = "OVRFLW"
    Else
    VOUT2STCReading = VOUT2STCReading - 4.5
    If VOUT2STCReading < 0 Then VOUT2STCReading = 0
    MainForm.STCVOUT2Display = Format$(VOUT2STCReading, "0.00" & " Ohms")
    
    End If

    If MainForm.STCTargetDisplay = "OVRFLW" Then
        If VOUT2STCReading > 40000000 Then
            MainForm.STCVOUT2PASS.Visible = True
            MainForm.STCVOUT2FAIL.Visible = False
            STCV2Result = True
        Else
            MainForm.STCVOUT2PASS.Visible = False
            MainForm.STCVOUT2FAIL.Visible = True
            SensorStatus(MainForm.SensorID) = FailedSTC
            STCV2Result = False
        End If
    Else
        If VOUT2STCReading <= 2 Then
            MainForm.STCVOUT2PASS.Visible = True
            MainForm.STCVOUT2FAIL.Visible = False
            STCV2Result = True
        Else
            MainForm.STCVOUT2PASS.Visible = False
            MainForm.STCVOUT2FAIL.Visible = True
            SensorStatus(MainForm.SensorID) = FailedSTC
            STCV2Result = False
        End If
    End If
Else

End If
    DoEvents

End Sub
Private Sub CHECKVOUT1()


Dim Vout1Reading As Double
Dim Vout2Reading As Double
       


'dual switch

    If BoardType = "HY" Then
        
        MainForm.OFFSETFROMCALLabel.Visible = False
        MainForm.OffsetFromCalDisplay.Visible = False
        MainForm.OFFSETDIFFDISPLAY.Visible = False
        MainForm.VoltageLabel.Visible = False
        MainForm.VOUT1OUTPUTERRORDISPLAY.Visible = False
        SwitchDualVout1Meas
        SetPSU2 Vsupply
        Vout1Reading = MeasureDigitalMultimeterVolts
        MainForm.VOUT1OutputDisplay = Format$(Vout1Reading, "0.0000")
           
                
        If Vout1Reading < 1 Then
            MainForm.VOUT1PASS.Visible = True
            MainForm.VOUT1FAIL.Visible = False
        Else
            MainForm.VOUT1PASS.Visible = False
            MainForm.VOUT1FAIL.Visible = True
            SensorStatus(MainForm.SensorID) = FailedSwitch
        End If
    
        SwitchDualVout2Meas
        
        Vout2Reading = MeasureDigitalMultimeterVolts
        MainForm.VOUT2OutputDisplay = Format$(Vout2Reading, "0.0000")
            
        If (Mid$(MainForm.ThirdMCSBarcode, 4, 1) = "0") Then
            If Vout2Reading < 1 Then
                MainForm.VOUT2PASS.Visible = True
                MainForm.VOUT2FAIL.Visible = False
            Else
                MainForm.VOUT2PASS.Visible = False
                MainForm.VOUT2FAIL.Visible = True
                SensorStatus(MainForm.SensorID) = FailedSwitch
            End If
'check output is less than 0.2v
        Else
            If Vout2Reading > 4 Then
                MainForm.VOUT2PASS.Visible = True
                MainForm.VOUT2FAIL.Visible = False
            Else
                MainForm.VOUT2PASS.Visible = False
                MainForm.VOUT2FAIL.Visible = True
                SensorStatus(MainForm.SensorID) = FailedSwitch
            End If
        End If
        
        VerifySupplyCurrent
                 
        DoEvents

        Exit Sub
    End If
        
    If SwitchHyst = True Then
    
        SwitchVout2MeasA
        SetPSU2 Vsupply
        Vout1Reading = MeasureDigitalMultimeterVolts
        MainForm.VOUT1OutputDisplay = Format$(Vout1Reading, "0.0000")
    
        If Vout1Reading > 0 And Vout1Reading < 3.8 Then
            MainForm.VOUT1PASS.Visible = True
            MainForm.VOUT1FAIL.Visible = False
        Else
            MainForm.VOUT1PASS.Visible = False
            MainForm.VOUT1FAIL.Visible = True
            SensorStatus(MainForm.SensorID) = FailedSwitch
        End If
    
        VerifySupplyCurrent
        
        DoEvents
        
        Exit Sub
    End If

    If OutputType = "C" Then
    
        If PinOutSwitch = "T" Then
            SwitchCurrentMeas
        ElseIf PinOutSwitch = "Q" Then
            SwitchCurrentMeasQ
        ElseIf PinOutSwitch = "M" Then
            SwitchCurrentMeasM
        ElseIf PinOutSwitch = "Z" Then
            SwitchCurrentMeasZ
        ElseIf PinOutSwitch = "A" Then
            SwitchCurrentMeasA
        ElseIf PinOutSwitch = "W" Then
            SwitchCurrentMeasW
        ElseIf PinOutSwitch = "L" Then
            SwitchCurrentMeasL
        ElseIf PinOutSwitch = "N" Then
            SwitchCurrentMeasN
        ElseIf PinOutSwitch = "F" Then
            SwitchCurrentMeasF
            If ConnectorType = "C1" Or ConnectorType = "C2" Then 'sumitomo
                SwitchCurrentMeas
            End If
        ElseIf PinOutSwitch = "C" Then
            SwitchCurrentMeasC
        ElseIf PinOutSwitch = "6" Then
            SwitchCurrentMeas6
        ElseIf PinOutSwitch = "U" Then 'ADDED BY DW 17/11/23
            SwitchCurrentMeasU
        End If

        If BarcodeUnits = "B" Then
            FSPressure = (Mid$(MainForm.SecondMCSBarcode, 17, 4))
        Else
            FSPressure = (Mid$(MainForm.SecondMCSBarcode, 16, 5) / 14.5)
        End If
        

            SetPSU2 Vsupply
            
            If MainForm.PODCheck = 1 Then
                Sleep 5000
            End If
        
            Vout1Reading = MeasureDigitalMultimeterAmps
            Vout1Reading = Vout1Reading * 1000

        If PinOutSwitch = "T" Then
            SwitchVout1Meas
        ElseIf PinOutSwitch = "Q" Then
            SwitchVout1MeasQ
        ElseIf PinOutSwitch = "M" Then
            SwitchVout1MeasM
        ElseIf PinOutSwitch = "Z" Then
            SwitchVout1MeasZ
        ElseIf PinOutSwitch = "A" Then
            SwitchVout1MeasA
        ElseIf PinOutSwitch = "W" Then
            SwitchVout1MeasW
        ElseIf PinOutSwitch = "L" Then
            SwitchVout1MeasL
        ElseIf PinOutSwitch = "N" Then
            SwitchVout1MeasN
        ElseIf PinOutSwitch = "F" Then
            SwitchVout1MeasF
            If ConnectorType = "C1" Or ConnectorType = "C2" Then 'sumitomo
                SwitchVout1Meas
            End If
        ElseIf PinOutSwitch = "C" Then
            SwitchVout1MeasC
        ElseIf PinOutSwitch = "6" Then
            SwitchVout1Meas6
        End If
    'ignore voltage output
        If MCSUnionType = "S05" Or Mid$(BoardType, 1, 1) = "L" Or ConnectorType = "H1" Then
            MainForm.CurrentDisplay.Visible = False
            MainForm.VoltageLabel.Visible = False
        Else
            MainForm.VoltageLabel.Visible = True
            MainForm.CurrentDisplay.Visible = True
            SetPSU2 Vsupply
            VoltageReading = MeasureDigitalMultimeterVolts
            MainForm.CurrentDisplay = Format$(VoltageReading, "0.00" & " V")
            
            If MainForm.ConnectorTypeDisplay = "Flying Lead" Or MainForm.ConnectorTypeDisplay = "Conduit" Then
    
            Else
                If VoltageReading < 1 Then
                    MainForm.CurrentPass.Visible = True
                Else
                    MainForm.CurrentFail.Visible = True
                    SensorStatus(MainForm.SensorID) = FailedCurrent
                End If
            End If
        End If

    Else
    
        If PinOutSwitch = "T" Then
            SwitchVout1Meas
        ElseIf PinOutSwitch = "Q" Then
            SwitchVout1MeasQ
        ElseIf PinOutSwitch = "M" Then
            SwitchVout1MeasM
        ElseIf PinOutSwitch = "Z" Then
            SwitchVout1MeasZ
        ElseIf PinOutSwitch = "A" Then
            SwitchVout1MeasA
        ElseIf PinOutSwitch = "W" Then
            SwitchVout1MeasW
        ElseIf PinOutSwitch = "L" Then
            SwitchVout1MeasL
        ElseIf PinOutSwitch = "N" Then
            SwitchVout1MeasN
        ElseIf PinOutSwitch = "F" Then
            SwitchVout1MeasF
            If ConnectorType = "C1" Or ConnectorType = "C2" Then 'sumitomo
                SwitchVout1Meas
            End If
        ElseIf PinOutSwitch = "C" Then
            SwitchVout1MeasC
        ElseIf PinOutSwitch = "6" Then
            SwitchVout1Meas6
        End If
       
        If BarcodeUnits = "B" Then
            FSPressure = (Mid$(MainForm.SecondMCSBarcode, 17, 4))
        Else
            FSPressure = (Mid$(MainForm.SecondMCSBarcode, 16, 5) / 14.5)
        End If
             
        SetPSU2 Vsupply
             
        If MainForm.PODCheck = 1 Then
            Sleep 5000
        End If
                
        Vout1Reading = MeasureDigitalMultimeterVolts
       
        Vout1Reading = Vout1Reading - 0.005 ' fudge for cable resistance
       
        VerifySupplyCurrent
    End If
    
    MainForm.VOUT1OutputDisplay = Format$(Vout1Reading, "0.0000")
    Vout1Error = (Vout1Reading - Offset) / OutputSpan * 100
    MainForm.VOUT1OUTPUTERRORDISPLAY = Format$(Vout1Error, "0.000")
          
    If Vout1Reading > LOWER And Vout1Reading < UPPER Then
        MainForm.VOUT1PASS.Visible = True
        MainForm.VOUT1FAIL.Visible = False
    Else
        MainForm.VOUT1PASS.Visible = False
        MainForm.VOUT1FAIL.Visible = True
        SensorStatus(MainForm.SensorID) = FailedOffset1
    End If
    
    DoEvents
    SetPSU2 0
End Sub
Private Sub CHECKVOUT2()
    
Dim Vout2Error As Double
Dim Vout2Reading As Double
Dim SwitchPSUValue As Double
   
'if temp output
    If (Mid$(MainForm.ThirdMCSBarcode, 3, 1) = "T") Then

        SetPSU2 Vsupply
'switch to vout 2

        If PinOutSwitch = "T" Then
            SwitchVout2Meas
        ElseIf PinOutSwitch = "Q" Then
            SwitchVout2MeasQ
        ElseIf PinOutSwitch = "3" Then
            SwitchVout2MeasQ
        ElseIf PinOutSwitch = "M" Then
            SwitchVout2MeasM
        ElseIf PinOutSwitch = "Z" Then
            SwitchVout2MeasZ
        ElseIf PinOutSwitch = "W" Then
            SwitchVout2MeasW
        ElseIf PinOutSwitch = "A" Then
            SwitchVout2MeasA
        ElseIf PinOutSwitch = "L" Then
            SwitchVout2MeasL
        ElseIf PinOutSwitch = "N" Then
            SwitchVout2MeasN
        ElseIf PinOutSwitch = "F" Then
            SwitchVout2MeasF
            If ConnectorType = "C1" Or ConnectorType = "C2" Then 'sumitomo
                SwitchVout2Meas
            End If
        ElseIf PinOutSwitch = "C" Then
            SwitchVout2MeasC
        ElseIf PinOutSwitch = "6" Then
            SwitchVout2Meas6
        End If
        
'Take vout2 reading
        Vout2Reading = MeasureDigitalMultimeterVolts
        MainForm.VOUT2OutputDisplay = Format$(Vout2Reading, "0.0000")
        Vout2Error = (Vout2Target - Vout2Reading) / Span2 * 100
        MainForm.VOUT2OUTPUTERRORDISPLAY = Format$(Vout2Error, "0.000")
        
        If Abs(Vout2Error) > Vout2errorlimit Then
            MainForm.VOUT2PASS.Visible = False
            MainForm.VOUT2FAIL.Visible = True
            SensorStatus(MainForm.SensorID) = FailedTemp
        Else
            MainForm.VOUT2PASS.Visible = True
            MainForm.VOUT2FAIL.Visible = False
            
        End If
    End If
    
'if Switch output
    If (Mid$(MainForm.ThirdMCSBarcode, 3, 1) = "S") And BoardType <> "HY" Then

        SetPSU2 Vsupply
        
        If SwitchHyst = True Then
            SwitchVout1SwMeasA
                            
     'Take vout2 reading
            Vout2Reading = MeasureDigitalMultimeterVolts
            MainForm.VOUT2OutputDisplay = Format$(Vout2Reading, "0.0000")
        
            If (Mid$(MainForm.ThirdMCSBarcode, 4, 1) = "0") Then
                If Vout2Reading < 0.2 Then
                    MainForm.VOUT2PASS.Visible = True
                    MainForm.VOUT2FAIL.Visible = False
                Else
                    MainForm.VOUT2PASS.Visible = False
                    MainForm.VOUT2FAIL.Visible = True
                    SensorStatus(MainForm.SensorID) = FailedSwitch
                End If
        'check output is less than 0.2v
            Else
                'If Vout2Reading > Vsupply - 6 Then
                If Vout2Reading > 4 Then
                    MainForm.VOUT2PASS.Visible = True
                    MainForm.VOUT2FAIL.Visible = False
                Else
                    MainForm.VOUT2PASS.Visible = False
                    MainForm.VOUT2FAIL.Visible = True
                    SensorStatus(MainForm.SensorID) = FailedSwitch
                End If
        'check output is greater than 4v
            End If
            
            SetPSU2 0
        Else
    
            CalculateSwitchTarget
            LowerSwitchTarget = PSUSwitchTarget - SWError
                      
            SetPSU3 LowerSwitchTarget
    
            If PinOutSwitch = "T" Then
                SwitchVout2SwMeasT
            ElseIf PinOutSwitch = "A" Then
                SwitchVout2SwMeasA
            ElseIf PinOutSwitch = "L" Then
                SwitchVout2SwMeasL
            ElseIf PinOutSwitch = "R" Then
                SwitchVout2SwMeasR
            End If
    
        'Take vout2 reading
            Vout2Reading = MeasureDigitalMultimeterVolts
            MainForm.VOUT2OutputDisplay = Format$(Vout2Reading, "0.0000")
        
            If (Mid$(MainForm.ThirdMCSBarcode, 4, 1) = "0") Then
                If Vout2Reading < 0.2 Then
                    MainForm.VOUT2PASS.Visible = True
                    MainForm.VOUT2FAIL.Visible = False
                Else
                    MainForm.VOUT2PASS.Visible = False
                    MainForm.VOUT2FAIL.Visible = True
                    SensorStatus(MainForm.SensorID) = FailedSwitch
                End If
        'check output is less than 0.2v
            Else
                'If Vout2Reading > Vsupply - 1 Then
                If Vout2Reading > 4 Then
                    MainForm.VOUT2PASS.Visible = True
                    MainForm.VOUT2FAIL.Visible = False
                Else
                    MainForm.VOUT2PASS.Visible = False
                    MainForm.VOUT2FAIL.Visible = True
                    SensorStatus(MainForm.SensorID) = FailedSwitch
                End If
        'check output is greater than 4v
            End If
                
            UpperSwitchTarget = PSUSwitchTarget + SWError
                    
            SetPSU3 UpperSwitchTarget
             
    'Take vout2 reading
            Vout2Reading = MeasureDigitalMultimeterVolts
            MainForm.VOUT2OUTPUTERRORDISPLAY = Format$(Vout2Reading, "0.0000")
        
            If (Mid$(MainForm.ThirdMCSBarcode, 4, 1) = "0") Then
                'If Vout2Reading > Vsupply - 1 Then
                If Vout2Reading > 4 Then
                    MainForm.VOUT2SWPASS.Visible = True
                    MainForm.VOUT2SWFAIL.Visible = False
                Else
                    MainForm.VOUT2SWPASS.Visible = False
                    MainForm.VOUT2SWFAIL.Visible = True
                    SensorStatus(MainForm.SensorID) = FailedSwitch
                End If
    'check output is greater than 4v
    
            Else
                If Vout2Reading < 0.2 Then
                    MainForm.VOUT2SWPASS.Visible = True
                    MainForm.VOUT2SWFAIL.Visible = False
                Else
                    MainForm.VOUT2SWPASS.Visible = False
                    MainForm.VOUT2SWFAIL.Visible = True
                    SensorStatus(MainForm.SensorID) = FailedSwitch
                End If
    'check output is less than 0.2v
            End If
            SetPSU2 0
            SetPSU3 0
        End If
    End If
  
End Sub
Private Sub CHECKVOUT2ONLY()

    Dim VSSTCReading As Double
    Dim GNDSTCReading As Double
    Dim VOUT2STCReading As Double
    Dim Vout2Reading  As Double
    
    MainForm.STCTargetDisplay = "OVRFLW"

    SwitchSTCVsMeas
           
'dummy reading
    VSSTCReading = MeasureDigitalMultimeterOhms
    Sleep 100
'take reading
    VSSTCReading = MeasureDigitalMultimeterOhms

    If VSSTCReading > 40000000 Then
        MainForm.STCVSDisplay = "OVRFLW"
    Else
        MainForm.STCVSDisplay = Format$(VSSTCReading, "0.00" & " Ohms")
    End If

    If VSSTCReading > 40000000 Then
        MainForm.STCVSPASS.Visible = True
        MainForm.STCVSFAIL.Visible = False
        STCVSResult = True
    Else
        MainForm.STCVSPASS.Visible = False
        MainForm.STCVSFAIL.Visible = True
        SensorStatus(MainForm.SensorID) = FailedSTC
        STCVSResult = False
    End If

DoEvents

'connect GND to stc

    SwitchSTCGndMeas

'take reading
    GNDSTCReading = MeasureDigitalMultimeterOhms

    If GNDSTCReading > 40000000 Then
        MainForm.STCGNDDisplay = "OVRFLW"
    Else
        MainForm.STCGNDDisplay = Format$(GNDSTCReading, "0.00" & " Ohms")
    End If


    If GNDSTCReading > 40000000 Then
        MainForm.STCGNDPASS.Visible = True
        MainForm.STCGNDFAIL.Visible = False
        STCGndResult = True
    Else
        MainForm.STCGNDPASS.Visible = False
        SensorStatus(MainForm.SensorID) = FailedSTC
        STCGndResult = False
    End If

DoEvents
    
    SwitchSTCVout2Meas

'take reading
    VOUT2STCReading = MeasureDigitalMultimeterOhms

    If VOUT2STCReading > 40000000 Then
        MainForm.STCVOUT2Display = "OVRFLW"
    Else
        VOUT2STCReading = VOUT2STCReading - 1
        MainForm.STCVOUT2Display = Format$(VOUT2STCReading, "0.00" & " Ohms")
    
    End If

    If MainForm.STCTargetDisplay = "OVRFLW" Then
        If VOUT2STCReading > 40000000 Then
            MainForm.STCVOUT2PASS.Visible = True
            MainForm.STCVOUT2FAIL.Visible = False
            STCV2Result = True
        Else
            MainForm.STCVOUT2PASS.Visible = False
            MainForm.STCVOUT2FAIL.Visible = True
            SensorStatus(MainForm.SensorID) = FailedSTC
            STCV2Result = False
        End If
    Else
        If VOUT2STCReading <= 2 Then
            MainForm.STCVOUT2PASS.Visible = True
            MainForm.STCVOUT2FAIL.Visible = False
            STCV2Result = True
        Else
            MainForm.STCVOUT2PASS.Visible = False
            MainForm.STCVOUT2FAIL.Visible = True
            SensorStatus(MainForm.SensorID) = FailedSTC
            STCV2Result = False
        End If
    End If

    SwitchVout2Meas

'Take vout2 reading
        Vout2Reading = MeasureDigitalMultimeterVolts
        MainForm.VOUT2OutputDisplay = Format$(Vout2Reading, "0.0000")
    
        If (Mid$(MainForm.ThirdMCSBarcode, 4, 7) = "000/100") Then
            If Vout2Reading < 0.2 Then
                MainForm.VOUT2PASS.Visible = True
                MainForm.VOUT2FAIL.Visible = False
            Else
                MainForm.VOUT2PASS.Visible = False
                MainForm.VOUT2FAIL.Visible = True
                SensorStatus(MainForm.SensorID) = FailedSwitch
            End If
'check output is less than 0.2v
        Else
            If Vout2Reading > 4 Then
                MainForm.VOUT2PASS.Visible = True
                MainForm.VOUT2FAIL.Visible = False
            Else
                MainForm.VOUT2PASS.Visible = False
                MainForm.VOUT2FAIL.Visible = True
                SensorStatus(MainForm.SensorID) = FailedSwitch
            End If
'check output is greater than 4v
        End If
        
        VerifySupplyCurrent
    
End Sub
Private Sub PASS_FAIL()

Dim l          As Long
Dim lFlags     As Long
Dim sSoundName As String
Dim id As Double

ReadyFlag2 = True               ' was ReadyFlag

If Retests = True Then

    If Abs(Vout1Error) > 0.4 Then
        SensorStatus(MainForm.SensorID) = FailedOffset1
    Else
        SensorStatus(MainForm.SensorID) = PASSED
    End If

End If

If Post25DayTest = True Then
    If Update25DayHoldResult = False Then
        Exit Sub
    End If
End If

If SensorStatus(MainForm.SensorID) <> PASSED Then
      
    sSoundName = "C:\offset setup files\failed.wav"
    PlaySound sSoundName, CLng(0), 1
    MainForm.PASSED.Visible = False
    MainForm.FAILED.Visible = True
    DoEvents
    
    If Abs(Vout1Error) > 1 Then
        MsgBox " EXCESSIVE FAIL RECONNECT AND RE-MEASURE"
    End If
    
Else

AddToHistoryLog "TestComplete = " & Now
EndTime = Now    ' was commnented out

If OffsetType = 1 And MainForm.DisablePrinterCheck = 0 Then
    PrintLabel
End If

    MainForm.PASSED.Visible = True
    MainForm.FAILED.Visible = False
    
End If

If Retests = False Then
    UpdateExcelWithIdResults
End If
    
AddToHistoryLogCDrive "Sub exited"

WorksOrder = (Mid$(MainForm.WorksOrderBarcode, 5, 15))
SensorNumber = MainForm.SensorID.Text
    
AddToHistoryLogCDrive "starting log entry"
    
LogEntry = WorksOrder & "  " & SensorNumber & " " & SensorStatus(MainForm.SensorID) & "  " & "VOUT 1 = " & MainForm.VOUT1OutputDisplay & "," & CSVOffset & "," & OffsetErrorPercent & "," & MainForm.STCVSDisplay & "," & MainForm.STCGNDDisplay & "," & MainForm.STCVOUT1Display & "," & MainForm.STCVOUT2Display & "," & MainForm.VOUT2OutputDisplay & "," & MainForm.CurrentDisplay
AddToHistoryLog LogEntry
    
AddToHistoryLogCDrive "log entry complete"
    
    id = MainForm.SensorID.Text
    id = id + 1
    MainForm.SensorID.Text = id
    SensorNumber = id
   
If PackOnly = False Then

    Dim i As Integer
    
    MainForm.NumberOfUsesDisplay = MainForm.NumberOfUsesDisplay + 1
    
    For i = 1 To NumberOfCableTypes
    
        If CableID(i) = MainForm.CableNumberDisplay Then
        
            CableUsage(i) = MainForm.NumberOfUsesDisplay
    
            If CableUsage(i) = 10000 Then
                Usage.Show
                MainForm.Hide
                
                
            End If
            Exit For
        End If
        
    Next
    
 AddToHistoryLogCDrive "update cable usage"
    
    WriteCableUsageFile ("C:\offset setup files\Cable Usage.txt")
    
 AddToHistoryLogCDrive "update cable usage complete"
    
End If

    AddToHistoryLog "FileSaveComplete = " & Now
    
    CompletedTime = Now
    
    EndTime = Format$(Now, "ttttt")
    TotalTime = EndTime - StartTime
    RecordTime = CompletedTime - EndTime
    
    LogEntry = "TestTime = " & TestTime & "  RecordTime = " & RecordTime
    AddToHistoryLog LogEntry
    
    sSoundName = "C:\offset setup files\complete.wav"
    PlaySound sSoundName, CLng(0), 1

'MainForm.TimeDisplay = Format$(TestTime, "s")

MainForm.EnableStart.Visible = True

AddToHistoryLogCDrive "Completed"


End Sub
Private Sub FindOffset()

    Dim FileHandle As Integer
    Dim FileName As String
    Dim BatchNumber As String
    Dim s As String
    Dim IDSTART As Double
    Dim RigFound As Boolean

    WorksOrder = (Mid$(MainForm.WorksOrderBarcode, 5, 15))
    
    RigFound = False

    On Error GoTo errhandler
    CSVLineRequest = 2
    FileName = "\\USVR8\Results\Production\CSV files for Offset\" & WorksOrder & "-" & "001" & "-" & "0000" & ".csv"
    FileHandle = FreeFile
    Open FileName For Input As #FileHandle
    CSVLine = 0
    While CSVLine < 100
        s = CSVInputLine2(FileHandle)
        CSVLine = CSVLine + 1
    If CSVLine = CSVLineRequest Then
        ReadCsvLine2 s
        If Error2 = "1" Then
            MainForm.RigDisplay = "SCR"
        ElseIf Error2 = "2" Then
            MainForm.RigDisplay = "UCR"
        Else
            MainForm.RigDisplay = "LOW"
        End If
    End If
       
    Wend
    Close #FileHandle
           
    If MainForm.RigDisplay = "SCR" Then

        If MainForm.SensorID <= 40 Then
            BatchNumber = "001"
            IDSTART = 0
        ElseIf MainForm.SensorID <= 80 Then
            BatchNumber = "041"
            IDSTART = 40
        ElseIf MainForm.SensorID <= 120 Then
            BatchNumber = "081"
            IDSTART = 80
        ElseIf MainForm.SensorID <= 160 Then
            BatchNumber = "121"
            IDSTART = 120
        ElseIf MainForm.SensorID <= 200 Then
            BatchNumber = "161"
            IDSTART = 160
        ElseIf MainForm.SensorID <= 240 Then
            BatchNumber = "201"
            IDSTART = 200
        ElseIf MainForm.SensorID <= 280 Then
            BatchNumber = "241"
            IDSTART = 240
        ElseIf MainForm.SensorID <= 320 Then
            BatchNumber = "281"
            IDSTART = 280
        ElseIf MainForm.SensorID <= 360 Then
            BatchNumber = "321"
            IDSTART = 320
        ElseIf MainForm.SensorID <= 400 Then
            BatchNumber = "361"
            IDSTART = 360
        ElseIf MainForm.SensorID <= 440 Then
            BatchNumber = "401"
            IDSTART = 400
        ElseIf MainForm.SensorID <= 480 Then
            BatchNumber = "441"
            IDSTART = 440
        ElseIf MainForm.SensorID <= 520 Then
            BatchNumber = "481"
            IDSTART = 480
        ElseIf MainForm.SensorID <= 560 Then
            BatchNumber = "521"
            IDSTART = 520
        ElseIf MainForm.SensorID <= 600 Then
            BatchNumber = "561"
            IDSTART = 560
        ElseIf MainForm.SensorID <= 640 Then
            BatchNumber = "601"
            IDSTART = 600
        ElseIf MainForm.SensorID <= 680 Then
            BatchNumber = "641"
            IDSTART = 640
        ElseIf MainForm.SensorID <= 720 Then
            BatchNumber = "681"
            IDSTART = 680
        ElseIf MainForm.SensorID <= 760 Then
            BatchNumber = "721"
            IDSTART = 720
        ElseIf MainForm.SensorID <= 800 Then
            BatchNumber = "761"
            IDSTART = 760
        ElseIf MainForm.SensorID <= 840 Then
            BatchNumber = "801"
            IDSTART = 800
        ElseIf MainForm.SensorID <= 880 Then
            BatchNumber = "841"
            IDSTART = 840
        ElseIf MainForm.SensorID <= 920 Then
            BatchNumber = "881"
            IDSTART = 880
        ElseIf MainForm.SensorID <= 960 Then
            BatchNumber = "921"
            IDSTART = 920
        ElseIf MainForm.SensorID <= 1000 Then
            BatchNumber = "961"
            IDSTART = 960
        ElseIf MainForm.SensorID <= 1040 Then
            BatchNumber = "1001"
            IDSTART = 1000
        Else
            MsgBox "id error"
        End If
    
    
    ElseIf MainForm.RigDisplay = "UCR" Then
        
        If MainForm.SensorID <= 42 Then
            BatchNumber = "001"
            IDSTART = 0
        ElseIf MainForm.SensorID <= 84 Then
            BatchNumber = "043"
            IDSTART = 42
        ElseIf MainForm.SensorID <= 126 Then
            BatchNumber = "085"
            IDSTART = 84
        ElseIf MainForm.SensorID <= 168 Then
            BatchNumber = "127"
            IDSTART = 126
        ElseIf MainForm.SensorID <= 210 Then
            BatchNumber = "169"
            IDSTART = 168
        ElseIf MainForm.SensorID <= 252 Then
            BatchNumber = "211"
            IDSTART = 210
        ElseIf MainForm.SensorID <= 294 Then
            BatchNumber = "253"
            IDSTART = 252
        ElseIf MainForm.SensorID <= 336 Then
            BatchNumber = "295"
            IDSTART = 294
        ElseIf MainForm.SensorID <= 378 Then
            BatchNumber = "337"
            IDSTART = 336
        ElseIf MainForm.SensorID <= 420 Then
            BatchNumber = "379"
            IDSTART = 378
        ElseIf MainForm.SensorID <= 462 Then
            BatchNumber = "421"
            IDSTART = 420
        ElseIf MainForm.SensorID <= 504 Then
            BatchNumber = "463"
            IDSTART = 462
        ElseIf MainForm.SensorID <= 546 Then
            BatchNumber = "505"
            IDSTART = 504
        ElseIf MainForm.SensorID <= 588 Then
            BatchNumber = "547"
            IDSTART = 546
        ElseIf MainForm.SensorID <= 630 Then
            BatchNumber = "589"
            IDSTART = 588
        ElseIf MainForm.SensorID <= 672 Then
            BatchNumber = "631"
            IDSTART = 630
        ElseIf MainForm.SensorID <= 714 Then
            BatchNumber = "673"
            IDSTART = 672
        ElseIf MainForm.SensorID <= 756 Then
            BatchNumber = "715"
            IDSTART = 714
        ElseIf MainForm.SensorID <= 798 Then
            BatchNumber = "757"
            IDSTART = 756
        ElseIf MainForm.SensorID <= 840 Then
            BatchNumber = "799"
            IDSTART = 798
        ElseIf MainForm.SensorID <= 882 Then
            BatchNumber = "841"
            IDSTART = 840
        ElseIf MainForm.SensorID <= 924 Then
            BatchNumber = "883"
            IDSTART = 882
        ElseIf MainForm.SensorID <= 966 Then
            BatchNumber = "925"
            IDSTART = 924
        ElseIf MainForm.SensorID <= 1008 Then
            BatchNumber = "967"
            IDSTART = 966
        Else
            MsgBox "id error"
        End If
    Else
        If MainForm.SensorID <= 14 Then
            BatchNumber = "001"
            IDSTART = 0
        ElseIf MainForm.SensorID <= 28 Then
            BatchNumber = "015"
            IDSTART = 14
        ElseIf MainForm.SensorID <= 42 Then
            BatchNumber = "029"
            IDSTART = 28
        ElseIf MainForm.SensorID <= 56 Then
            BatchNumber = "043"
            IDSTART = 42
        ElseIf MainForm.SensorID <= 70 Then
            BatchNumber = "057"
            IDSTART = 56
        ElseIf MainForm.SensorID <= 84 Then
            BatchNumber = "071"
            IDSTART = 70
        ElseIf MainForm.SensorID <= 98 Then
            BatchNumber = "085"
            IDSTART = 84
        ElseIf MainForm.SensorID <= 112 Then
            BatchNumber = "099"
            IDSTART = 98
        ElseIf MainForm.SensorID <= 126 Then
            BatchNumber = "113"
            IDSTART = 112
        ElseIf MainForm.SensorID <= 140 Then
            BatchNumber = "127"
            IDSTART = 126
        ElseIf MainForm.SensorID <= 154 Then
            BatchNumber = "141"
            IDSTART = 140
        ElseIf MainForm.SensorID <= 168 Then
            BatchNumber = "155"
            IDSTART = 154
        ElseIf MainForm.SensorID <= 182 Then
            BatchNumber = "169"
            IDSTART = 168
        ElseIf MainForm.SensorID <= 196 Then
            BatchNumber = "183"
            IDSTART = 182
        ElseIf MainForm.SensorID <= 210 Then
            BatchNumber = "197"
            IDSTART = 196
        ElseIf MainForm.SensorID <= 224 Then
            BatchNumber = "211"
            IDSTART = 210
        ElseIf MainForm.SensorID <= 238 Then
            BatchNumber = "225"
            IDSTART = 224
        ElseIf MainForm.SensorID <= 252 Then
            BatchNumber = "239"
            IDSTART = 238
        Else
            MsgBox "id error"
        End If
    End If
     
    On Error GoTo errhandler
    CSVLineRequest = (MainForm.SensorID - IDSTART) + 50
    FileName = "\\USVR8\Results\Production\CSV files for Offset\" & WorksOrder & "-" & BatchNumber & "-" & "0000" & ".csv"
    FileHandle = FreeFile
    Open FileName For Input As #FileHandle
    CSVLine = 0
    While EOF(FileHandle) = False
        s = CSVInputLine(FileHandle)
        If PartNotCal = True Then
            Close #FileHandle
            Exit Sub
        End If
    Wend
    Close #FileHandle

    Exit Sub

errhandler:

    MsgBox "Can't Find File - Check UCR"
    PartNotCal = True
CSVOffsetError = 100
End Sub
Public Function CSVInputLine2(ByVal FileHandle As Integer) As String
    CSVInputLine2 = vbNullString
    Dim NewCharacter As String
    Dim s As String
    
    While (EOF(FileHandle) = False) And (NewCharacter <> vbCr)
        NewCharacter = Input(1, FileHandle)
        CSVInputLine2 = CSVInputLine2 & NewCharacter
    Wend
    

End Function
Public Function CSVInputLine(ByVal FileHandle As Integer) As String
    CSVInputLine = vbNullString
    Dim NewCharacter As String
    Dim s As String
    
    While (EOF(FileHandle) = False) And (NewCharacter <> vbCr)
        NewCharacter = Input(1, FileHandle)
        If NewCharacter = vbLf Then
            Exit Function
        Else
            CSVInputLine = CSVInputLine & NewCharacter
        End If
    Wend
    CSVLine = CSVLine + 1
    If CSVLine = CSVLineRequest Then
    s = CSVInputLine
        ReadCsvLine s
        If PartNotCal = True Then
            Exit Function
        End If
    End If
End Function
Private Function ReadCsvLine(ByVal s As String)

Dim Error As String
Dim SplitValues() As String
    
    SplitValues = Split(s, ",")
    Error = SplitValues(1)
    
    If Error <> "SSPassed" Then
    
        MsgBox "PART DID NOT PASS CALIBRATION"
        PartNotCal = True
        SensorStatus(MainForm.SensorID) = DidNotPassCalibration
    Exit Function
    

    Else
        If OutputType = "C" Then
            CSVOffset = Format$((SplitValues(34) * 1000), "0.0000")
        Else
            CSVOffset = Format$(SplitValues(34), "0.0000")
        End If
        
        CSVOffsetError = Format$(SplitValues(35), "0.0000")
    End If
    
End Function
Private Function ReadCsvLine2(ByVal s As String)


Dim SplitValues() As String
    
    SplitValues = Split(s, " ")
    Error2 = SplitValues(0)
     
End Function
Private Sub CalculateSwitchTarget()

    If BarcodeUnits = "B" Then
        SWPressure = (Mid$(MainForm.ThirdMCSBarcode, 17, 4))
    Else
        SWPressure = (Mid$(MainForm.ThirdMCSBarcode, 16, 5) / 14.5)
    End If

    PSUSwitchTarget = (SWPressure / FSPressure * OutputSpan) + Offset
 '   SWError = OutputSpan / 100 * 0.9
    SWError = 0.08
End Sub
Public Sub TestSTC()

    Dim VSSTCReading  As Double
    Dim VOUT1STCReading As Double
    Dim VOUT2STCReading As Double
    Dim GNDSTCReading As Double

'connect VS to stc
    SetPSU2 0
    
    SwitchSTCVsMeas

'take reading
    VSSTCReading = MeasureDigitalMultimeterOhms
    MainForm.STCVSDisplay = VSSTCReading
    
    If VSSTCReading > 10000 Then
        MainForm.STCVSDisplay = "OVRFLW"
        MainForm.STCVSPASS.Visible = True
        MainForm.STCVSFAIL.Visible = False
    Else
        MainForm.STCVSPASS.Visible = False
        MainForm.STCVSFAIL.Visible = True
        SensorStatus(1) = FailedSTC
    End If

DoEvents

'connect GND to stc

    SwitchSTCGndMeas
        
'take reading
    GNDSTCReading = MeasureDigitalMultimeterOhms
    MainForm.STCGNDDisplay = GNDSTCReading
    
    If GNDSTCReading > 1000 Then
        MainForm.STCGNDDisplay = "OVRFLW"
        MainForm.STCGNDPASS.Visible = True
        MainForm.STCGNDFAIL.Visible = False
    Else
        MainForm.STCGNDPASS.Visible = False
        MainForm.STCGNDFAIL.Visible = True
        SensorStatus(1) = FailedSTC
    End If
DoEvents
'connect Vout1 to stc

    SwitchSTCVout1Meas
        
'take reading
    VOUT1STCReading = MeasureDigitalMultimeterOhms
    MainForm.STCVOUT1Display = VOUT1STCReading
    
    If VOUT1STCReading > 10000 Then
        MainForm.STCVOUT1Display = "OVRFLW"
        MainForm.STCVOUT1PASS.Visible = True
        MainForm.STCVOUT1FAIL.Visible = False
    Else
        MainForm.STCVOUT1PASS.Visible = False
        MainForm.STCVOUT1FAIL.Visible = True
        SensorStatus(1) = FailedSTC
    End If
    
DoEvents

'connect Vout2 to stc

    SwitchSTCVout2Meas
        
'take reading
    VOUT2STCReading = MeasureDigitalMultimeterOhms
    MainForm.STCVOUT2Display = VOUT2STCReading
    
    If VOUT2STCReading > 10000 Then
        MainForm.STCVOUT2Display = "OVRFLW"
        MainForm.STCVOUT2PASS.Visible = True
        MainForm.STCVOUT2FAIL.Visible = False
    Else
        MainForm.STCVOUT2PASS.Visible = False
        MainForm.STCVOUT2FAIL.Visible = True
        SensorStatus(1) = FailedSTC
    End If
    
    
End Sub
Public Sub TestSTC2()

    Dim VSSTCReading  As Double
    Dim VOUT1STCReading As Double
    Dim VOUT2STCReading As Double
    Dim GNDSTCReading As Double

'connect VS to stc
    
    SetPSU2 0
    
    SwitchSTCVsMeas

'take reading
    VSSTCReading = MeasureDigitalMultimeterOhms
    VSSTCReading = VSSTCReading - 4.5
    MainForm.STCVSDisplay = VSSTCReading
    
    If VSSTCReading < 2 Then
        MainForm.STCVSPASS.Visible = True
        MainForm.STCVSFAIL.Visible = False
    Else
        MainForm.STCVSPASS.Visible = False
        MainForm.STCVSFAIL.Visible = True
        SensorStatus(1) = FailedSTC
    End If

DoEvents

'connect GND to stc

    SwitchSTCGndMeas

'take reading
    
    GNDSTCReading = MeasureDigitalMultimeterOhms
    GNDSTCReading = GNDSTCReading - 4.5
    MainForm.STCGNDDisplay = GNDSTCReading
    
    If GNDSTCReading < 2 Then
        MainForm.STCGNDPASS.Visible = True
        MainForm.STCGNDFAIL.Visible = False
    Else
        MainForm.STCGNDPASS.Visible = False
        MainForm.STCGNDFAIL.Visible = True
        SensorStatus(1) = FailedSTC
    End If
DoEvents
'connect Vout1 to stc

    SwitchSTCVout1Meas
    
'take reading
    
    VOUT1STCReading = MeasureDigitalMultimeterOhms
    VOUT1STCReading = VOUT1STCReading - 4.5
    MainForm.STCVOUT1Display = VOUT1STCReading
    
    If VOUT1STCReading < 2 Then
        MainForm.STCVOUT1PASS.Visible = True
        MainForm.STCVOUT1FAIL.Visible = False
    Else
        MainForm.STCVOUT1PASS.Visible = False
        MainForm.STCVOUT1FAIL.Visible = True
        SensorStatus(1) = FailedSTC
    End If
    
DoEvents

'connect Vout2 to stc

    SwitchSTCVout2Meas

'take reading
    
    VOUT2STCReading = MeasureDigitalMultimeterOhms
    VOUT2STCReading = VOUT2STCReading - 4.5
    MainForm.STCVOUT2Display = VOUT2STCReading
    
    If VOUT2STCReading < 2 Then
        
        MainForm.STCVOUT2PASS.Visible = True
        MainForm.STCVOUT2FAIL.Visible = False
    Else
        MainForm.STCVOUT2PASS.Visible = False
        MainForm.STCVOUT2FAIL.Visible = True
        SensorStatus(1) = FailedSTC
    End If
    
    
End Sub
Public Sub TestVout1and2()

    Dim Vout1Reading As Double
    Dim Vout2Reading As Double

'Connect to vout 1
    SetPSU2 24
    SwitchVout1Meas

    Vout1Reading = MeasureDigitalMultimeterVolts

    MainForm.VOUT1OutputDisplay = Format$(Vout1Reading, "0.0000")

    If Vout1Reading > 0.99 And Vout1Reading < 1.02 Then
        MainForm.VOUT1PASS.Visible = True
        MainForm.VOUT1FAIL.Visible = False
    Else
        MainForm.VOUT1PASS.Visible = False
        MainForm.VOUT1FAIL.Visible = True
        SensorStatus(1) = FailedOffset1
    End If
        
    DoEvents
'Connect to vout 2


    SwitchVout2Meas

    Vout2Reading = MeasureDigitalMultimeterVolts

    MainForm.VOUT2OutputDisplay = Format$(Vout2Reading, "0.0000")

    If Vout2Reading > 2 And Vout2Reading < 3 Then
        MainForm.VOUT2PASS.Visible = True
        MainForm.VOUT2FAIL.Visible = False
    Else
        MainForm.VOUT2PASS.Visible = False
        MainForm.VOUT2FAIL.Visible = True
        SensorStatus(1) = FailedTemp
    End If

    DoEvents
  
End Sub
Public Sub TestOringAndRest()

    ORingRequired = False
    RestrictorRequired = False
    ChangeVisionProgram 35
    RecieveVision
    
End Sub
Public Sub TestOringAndRest2()

    ORingRequired = True
    RestrictorRequired = True
    ChangeVisionProgram 35
    RecieveVision

End Sub

Public Sub CheckLoad()

If LoadValue > 500 Then
LoadType = "K   PullUp"
LoadValue = LoadValue - 500
Else
LoadType = "K   PullDown"
End If
MainForm.LoadTypeDisplay.Caption = LoadType
  
If LoadValue = 6 Then
LoadValue = 1.7
End If

If LoadValue = 7 Then
LoadValue = 2
End If
 
If LoadValue = 8 Then
LoadValue = 2.5
End If

If LoadValue = 9 Then
LoadValue = 2.7
End If

If LoadValue = 10 Then
LoadValue = 3
End If

If LoadValue = 11 Then
LoadValue = 3.2
End If

If LoadValue = 12 Then
LoadValue = 3.7
End If

If LoadValue = 13 Then
LoadValue = 4
End If

If LoadValue = 14 Then
LoadValue = 4.2
End If

If LoadValue = 15 Then
LoadValue = 4.5
End If

If LoadValue = 16 Then
LoadValue = 5
End If

If LoadValue = 17 Then
LoadValue = 5.2
End If

If LoadValue = 18 Then
LoadValue = 5.5
End If

If LoadValue = 19 Then
LoadValue = 5.7
End If

If LoadValue = 20 Then
LoadValue = 6.2
End If

If LoadValue = 21 Then
LoadValue = 6.5
End If

If LoadValue = 22 Then
LoadValue = 6.7
End If

If LoadValue = 23 Then
LoadValue = 7
End If

If LoadValue = 24 Then
LoadValue = 7.5
End If

If LoadValue = 25 Then
LoadValue = 7.7
End If

If LoadValue = 26 Then
LoadValue = 8
End If

If LoadValue = 27 Then
LoadValue = 8.2
End If

If LoadValue = 28 Then
LoadValue = 8.7
End If

If LoadValue = 29 Then
LoadValue = 9
End If

If LoadValue = 30 Then
LoadValue = 9.2
End If

If LoadValue = 31 Then
LoadValue = 9.5
End If

If LoadValue = 32 Then
LoadValue = 10
End If

If LoadValue = 33 Then
LoadValue = 10.3
End If

If LoadValue = 34 Then
LoadValue = 10.5
End If

If LoadValue = 35 Then
LoadValue = 10.7
End If

If LoadValue = 36 Then
LoadValue = 11.3
End If

If LoadValue = 37 Then
LoadValue = 11.5
End If

If LoadValue = 38 Then
LoadValue = 11.7
End If

If LoadValue = 39 Then
LoadValue = 12
End If

If LoadValue = 40 Then
LoadValue = 12.5
End If

If LoadValue = 41 Then
LoadValue = 12.7
End If

If LoadValue = 42 Then
LoadValue = 13
End If

If LoadValue = 43 Then
LoadValue = 13.2
End If

If LoadValue = 44 Then
LoadValue = 13.7
End If

If LoadValue = 45 Then
LoadValue = 14
End If

If LoadValue = 46 Then
LoadValue = 14.2
End If

If LoadValue = 47 Then
LoadValue = 14.5
End If

If LoadValue = 48 Then
LoadValue = 15
End If

If LoadValue = 49 Then
LoadValue = 15.2
End If

If LoadValue = 50 Then
LoadValue = 15.5
End If

If LoadValue = 51 Then
LoadValue = 15.7
End If

If LoadValue = 52 Then
LoadValue = 16.2
End If

If LoadValue = 53 Then
LoadValue = 16.5
End If

If LoadValue = 54 Then
LoadValue = 16.7
End If

If LoadValue = 55 Then
LoadValue = 17
End If

If LoadValue = 56 Then
LoadValue = 17.5
End If

If LoadValue = 57 Then
LoadValue = 17.7
End If

If LoadValue = 58 Then
LoadValue = 18
End If

If LoadValue = 59 Then
LoadValue = 18.2
End If

If LoadValue = 60 Then
LoadValue = 18.7
End If

If LoadValue = 61 Then
LoadValue = 19
End If

If LoadValue = 62 Then
LoadValue = 19.2
End If

If LoadValue = 63 Then
LoadValue = 19.5
End If

If LoadValue = 64 Then
LoadValue = 20
End If

If LoadValue = 65 Then
LoadValue = 20.3
End If

If LoadValue = 66 Then
LoadValue = 20.5
End If

If LoadValue = 67 Then
LoadValue = 20.75
End If

If LoadValue = 68 Then
LoadValue = 21.25
End If

If LoadValue = 69 Then
LoadValue = 21.5
End If

If LoadValue = 70 Then
LoadValue = 21.7
End If

If LoadValue = 71 Then
LoadValue = 22
End If

If LoadValue = 72 Then
LoadValue = 22.5
End If

If LoadValue = 73 Then
LoadValue = 22.7
End If

If LoadValue = 74 Then
LoadValue = 23
End If

If LoadValue = 75 Then
LoadValue = 23.2
End If

If LoadValue = 76 Then
LoadValue = 23.7
End If

If LoadValue = 77 Then
LoadValue = 24
End If

If LoadValue = 78 Then
LoadValue = 24.2
End If

If LoadValue = 79 Then
LoadValue = 24.5
End If

If LoadValue = 80 Then
LoadValue = 25
End If

If LoadValue = 81 Then
LoadValue = 25.2
End If

If LoadValue = 82 Then
LoadValue = 25.5
End If

If LoadValue = 83 Then
LoadValue = 25.7
End If

If LoadValue = 84 Then
LoadValue = 26.2
End If

If LoadValue = 85 Then
LoadValue = 26.5
End If

If LoadValue = 86 Then
LoadValue = 26.7
End If

If LoadValue = 87 Then
LoadValue = 27
End If

If LoadValue = 88 Then
LoadValue = 27.5
End If

If LoadValue = 89 Then
LoadValue = 27.7
End If

If LoadValue = 90 Then
LoadValue = 28
End If

If LoadValue = 91 Then
LoadValue = 28.2
End If

If LoadValue = 92 Then
LoadValue = 28.7
End If

If LoadValue = 93 Then
LoadValue = 29
End If

If LoadValue = 94 Then
LoadValue = 29.2
End If

If LoadValue = 95 Then
LoadValue = 29.5
End If

If LoadValue = 96 Then
LoadValue = 30
End If

If LoadValue = 97 Then
LoadValue = 30.25
End If

If LoadValue = 98 Then
LoadValue = 30.5
End If

If LoadValue = 99 Then
LoadValue = 30.75
End If

If LoadValue = 100 Then
LoadValue = 31.2
End If

If LoadValue = 101 Then
LoadValue = 31.5
End If

If LoadValue = 102 Then
LoadValue = 31.7
End If

If LoadValue = 103 Then
LoadValue = 32
End If

If LoadValue = 104 Then
LoadValue = 32.5
End If

If LoadValue = 105 Then
LoadValue = 32.7
End If

If LoadValue = 106 Then
LoadValue = 33
End If

If LoadValue = 107 Then
LoadValue = 33.2
End If

If LoadValue = 108 Then
LoadValue = 33.7
End If

If LoadValue = 109 Then
LoadValue = 34
End If

If LoadValue = 110 Then
LoadValue = 34.2
End If

If LoadValue = 111 Then
LoadValue = 34.5
End If

If LoadValue = 112 Then
LoadValue = 35
End If

If LoadValue = 113 Then
LoadValue = 35.2
End If

If LoadValue = 114 Then
LoadValue = 35.5
End If

If LoadValue = 115 Then
LoadValue = 35.7
End If

If LoadValue = 116 Then
LoadValue = 36.2
End If

If LoadValue = 117 Then
LoadValue = 36.5
End If

If LoadValue = 118 Then
LoadValue = 36.7
End If

If LoadValue = 119 Then
LoadValue = 37
End If

If LoadValue = 120 Then
LoadValue = 37.5
End If

If LoadValue = 121 Then
LoadValue = 37.7
End If

If LoadValue = 122 Then
LoadValue = 38
End If

If LoadValue = 123 Then
LoadValue = 38.2
End If

If LoadValue = 124 Then
LoadValue = 38.7
End If

If LoadValue = 125 Then
LoadValue = 39
End If

If LoadValue = 126 Then
LoadValue = 39.2
End If

If LoadValue = 127 Then
LoadValue = 39.5
End If

If LoadValue = 128 Then
LoadValue = 40.2
End If

If LoadValue = 129 Then
LoadValue = 40.5
End If

If LoadValue = 130 Then
LoadValue = 40.7
End If

If LoadValue = 131 Then
LoadValue = 41
End If

If LoadValue = 132 Then
LoadValue = 41.5
End If

If LoadValue = 133 Then
LoadValue = 41.7
End If

If LoadValue = 134 Then
LoadValue = 42
End If

If LoadValue = 135 Then
LoadValue = 42.2
End If

If LoadValue = 136 Then
LoadValue = 42.7
End If

If LoadValue = 137 Then
LoadValue = 43
End If

If LoadValue = 138 Then
LoadValue = 43.2
End If

If LoadValue = 139 Then
LoadValue = 43.5
End If

If LoadValue = 140 Then
LoadValue = 44
End If

If LoadValue = 141 Then
LoadValue = 44.2
End If

If LoadValue = 142 Then
LoadValue = 44.5
End If

If LoadValue = 143 Then
LoadValue = 44.7
End If

If LoadValue = 144 Then
LoadValue = 45.2
End If

If LoadValue = 145 Then
LoadValue = 45.5
End If

If LoadValue = 146 Then
LoadValue = 45.7
End If

If LoadValue = 147 Then
LoadValue = 46
End If

If LoadValue = 148 Then
LoadValue = 46.5
End If

If LoadValue = 149 Then
LoadValue = 46.7
End If

If LoadValue = 150 Then
LoadValue = 47
End If

If LoadValue = 151 Then
LoadValue = 47.2
End If

If LoadValue = 152 Then
LoadValue = 47.7
End If

If LoadValue = 153 Then
LoadValue = 48
End If

If LoadValue = 154 Then
LoadValue = 48.2
End If

If LoadValue = 155 Then
LoadValue = 48.5
End If

If LoadValue = 156 Then
LoadValue = 49
End If

If LoadValue = 157 Then
LoadValue = 49.2
End If

If LoadValue = 158 Then
LoadValue = 49.5
End If

If LoadValue = 159 Then
LoadValue = 49.7
End If

If LoadValue = 160 Then
LoadValue = 50.2
End If

If LoadValue = 161 Then
LoadValue = 50.5
End If

If LoadValue = 162 Then
LoadValue = 50.7
End If

If LoadValue = 163 Then
LoadValue = 51
End If

If LoadValue = 164 Then
LoadValue = 51.5
End If

If LoadValue = 165 Then
LoadValue = 51.7
End If

If LoadValue = 166 Then
LoadValue = 52
End If

If LoadValue = 167 Then
LoadValue = 52.2
End If

If LoadValue = 168 Then
LoadValue = 52.7
End If

If LoadValue = 169 Then
LoadValue = 53
End If

If LoadValue = 170 Then
LoadValue = 53.2
End If

If LoadValue = 171 Then
LoadValue = 53.4
End If

If LoadValue = 172 Then
LoadValue = 54
End If

If LoadValue = 173 Then
LoadValue = 54.2
End If

If LoadValue = 174 Then
LoadValue = 54.4
End If

If LoadValue = 175 Then
LoadValue = 54.7
End If

If LoadValue = 176 Then
LoadValue = 55.2
End If

If LoadValue = 177 Then
LoadValue = 55.5
End If

If LoadValue = 178 Then
LoadValue = 55.7
End If

If LoadValue = 179 Then
LoadValue = 56
End If

If LoadValue = 180 Then
LoadValue = 56.4
End If

If LoadValue = 181 Then
LoadValue = 56.7
End If

If LoadValue = 182 Then
LoadValue = 57
End If

If LoadValue = 183 Then
LoadValue = 57.2
End If

If LoadValue = 184 Then
LoadValue = 57.7
End If

If LoadValue = 185 Then
LoadValue = 58
End If

If LoadValue = 186 Then
LoadValue = 58.2
End If

If LoadValue = 187 Then
LoadValue = 58.4
End If

If LoadValue = 188 Then
LoadValue = 59
End If

If LoadValue = 189 Then
LoadValue = 59.3
End If

If LoadValue = 190 Then
LoadValue = 59.5
End If

If LoadValue = 191 Then
LoadValue = 59.7
End If

If LoadValue = 192 Then
LoadValue = 60.2
End If

If LoadValue = 193 Then
LoadValue = 60.5
End If

If LoadValue = 194 Then
LoadValue = 60.7
End If

If LoadValue = 195 Then
LoadValue = 61
End If

If LoadValue = 196 Then
LoadValue = 61.5
End If

If LoadValue = 197 Then
LoadValue = 61.7
End If

If LoadValue = 198 Then
LoadValue = 62
End If

If LoadValue = 199 Then
LoadValue = 62.2
End If

If LoadValue = 200 Then
LoadValue = 62.7
End If

If LoadValue = 201 Then
LoadValue = 63
End If

If LoadValue = 202 Then
LoadValue = 63.2
End If

If LoadValue = 203 Then
LoadValue = 63.5
End If

If LoadValue = 204 Then
LoadValue = 64
End If

If LoadValue = 205 Then
LoadValue = 64.2
End If

If LoadValue = 206 Then
LoadValue = 64.5
End If

If LoadValue = 207 Then
LoadValue = 64.7
End If

If LoadValue = 208 Then
LoadValue = 65.2
End If

If LoadValue = 209 Then
LoadValue = 65.5
End If

If LoadValue = 210 Then
LoadValue = 65.7
End If

If LoadValue = 211 Then
LoadValue = 66
End If

If LoadValue = 212 Then
LoadValue = 66.5
End If

If LoadValue = 213 Then
LoadValue = 66.7
End If

If LoadValue = 214 Then
LoadValue = 67
End If

If LoadValue = 215 Then
LoadValue = 67.2
End If

If LoadValue = 216 Then
LoadValue = 67.7
End If

If LoadValue = 217 Then
LoadValue = 68
End If

If LoadValue = 218 Then
LoadValue = 68.2
End If

If LoadValue = 219 Then
LoadValue = 68.5
End If

If LoadValue = 220 Then
LoadValue = 69
End If

If LoadValue = 221 Then
LoadValue = 69.2
End If

If LoadValue = 222 Then
LoadValue = 69.5
End If

If LoadValue = 223 Then
LoadValue = 69.7
End If

If LoadValue = 224 Then
LoadValue = 70.2
End If

If LoadValue = 225 Then
LoadValue = 70.5
End If

If LoadValue = 226 Then
LoadValue = 70.7
End If

If LoadValue = 227 Then
LoadValue = 71
End If

If LoadValue = 228 Then
LoadValue = 71.5
End If

If LoadValue = 229 Then
LoadValue = 71.7
End If

If LoadValue = 230 Then
LoadValue = 72
End If

If LoadValue = 231 Then
LoadValue = 72.2
End If

If LoadValue = 232 Then
LoadValue = 72.7
End If

If LoadValue = 233 Then
LoadValue = 73
End If

If LoadValue = 234 Then
LoadValue = 73.2
End If

If LoadValue = 235 Then
LoadValue = 73.5
End If

If LoadValue = 236 Then
LoadValue = 74
End If

If LoadValue = 237 Then
LoadValue = 74.2
End If

If LoadValue = 238 Then
LoadValue = 74.5
End If

If LoadValue = 239 Then
LoadValue = 74.7
End If

If LoadValue = 240 Then
LoadValue = 75
End If

If LoadValue = 241 Then
LoadValue = 75.5
End If

If LoadValue = 242 Then
LoadValue = 75.7
End If

If LoadValue = 243 Then
LoadValue = 76
End If

If LoadValue = 244 Then
LoadValue = 76.5
End If

If LoadValue = 245 Then
LoadValue = 76.7
End If

If LoadValue = 246 Then
LoadValue = 77
End If

If LoadValue = 247 Then
LoadValue = 77.2
End If

If LoadValue = 248 Then
LoadValue = 77
End If

If LoadValue = 249 Then
LoadValue = 78
End If

If LoadValue = 250 Then
LoadValue = 78.2
End If

If LoadValue = 251 Then
LoadValue = 78.5
End If

If LoadValue = 252 Then
LoadValue = 79
End If

If LoadValue = 253 Then
LoadValue = 79.2
End If

If LoadValue = 254 Then
LoadValue = 79.5
End If

If LoadValue = 255 Then
LoadValue = 79.7
End If
 
End Sub

Public Sub VerifySupplyCurrent(Optional ByVal DefaultLower As Double = 2#, Optional ByVal DefaultUpper As Double = 5.6)
    Dim SupplyCurrent As Double
    Dim LowerLim As Double
    Dim UpperLim As Double
    Dim ProductRange As String
    
    SupplyCurrent = ReadCurrent()
    SupplyCurrent = SupplyCurrent * 1000
    MainForm.CurrentDisplay = Format$(SupplyCurrent, "0.00" & " mA")
    
    ProductRange = Mid$(WorksOrder, 1, 2)
    GetCurrentLimits ProductRange, LowerLim, UpperLim, DefaultLower, DefaultUpper
    
    If SupplyCurrent > LowerLim And SupplyCurrent < UpperLim Then
        MainForm.CurrentPass.Visible = True
    Else
        MainForm.CurrentFail.Visible = True
        SensorStatus(MainForm.SensorID) = FailedCurrent
    End If
End Sub


