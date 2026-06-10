Attribute VB_Name = "DigitalMultimeterControl"
Option Explicit
Public Const PortNumber As Integer = 16
Public Sub SendDMMCommand(ByVal Command As String)
    Dim Status As Integer

    send PortNumber, Command, Status
    
End Sub
Public Function ReceiveDMMReply() As String
  'Read DMM status and report IEEE status in stat%
    Dim Status As Integer
    Dim l As Integer
    Dim Reply As String

    enter Reply, 255, l, PortNumber, Status
 
    ReceiveDMMReply = Reply
End Function
' This subroutine initialises the DMM for general use by the program.
' It does not set it for any particular measurement type (current or voltage) since
' this can vary from one reading to the next as two different calibration
' types can be run at the same time with the measurments potentially interleaved
'
Public Sub InitialiseDMM()
  
  Dim Reply As String
  
        'Reset and check device is present
        SendDMMCommand "*CLS"   ' clear status
        SendDMMCommand "*RST"   ' reset
        
        SendDMMCommand "*IDN?"
        
End Sub
Public Function MeasureDigitalMultimeterVolts() As Double
    Dim Reply As String
    
        SendDMMCommand ":FUNC 'VOLTAGE:DC'"
        SendDMMCommand ":VOLTAGE:DC:RANGE:AUTO 1"
        SendDMMCommand "READ?"
        Reply = ReceiveDMMReply()
        MeasureDigitalMultimeterVolts = Val(Reply)
        
End Function
Public Function MeasureDigitalMultimeterAmps() As Double
    Dim Reply As String
   
        SendDMMCommand ":FUNC 'CURRENT:DC'"
        SendDMMCommand ":CURRENT:DC:RANGE:AUTO 1"
        SendDMMCommand "READ?"
        Reply = ReceiveDMMReply()
        MeasureDigitalMultimeterAmps = Val(Reply)
        
End Function
Public Function MeasureDigitalMultimeterOhms() As Double
    Dim Reply As String
     
        SendDMMCommand ":FUNC 'RES'"
        SendDMMCommand ":SENS:RES:NPLC 1"
        SendDMMCommand ":SENS:RES:RANG:AUTO ON"
        SendDMMCommand "READ?"
        Reply = ReceiveDMMReply()
        MeasureDigitalMultimeterOhms = Val(Reply)
        
      
End Function
Public Function MeasureTemp() As Double
    Dim Reply As String
    Dim ohms As Double
    
        Sleep Relay_Delay
        SendDMMCommand ":FUNC 'RES'"
        SendDMMCommand ":SENS:RES:NPLC 0.1"
        SendDMMCommand ":SENS:RES:RANG 100"
        SendDMMCommand "READ?"
        Reply = ReceiveDMMReply()
        ohms = Val(Reply) - 1.5
        
        MeasureTemp = 20 + (ohms - 107.8) / 0.382
       
End Function
Public Function CheckLoadOn() As Boolean

    Dim reading As Double
    
        CheckLoadOn = False

' check gnd to vout1

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@102,105,135)"
        Sleep Relay_Delay
        reading = MeasureDigitalMultimeterOhms
        If reading < 100000 Then
        CheckLoadOn = True
        End If
        
'check VS to vout1

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,127,137)"
        Sleep Relay_Delay
        reading = MeasureDigitalMultimeterOhms
        If reading < 100000 Then
        CheckLoadOn = True
        End If
           
End Function
Public Function PSUCheck() ' 5V accuracy check
       
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@101,103,110,113)"
        Sleep Relay_Delay
   
End Function
Public Function SwitchToCheckPSUT()
       
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@135,118,121,233)"
        Sleep Relay_Delay
   
End Function
Public Function SwitchToCheckPSUL()
       
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@135,126,129,233)"
        Sleep Relay_Delay
    
End Function
Public Function SwitchToCheckPSUA()
       
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@143,118,121,233)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchToCheckPSUR()
       
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@143,126,129,233)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchPackOffsetMeas() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,118,137,143)"
        Sleep Relay_Delay
                
End Function
Public Function SwitchTempMeas() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@105,106,111,116)"
        Sleep Relay_Delay
                
End Function
Public Function SwitchCurrentMeas() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,118,129)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeasA() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,118,129)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeasQ() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,126,137)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeasM() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,134,129)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeasZ() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,134,129)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeasW() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,118,137)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeasL() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,126,121)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeasF() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,129,142)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeasC() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,142,129)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeas6() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,126,137)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeasN() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,126,145)"
        Sleep Relay_Delay
        
End Function
Public Function SwitchCurrentMeasU() As Double 'added by DW 17/11/23
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@109,112,134,121)"
        Sleep Relay_Delay
        
End Function

Public Function SwitchVout1Meas() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,135,121,118)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout1MeasA() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,143,121,118)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout1MeasQ() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@133,143,129,126)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout1MeasM() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,119,137,134)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout1MeasZ() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,119,137,134)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout1MeasW() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@133,127,121,118)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout1MeasL() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@117,135,129,126)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout1MeasF() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,135,145,142)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout1MeasC() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,119,145,142)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout1Meas6() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@133,119,129,126)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout1MeasN() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@141,135,129,126)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchDualVout1Meas() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@126,129,117,143)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchDualVout2Meas() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@126,129,117,135)"
        Sleep Relay_Delay
                 
End Function
Public Function SwitchVout2Meas() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,143,121,118)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2MeasA() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,135,121,118)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2MeasQ() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@133,119,129,126)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2MeasM() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,143,137,134)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2MeasZ() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,143,137,134)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2MeasW() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@133,143,121,118)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2MeasL() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@117,143,129,126)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2MeasF() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,119,145,142)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2MeasC() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,135,145,142)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2Meas6() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@133,143,129,126)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2MeasN() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@133,143,129,126)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2SwMeasT() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,143,121,118,233)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout1SwMeasA() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,143,121,118,233)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2SwMeasA() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@125,135,121,118,233)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2SwMeasL() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@117,143,129,126,233)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchVout2SwMeasR() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@117,135,129,126,233)"
        Sleep Relay_Delay
             
End Function
Public Function SwitchSTCVsMeas() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeasA() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeasQ() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeasM() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeasZ() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeasW() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeasL() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeasF() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeasC() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeas6() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeasN() As Double

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVsMeasU() As Double 'added by DW 17/11/23

        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function

Public Function SwitchSTCGndMeas() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeasA() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeasQ() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeasM() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeasZ() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeasW() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeasL() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeasF() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeasC() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeas6() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeasN() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCGndMeasU() As Double 'added by DW 17/11/23
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function

Public Function SwitchSTCVout1Meas() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1MeasA() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1MeasQ() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1MeasM() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1MeasZ() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1MeasW() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,127)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1MeasL() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1MeasF() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1MeasC() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1Meas6() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1MeasN() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout1MeasU() As Double 'added by DW 17/11/23
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function

Public Function SwitchSTCVout2Meas() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout2MeasA() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout2MeasQ() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout2MeasM() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout2MeasZ() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout2MeasW() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout2MeasL() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout2MeasF() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout2MeasC() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,135)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout2Meas6() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchSTCVout2MeasN() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,119)"
        Sleep Relay_Delay
               
End Function
Public Function SwitchPackInsulation() As Double
   
        OpenAllSwitches
        Sleep Relay_Delay
        SendDMMCommand ":ROUT:MULT:CLOS (@115,113,143)"
        Sleep Relay_Delay
               
End Function
Public Function OpenAllSwitches() As Double
   
        SendDMMCommand ":ROUT:OPEN:ALL"
              
End Function

