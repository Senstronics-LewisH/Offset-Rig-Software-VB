Attribute VB_Name = "InputsOutputs"
Option Explicit

Global card As Integer
Public IsInterlockClosed As Boolean
Public PartPresent As Boolean
Public OringPresent As Boolean
Public RestrictorPresent As Boolean
Public OutputSum As Double
Public Sub Initialise7250()
  
  Register_Card PCI_7250, 0

End Sub
Public Sub CheckInputs()
' modified by DW for toggle switch operation

  Dim Result As Integer
  Dim di_data As Long
  Dim i As Integer
  Dim p As Integer
  Dim Port As Integer
  
    Port = 0
    card = 0
        
    If IOCard = False Then
        Exit Sub
    Else
        
        Result = DI_ReadPort(card, Port, di_data)
        For i = 0 To 7
        p = di_data Mod 2
        If i = 0 Then
           
        If p = 0 Then
           ' MainForm.InterlockStatusLabel = " Open"
            IsInterlockClosed = False
            ReadyFlag = True
        Else 'p = 1
           ' MainForm.InterlockStatusLabel = " Closed"
            IsInterlockClosed = True
                If ReadyFlag = True And DetailChecked = True Then
                    ReadyFlag = False
                    START
                   ElseIf i = 1 Then
               End If
            End If
           
          End If

            di_data = CInt(Int(di_data / 2))
        Next
        
    End If

End Sub
