VERSION 5.00
Begin VB.Form Outputs 
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Cycles 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   17
      Text            =   "10"
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Prog5Button 
      Caption         =   "Set Prog 5"
      Height          =   615
      Left            =   4200
      TabIndex        =   15
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Purge Marker"
      Height          =   615
      Left            =   1560
      TabIndex        =   14
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Prog4Button 
      Caption         =   "Set Prog 4"
      Height          =   615
      Left            =   4200
      TabIndex        =   13
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton Prog3Button 
      Caption         =   "Set Prog 3"
      Height          =   615
      Left            =   4200
      TabIndex        =   12
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Camera Feedback"
      Height          =   1935
      Left            =   5400
      TabIndex        =   5
      Top             =   840
      Width           =   3495
      Begin VB.Label PartPresentStatusLabel 
         Caption         =   "-------"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Part Present"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.Label RestrictorStatusLabel 
         Caption         =   "-------"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Restrictor"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label OringStatusLabel 
         Caption         =   "-------"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "O-Ring"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.CommandButton Prog2Button 
      Caption         =   "Set Prog 2"
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Prog1Button 
      Caption         =   "Set Prog 1"
      Height          =   615
      Left            =   4200
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton ResetCameraButton 
      Caption         =   "Reset Camera"
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton TrigggerMarkerButton 
      Caption         =   "Trigger Marker"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton TriggerCameraButton 
      Caption         =   "Trigger Camera"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Number Of Cycles Before Purge"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   4920
      Width           =   1455
   End
End
Attribute VB_Name = "Outputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()

    TriggerPurge
    
End Sub
Private Sub ResetCameraButton_Click()

    ResetCamera

End Sub
Private Sub TriggerCameraButton_Click()

    TriggerCamera
    
End Sub
Private Sub TrigggerMarkerButton_Click()

    TriggerMarker

End Sub
Private Sub Prog1Button_Click()

    PROG1
    
End Sub
Private Sub Prog2Button_Click()

    PROG2
    
End Sub
Private Sub Prog3Button_Click()

    PROG3
    
End Sub
Private Sub Prog4Button_Click()

    PROG4
    
End Sub
Private Sub Prog5Button_Click()

    PROG5
    
End Sub
