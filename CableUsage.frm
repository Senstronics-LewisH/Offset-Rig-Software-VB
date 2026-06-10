VERSION 5.00
Begin VB.Form Usage 
   Caption         =   "Cable Usage Reset"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox pbox2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      IMEMode         =   3  'DISABLE
      Left            =   3120
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3600
      Width           =   6135
   End
   Begin VB.CommandButton Cmdpword2 
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   0
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "TEAM LEADER - PLEASE ENTER PASSWORD BELOW AND PRESS ENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      TabIndex        =   3
      Top             =   2160
      Width           =   10215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "PLEASE SEE YOUR TEAM LEADER FOR A REPLACEMENT CABLE AND COUNTER RESET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   10935
   End
End
Attribute VB_Name = "Usage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdpword2_Click()
Const PassWord As String = "sens666"
         
    If Usage.pbox2.Text = PassWord Then
     CableUsage(i) = 0
     MainForm.NumberOfUsesDisplay = CableUsage(i)
     Usage.pbox2.Text = ""
     Usage.Hide
     MainForm.Show
    End If
End Sub


