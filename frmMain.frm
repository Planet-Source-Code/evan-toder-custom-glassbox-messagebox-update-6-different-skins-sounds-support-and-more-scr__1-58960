VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin CustomMessageBox.GlassBox GlassBox1 
      Left            =   3465
      Top             =   2025
      _extentx        =   423
      _extenty        =   423
      picture         =   "frmMain.frx":0000
      pictureoffsety  =   5
   End
   Begin VB.OptionButton Option1 
      Caption         =   "glass yellow"
      Height          =   285
      Index           =   5
      Left            =   225
      TabIndex        =   7
      Top             =   2880
      Width           =   2265
   End
   Begin VB.OptionButton Option1 
      Caption         =   "glass white"
      Height          =   285
      Index           =   4
      Left            =   225
      TabIndex        =   6
      Top             =   2610
      Width           =   2265
   End
   Begin VB.OptionButton Option1 
      Caption         =   "glass blue"
      Height          =   285
      Index           =   3
      Left            =   225
      TabIndex        =   5
      Top             =   2340
      Width           =   2265
   End
   Begin VB.OptionButton Option1 
      Caption         =   "glass orange"
      Height          =   285
      Index           =   2
      Left            =   225
      TabIndex        =   4
      Top             =   2070
      Width           =   2265
   End
   Begin VB.OptionButton Option1 
      Caption         =   "glass candy"
      Height          =   285
      Index           =   1
      Left            =   225
      TabIndex        =   3
      Top             =   1800
      Width           =   2265
   End
   Begin VB.OptionButton Option1 
      Caption         =   "glass red"
      Height          =   285
      Index           =   0
      Left            =   225
      TabIndex        =   2
      Top             =   1530
      Width           =   2265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Test"
      Height          =   285
      Left            =   225
      TabIndex        =   1
      Top             =   990
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "You can draw any picture on the custom message box. any size. AND you can specify a color that is transparent or not painted"
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   315
      TabIndex        =   0
      Top             =   45
      Width           =   4065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim optionClicker    As Long

Private Sub Command1_Click()
 
 Dim strMsg  As String
 Dim retVal  As Long
 
 strMsg = "This custom messagebox  has the powerful features " & _
          "of a regular messagebox" & vbCrLf & _
          "such as autosizing to fit the message, AND..." & vbCrLf & _
           Space(20) & "+ the visual superiority is being displayed now" & vbCrLf & _
           Space(20) & "+ you must supply the ""vbCrLf"" which means total control" & vbCrLf & _
           Space(20) & "+ you can display any type or size picture you wish"
           
 'move picture down 30 pixels from default
 GlassBox1.pictureOffsetY = 30
  
' If GlassBox1.Message( _
'    strMsg, "this is the caption", YesNo, sndCustom, _
'    App.Path & "\testing.wav") = vbYes Then
 If GlassBox1.Message( _
    strMsg, "this is the caption", YesNo, sndExclamation _
    ) = vbYes Then
     Debug.Print "YES!!!"
 Else
     Debug.Print "NO!!!"
 End If

End Sub

'Private Sub Form_Unload(Cancel As Integer)
'  timerDisplay.Enabled = False
'End Sub

Private Sub Option1_Click(Index As Integer)
 GlassBox1.GlassBoxColor = Index
End Sub
'
'Private Sub timerDisplay_Timer()
'   optionClicker = optionClicker + 1
'   Option1_Click (optionClicker)
'   If optionClicker >= 5 Then optionClicker = 0
'End Sub
