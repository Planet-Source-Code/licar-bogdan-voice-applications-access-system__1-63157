VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRecord 
   BackColor       =   &H80000012&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Define commands"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5760
   Icon            =   "frmRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.Timer tmrTime 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   3720
         Top             =   240
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000007&
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   3720
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FF0000&
         Height          =   495
         Left            =   4320
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblDone 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Done"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "System"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "Record your commands and choose a path Voice Access to run."
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblRec 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Record"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   495
         Left            =   1920
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblAdd 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   4440
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         Height          =   1215
         Left            =   4320
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblStop 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   1680
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF0000&
         Height          =   495
         Left            =   3120
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF0000&
         Height          =   495
         Left            =   480
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblBrowse 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   1680
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tm As Date, Recording As Boolean

Private Sub ResetLabel()
lblTime.Caption = CDate(Time - Time)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRec.ForeColor = &HFF&
lblStop.ForeColor = &HFFFF&
lblAdd.ForeColor = &HFFFF&
lblBrowse.ForeColor = &HFFFF&
lblDone.ForeColor = &HFFFF&
End Sub

Private Sub lblBrowse_Click()
CD1.Filter = "Executables (*.exe)|*.exe"
CD1.ShowOpen
Text1.Text = CD1.filename
End Sub

Private Sub lblBrowse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBrowse.ForeColor = &H80FF&
End Sub

Private Sub lblDone_Click()
Unload Me
End Sub

Private Sub lblDone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDone.ForeColor = &H80FF&
End Sub

Private Sub lblRec_Click()
RecordWave
ResetLabel
Tm = Time: tmrTime.Enabled = True
Recording = True: Me.SetFocus
End Sub

Private Sub lblRec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRec.ForeColor = &H80&
End Sub

Private Sub lblAdd_Click()
Dim i As Integer, str2 As String, str As String
    
    'Save the recorded command and write it in the ini file
    str2 = GetFileName(Text1.Text, False)
    SaveWave CommandsFld & str2 & ".wav"
    
    Open (CommandsFld & "Commands.ini") For Input As #1
        str = Input(LOF(1), 1)
    Close #1
    
    str2 = GetShortName(CommandsFld & str2 & ".wav")
    If Len(GetFileName(Text1, False)) > 6 Then str2 = Left$(str2, Len(str2) - 5) & "1.wav"
    Open (CommandsFld & "Commands.ini") For Output As #1
        Print #1, str
        Print #1, Text1 & "=" & str2
    Close #1
    
    MsgBox "A command to the following application was added: " & Text1, vbInformation
    Text1 = ""
    
    ResetLabel
End Sub

Private Sub lblAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.ForeColor = &H80FF&
End Sub

Private Sub lblStop_Click()
StopWave
ResetLabel: tmrTime.Enabled = False
Recording = False: Me.SetFocus
End Sub

Private Sub Form_Load()
ResetLabel
End Sub

Private Sub lblStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStop.ForeColor = &H80FF&
End Sub

Private Sub tmrTime_Timer()
lblTime.Caption = CDate(Time - Tm)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If Recording = False And KeyAscii = 13 Then
    lblRec_Click
ElseIf Recording = True And KeyAscii = 13 Then
    lblStop_Click
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRec.ForeColor = &HFF&
lblStop.ForeColor = &HFFFF&
lblAdd.ForeColor = &HFFFF&
lblBrowse.ForeColor = &HFFFF&
End Sub

