VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Your wish is my command"
   ClientHeight    =   1650
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3720
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Timer tmrTime 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   3120
         Top             =   1200
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
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
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   495
         Left            =   120
         Top             =   960
         Width           =   1095
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
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         Height          =   495
         Left            =   1200
         Top             =   960
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
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF0000&
         Height          =   1215
         Left            =   2280
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblOK 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuRec 
         Caption         =   "&Record New Commands"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuCommands 
         Caption         =   "&Commands Management"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu systray 
      Caption         =   "&Move In The Systray "
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                         *** Command Your Computer ***
'
'Author:      Licar Bogdan (copyright).
'
'Description: After recording some commands and assigning them an exe file, you just
'             have to repeat the command whenever you want to run the exe. It bases
'             on my previous program, Voice Recognition. It works when the commands are
'             pronounced almost in the same way. Error ranges may be varied, in order
'             to obtain a more satisfying result. I think these error ranges give the
'             right result (with exceptions, of course). It has the advantage that
'             no external ocx's and dll's are used.
'
'Purpose:     Again, I got this idea from movies, where one had to say a predefined
'             command to run a program, or to make the computer do something that
'             would take to much time by normal ways (i.e. clicking and clicking).
'             Tyler says: The Importance of Being Idle.
'             Tyler says: HAL, open the pod doors.
'
'NOTE:        My advice is, if you define many commands, to record them different
'             from eachother. I mean, if there are 2 commands: "Go to doom.exe" and
'             "Go to room.exe", the program won't distinguish them and will make confusion.
'             So, record sound of different length and pronounced in different ways.

Option Explicit

Dim Tm As Date, Recording As Boolean, TmpPath As String

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CloseSysTray

'Set the sound free and delete the temporal recorded one.
ResetWave
Kill TmpPath
End
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRec.ForeColor = &HFF&
lblStop.ForeColor = &HFFFF&
lblOK.ForeColor = &HFFFF&
End Sub

Private Sub lblOK_Click()
Dim str As String, str2() As String, ShortSC As String
On Error Resume Next

'Search a wave with the same characteristics as of the one recorded
ShortSC = GetShortName(SearchCommand(TmpPath))
If ShortSC <> "" Then
    
    'Search for the path that the program has to run in the ini file
    Open (CommandsFld & "Commands.ini") For Input As #1
        Do While EOF(1) = False
            Input #1, str
            str2 = Split(str, "=")

            If (Left$(str2(1), Len(str2(1)) - 6) = Left$(ShortSC, Len(ShortSC) - 6)) Then
                If str <> "" Then
                    Shell str2(0), vbNormalFocus
                    Close #1
                    Exit Sub
                End If
            End If
            
        Loop
    Close #1
Else
    MsgBox "No matching command was found.", vbInformation
End If
End Sub

Private Sub ResetLabel()
lblTime.Caption = CDate(Time - Time)
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
lblOK.ForeColor = &HFFFF&
On Error Resume Next

    Select Case X
    
        Case 7695:      'Leftclick
        frmMain.Show
        lblRec_Click

        Case 7740:      'Rightclick
        PopupMenu mnuSettings

    End Select
End Sub

Private Sub lblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOK.ForeColor = &H80FF&
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

Private Sub lblStop_Click()
StopWave
SaveWave TmpPath
ResetLabel: tmrTime.Enabled = False
Recording = False: Me.SetFocus
End Sub

Private Sub Form_Load()
ResetLabel
GetWaves CommandsFld, Waves             'Get all defined commands

TmpPath = App.Path & "\Tmp.wav"         'Constants
CommandsFld = App.Path & "\Commands\"
End Sub

Private Sub lblStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStop.ForeColor = &H80FF&
End Sub

Private Sub mnuCommands_Click()
frmCommands.Show vbNormalFocus
End Sub

Private Sub mnuQuit_Click()
Unload Me
End
End Sub

Private Sub mnuRec_Click()
frmRecord.Show vbNormalFocus
End Sub

Private Sub systray_Click()
Me.Hide
ShowInSysTray
End Sub

Private Sub tmrTime_Timer()
lblTime.Caption = CDate(Time - Tm)
End Sub
