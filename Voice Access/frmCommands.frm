VERSION 5.00
Begin VB.Form frmCommands 
   BackColor       =   &H80000007&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Commands Management"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4050
   Icon            =   "frmCommands.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   3960
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H00FFFFFF&
      Height          =   3930
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4095
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   4560
      Width           =   135
   End
   Begin VB.CommandButton cmdDel 
      Appearance      =   0  'Flat
      Caption         =   "Delete"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   2055
   End
End
Attribute VB_Name = "frmCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ini As String, i As Integer

Private Sub cmdDel_Click()
Dim Path1 As String, ind As Integer
    'Delete a defined command
    ind = List1.ListIndex
    Path1 = List1.List(ind) & "=" & List2.List(ind)
    
    Ini = Replace(Ini, Path1, "")
    Open CommandsFld & "Commands.ini" For Output As #1
        Print #1, Ini
    Close #1
    List1.RemoveItem (ind)
    
    ResetWave
    Kill List2.List(ind)
    List2.RemoveItem (ind)
End Sub

Private Sub cmdPlay_Click()
sndPlaySound List2.List(List1.ListIndex), 1
End Sub

Private Sub Form_Load()
Dim str As String, str2() As String, PrecStr() As String, i As Integer, count As Integer
ReDim str2(1): ReDim PrecStr(1)

'Get all predefined commands
Open CommandsFld & "Commands.ini" For Input As #1
Do While EOF(1) = False
    Input #1, str
    If str <> "" Then
        
        For i = 1 To UBound(PrecStr)
            If str = PrecStr(i) Then GoTo Nxt
        Next i
        
        count = count + 1
        ReDim Preserve PrecStr(count)
        PrecStr(count) = str
        
        Ini = Ini & vbCrLf & str
        str2 = Split(str, "=")
        List1.AddItem str2(0)
        List2.AddItem str2(1)
    End If
    
Nxt:
Loop
Close #1

End Sub

Private Sub Form_Resize()
On Error Resume Next
    List1.Width = Me.Width
    List1.Height = Me.Height - 820
    cmdDel.Top = Me.Height - 840
    cmdPlay.Top = cmdDel.Top
    
    cmdPlay.Width = Me.Width / 2
    cmdDel.Width = cmdPlay.Width
    cmdPlay.Left = cmdDel.Width
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ind As Integer
    ind = Y / 240
    List1.ToolTipText = List1.List(ind)
End Sub
