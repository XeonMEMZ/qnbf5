VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":6988A
   ScaleHeight     =   2205
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer3 
      Interval        =   2
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   2
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   120
      Top             =   120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A2D581&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   -9100
      Top             =   2100
      Width           =   9100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1
Dim qdtm As Integer
Dim fxs As String

Private Sub Form_Load()
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 0, LWA_ALPHA
fdczt = False
Timer3.Enabled = False
Open "data\fx.txt" For Input As #1
Line Input #1, fxs
Close #1
If fxs = "5" Then
 If Weekday(Date, 2) = 5 Then
  fricls = 1
 Else
  fricls = 0
 End If
ElseIf fxs = "6" Then
 If Weekday(Date, 2) = 6 Then
  fricls = 1
 Else
  fricls = 0
 End If
ElseIf fxs = "7" Then
 If Weekday(Date, 2) = 7 Then
  fricls = 1
 Else
  fricls = 0
 End If
End If
Form2.Show
End Sub

Private Sub Timer1_Timer()
If Shape1.Left + 90 > 0 Then
 Timer3.Enabled = True
Else
 Shape1.Left = Shape1.Left + 90
End If
End Sub

Private Sub Timer2_Timer()
If qdtm + 5 <= 255 Then
 qdtm = qdtm + 5
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), qdtm, LWA_ALPHA
Else
 qdtm = 255
 Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If qdtm - 5 >= 0 Then
 qdtm = qdtm - 5
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), qdtm, LWA_ALPHA
Else
 qdtm = 0
 Unload Me
 Timer3.Enabled = False
End If
End Sub
