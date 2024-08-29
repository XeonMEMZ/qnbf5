VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Form11"
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   LinkTopic       =   "Form11"
   ScaleHeight     =   735
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   2
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer Timer5 
      Interval        =   2
      Left            =   2040
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   1830
      Picture         =   "Form11.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "地理"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim speed As Integer
Dim bg As String

Private Sub Command1_Click()
kbcout = 11
kbtext = Label1.Caption
kbcolr = Me.BackColor
Form16.Show
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width
Me.Top = Form2.Height * 9
Open "data\f11color.txt" For Input As #1
Line Input #1, bg
Close #1
Me.BackColor = bg
speed = 1
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Label1.Caption = allclass("c9")
allkbh = allkbh + Me.Height
End Sub

Private Sub Label1_DblClick()
Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()
If speed > 0 Then
 If Me.Left - speed >= (Screen.Width - Me.Width + 780) + (Me.Width / 3) Then
  Me.Left = Me.Left - speed
  speed = speed + 5
 Else
  Me.Left = Me.Left - speed
  speed = speed - 5
 End If
Else
 speed = 1
 Me.Left = Screen.Width - Me.Width + 780
 If Val(allclass("ct")) > 9 Then
  Form12.Show
 Else
  Form17.Show
 End If
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If speed > 0 Then
 If Me.Left - speed >= Screen.Width - Me.Width + (780 / 3) Then
  Me.Left = Me.Left - speed
  speed = speed + 2
 Else
  Me.Left = Me.Left - speed
  speed = speed - 4
 End If
Else
 speed = 1
 Me.Left = Screen.Width - Me.Width
 Timer4.Enabled = True
 Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
Timer5.Enabled = True
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
If speed > 0 Then
 If Me.Left - speed <= Screen.Width - Me.Width + (780 / 2.5) Then
  Me.Left = Me.Left + speed
  speed = speed + 2
 Else
  Me.Left = Me.Left + speed
  speed = speed - 1.5
 End If
Else
 speed = 1
 Me.Left = Screen.Width - Me.Width + 780
 Timer5.Enabled = False
End If
End Sub

Public Function allclass(s$) As String
Dim classlist As String
Open "data\z" & Weekday(Date, 2) & ".txt" For Input As #1
Line Input #1, classlist
Close #1
allclass = Trim(Mid(classlist, InStr(classlist, CStr(s)) + 2, 3))
End Function

