VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7395
   LinkTopic       =   "Form2"
   ScaleHeight     =   735
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer6 
      Interval        =   500
      Left            =   2520
      Top             =   120
   End
   Begin VB.Timer Timer5 
      Interval        =   2
      Left            =   2040
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   2
      Left            =   1080
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   6750
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "星期日"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "2024/12/12"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim speed As Integer
Dim dskop As Integer
Dim disk As Integer

Private Sub Command1_Click()
Form15.Show
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width
Me.Top = 0
speed = 1
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
allkbh = allkbh + Me.Height
disk = Drive1.ListCount
dskop = 0
Label1.Caption = Date
If Len(Time) = 7 Then
 Label2.Caption = "0" & Time
Else
 Label2.Caption = Time
End If
If Weekday(Date, 2) = 1 Then
 Label3.Caption = "星期一"
ElseIf Weekday(Date, 2) = 2 Then
 Label3.Caption = "星期二"
ElseIf Weekday(Date, 2) = 3 Then
 Label3.Caption = "星期三"
ElseIf Weekday(Date, 2) = 4 Then
 Label3.Caption = "星期四"
ElseIf Weekday(Date, 2) = 5 Then
 Label3.Caption = "星期五"
ElseIf Weekday(Date, 2) = 6 Then
 Label3.Caption = "星期六"
ElseIf Weekday(Date, 2) = 7 Then
 Label3.Caption = "星期日"
End If
End Sub

Private Sub Label3_DblClick()
Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()
If speed > 0 Then
 If Me.Left - speed >= (Screen.Width - Me.Width + 780) + (Me.Width / 3) Then
  Me.Left = Me.Left - speed
  speed = speed + 5
 Else
  Me.Left = Me.Left - speed
  speed = speed - 8.7
 End If
Else
 speed = 1
 Me.Left = Screen.Width - Me.Width + 780
 Form3.Show
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Drive1.Refresh
If Drive1.ListCount > disk Then
 If dskop = 0 Then
  Dim upf As String
  Open "data\upf.txt" For Input As #1
  Line Input #1, upf
  Close #1
  Call Shell("cmd /c start " & upf)
  fdctime = 2
  fdctext = "正在为您打开U盘"
  Form18.Show
  dskop = 1
 End If
Else
 dskop = 0
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

Private Sub Timer6_Timer()
Label1.Caption = Date
If Len(Time) = 7 Then
 Label2.Caption = "0" & Time
Else
 Label2.Caption = Time
End If
If Weekday(Date, 2) = 1 Then
 Label3.Caption = "星期一"
ElseIf Weekday(Date, 2) = 2 Then
 Label3.Caption = "星期二"
ElseIf Weekday(Date, 2) = 3 Then
 Label3.Caption = "星期三"
ElseIf Weekday(Date, 2) = 4 Then
 Label3.Caption = "星期四"
ElseIf Weekday(Date, 2) = 5 Then
 Label3.Caption = "星期五"
ElseIf Weekday(Date, 2) = 6 Then
 Label3.Caption = "星期六"
ElseIf Weekday(Date, 2) = 7 Then
 Label3.Caption = "星期日"
End If
End Sub
