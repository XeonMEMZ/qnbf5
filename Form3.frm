VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7395
   LinkTopic       =   "Form3"
   ScaleHeight     =   735
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer6 
      Interval        =   1000
      Left            =   2520
      Top             =   120
   End
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
      Left            =   6750
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "语文"
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
      Left            =   4920
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "分钟"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "距下课还有"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim speed As Integer
Dim bg As String

Private Sub Command1_Click()
kbcout = 3
kbtext = Label4.Caption
kbcolr = Me.BackColor
Form16.Show
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width
Me.Top = Form2.Height
Open "data\f3color.txt" For Input As #1
Line Input #1, bg
Close #1
Me.BackColor = bg
speed = 1
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Label4.Caption = allclass("c1")
allkbh = allkbh + Me.Height
If fricls = 0 Then
 If Time < CDate(zxtime("zm1s")) Then
  If Hour(CDate(zxtime("zm1s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm1s")) <= Time And Time < CDate(zxtime("zm1x")) Then
  If Hour(CDate(zxtime("zm1x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm1x")) <= Time And Time < CDate(zxtime("zm2s")) Then
  If Hour(CDate(zxtime("zm2s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm2s")) <= Time And Time < CDate(zxtime("zm2x")) Then
  If Hour(CDate(zxtime("zm2x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm2x")) <= Time And Time < CDate(zxtime("zm3s")) Then
  If Hour(CDate(zxtime("zm3s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm3s")) <= Time And Time < CDate(zxtime("zm3x")) Then
  If Hour(CDate(zxtime("zm3x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm3x")) <= Time And Time < CDate(zxtime("zm4s")) Then
  If Hour(CDate(zxtime("zm4s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zm4x")) Then
  If Hour(CDate(zxtime("zm4x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4x")) - Time) + 1
  End If
 End If
 If Val(allclass("ct")) > 4 Then
  If CDate(zxtime("zm4x")) <= Time And Time < CDate(zxtime("zm5s")) Then
   If Hour(CDate(zxtime("zm5s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm5s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm5s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm5s")) <= Time And Time < CDate(zxtime("zm5x")) Then
   If Hour(CDate(zxtime("zm5x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm5x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm5x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 5 Then
  If CDate(zxtime("zm5x")) <= Time And Time < CDate(zxtime("zm6s")) Then
   If Hour(CDate(zxtime("zm6s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm6s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm6s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm6s")) <= Time And Time < CDate(zxtime("zm6x")) Then
   If Hour(CDate(zxtime("zm6x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm6x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm6x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 6 Then
  If CDate(zxtime("zm6x")) <= Time And Time < CDate(zxtime("zm7s")) Then
   If Hour(CDate(zxtime("zm7s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm7s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm7s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm7s")) <= Time And Time < CDate(zxtime("zm7x")) Then
   If Hour(CDate(zxtime("zm7x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm7x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm7x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 7 Then
  If CDate(zxtime("zm7x")) <= Time And Time < CDate(zxtime("zm8s")) Then
   If Hour(CDate(zxtime("zm8s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm8s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm8s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm8s")) <= Time And Time < CDate(zxtime("zm8x")) Then
   If Hour(CDate(zxtime("zm8x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm8x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm8x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 8 Then
  If CDate(zxtime("zm8x")) <= Time And Time < CDate(zxtime("zm9s")) Then
   If Hour(CDate(zxtime("zm9s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm9s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm9s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm9s")) <= Time And Time < CDate(zxtime("zm9x")) Then
   If Hour(CDate(zxtime("zm9x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm9x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm9x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 9 Then
  If CDate(zxtime("zm9x")) <= Time And Time < CDate(zxtime("zm10s")) Then
   If Hour(CDate(zxtime("zm10s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm10s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm10s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm10s")) <= Time And Time < CDate(zxtime("zm10x")) Then
   If Hour(CDate(zxtime("zm10x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm10x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm10x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 10 Then
  If CDate(zxtime("zm10x")) <= Time And Time < CDate(zxtime("zm11s")) Then
   If Hour(CDate(zxtime("zm11s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm11s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm11s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm11s")) <= Time And Time < CDate(zxtime("zm11x")) Then
   If Hour(CDate(zxtime("zm11x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm11x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm11x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 11 Then
  If CDate(zxtime("zm11x")) <= Time And Time < CDate(zxtime("zm12s")) Then
   If Hour(CDate(zxtime("zm12s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm12s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm12s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm12s")) <= Time And Time < CDate(zxtime("zm12x")) Then
   If Hour(CDate(zxtime("zm12x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm12x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm12x")) - Time) + 1
   End If
  End If
 End If
Else
  If Time < CDate(zxtime("zm1s")) Then
  If Hour(CDate(zxtime("zm1s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm1s")) <= Time And Time < CDate(zxtime("zm1x")) Then
  If Hour(CDate(zxtime("zm1x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm1x")) <= Time And Time < CDate(zxtime("zm2s")) Then
  If Hour(CDate(zxtime("zm2s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm2s")) <= Time And Time < CDate(zxtime("zm2x")) Then
  If Hour(CDate(zxtime("zm2x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm2x")) <= Time And Time < CDate(zxtime("zm3s")) Then
  If Hour(CDate(zxtime("zm3s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm3s")) <= Time And Time < CDate(zxtime("zm3x")) Then
  If Hour(CDate(zxtime("zm3x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm3x")) <= Time And Time < CDate(zxtime("zm4s")) Then
  If Hour(CDate(zxtime("zm4s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zm4x")) Then
  If Hour(CDate(zxtime("zm4x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4x")) - Time) + 1
  End If
 End If
 If Val(allclass("ct")) > 4 Then
  If CDate(zxtime("zm4x")) <= Time And Time < CDate(zxtime("zf5s")) Then
   If Hour(CDate(zxtime("zf5s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zf5s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zf5s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zf5s")) <= Time And Time < CDate(zxtime("zf5x")) Then
   If Hour(CDate(zxtime("zf5x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zf5x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zf5x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 5 Then
  If CDate(zxtime("zf5x")) <= Time And Time < CDate(zxtime("zf6s")) Then
   If Hour(CDate(zxtime("zf6s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zf6s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zf6s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zf6s")) <= Time And Time < CDate(zxtime("zf6x")) Then
   If Hour(CDate(zxtime("zf6x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zf6x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zf6x")) - Time) + 1
   End If
  End If
 End If
End If
End Sub

Private Sub Label4_DblClick()
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
 Form4.Show
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

Public Function zxtime(s$) As String
Dim timelist As String
Open "data\timelist.txt" For Input As #1
Line Input #1, timelist
Close #1
If Len(s) = 4 Then
 zxtime = Trim(Mid(timelist, InStr(timelist, CStr(s)) + 4, 9))
Else
 zxtime = Trim(Mid(timelist, InStr(timelist, CStr(s)) + 5, 9))
End If
End Function

Private Sub Timer6_Timer()
If fricls = 0 Then
 If Time < CDate(zxtime("zm1s")) Then
  If Hour(CDate(zxtime("zm1s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm1s")) <= Time And Time < CDate(zxtime("zm1x")) Then
  If Hour(CDate(zxtime("zm1x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm1x")) <= Time And Time < CDate(zxtime("zm2s")) Then
  If Hour(CDate(zxtime("zm2s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm2s")) <= Time And Time < CDate(zxtime("zm2x")) Then
  If Hour(CDate(zxtime("zm2x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm2x")) <= Time And Time < CDate(zxtime("zm3s")) Then
  If Hour(CDate(zxtime("zm3s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm3s")) <= Time And Time < CDate(zxtime("zm3x")) Then
  If Hour(CDate(zxtime("zm3x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm3x")) <= Time And Time < CDate(zxtime("zm4s")) Then
  If Hour(CDate(zxtime("zm4s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zm4x")) Then
  If Hour(CDate(zxtime("zm4x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4x")) - Time) + 1
  End If
 End If
 If Val(allclass("ct")) > 4 Then
  If CDate(zxtime("zm4x")) <= Time And Time < CDate(zxtime("zm5s")) Then
   If Hour(CDate(zxtime("zm5s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm5s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm5s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm5s")) <= Time And Time < CDate(zxtime("zm5x")) Then
   If Hour(CDate(zxtime("zm5x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm5x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm5x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 5 Then
  If CDate(zxtime("zm5x")) <= Time And Time < CDate(zxtime("zm6s")) Then
   If Hour(CDate(zxtime("zm6s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm6s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm6s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm6s")) <= Time And Time < CDate(zxtime("zm6x")) Then
   If Hour(CDate(zxtime("zm6x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm6x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm6x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 6 Then
  If CDate(zxtime("zm6x")) <= Time And Time < CDate(zxtime("zm7s")) Then
   If Hour(CDate(zxtime("zm7s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm7s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm7s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm7s")) <= Time And Time < CDate(zxtime("zm7x")) Then
   If Hour(CDate(zxtime("zm7x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm7x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm7x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 7 Then
  If CDate(zxtime("zm7x")) <= Time And Time < CDate(zxtime("zm8s")) Then
   If Hour(CDate(zxtime("zm8s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm8s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm8s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm8s")) <= Time And Time < CDate(zxtime("zm8x")) Then
   If Hour(CDate(zxtime("zm8x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm8x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm8x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 8 Then
  If CDate(zxtime("zm8x")) <= Time And Time < CDate(zxtime("zm9s")) Then
   If Hour(CDate(zxtime("zm9s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm9s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm9s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm9s")) <= Time And Time < CDate(zxtime("zm9x")) Then
   If Hour(CDate(zxtime("zm9x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm9x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm9x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 9 Then
  If CDate(zxtime("zm9x")) <= Time And Time < CDate(zxtime("zm10s")) Then
   If Hour(CDate(zxtime("zm10s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm10s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm10s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm10s")) <= Time And Time < CDate(zxtime("zm10x")) Then
   If Hour(CDate(zxtime("zm10x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm10x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm10x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 10 Then
  If CDate(zxtime("zm10x")) <= Time And Time < CDate(zxtime("zm11s")) Then
   If Hour(CDate(zxtime("zm11s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm11s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm11s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm11s")) <= Time And Time < CDate(zxtime("zm11x")) Then
   If Hour(CDate(zxtime("zm11x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm11x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm11x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 11 Then
  If CDate(zxtime("zm11x")) <= Time And Time < CDate(zxtime("zm12s")) Then
   If Hour(CDate(zxtime("zm12s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm12s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zm12s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zm12s")) <= Time And Time < CDate(zxtime("zm12x")) Then
   If Hour(CDate(zxtime("zm12x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm12x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zm12x")) - Time) + 1
   End If
  End If
 End If
Else
  If Time < CDate(zxtime("zm1s")) Then
  If Hour(CDate(zxtime("zm1s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm1s")) <= Time And Time < CDate(zxtime("zm1x")) Then
  If Hour(CDate(zxtime("zm1x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm1x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm1x")) <= Time And Time < CDate(zxtime("zm2s")) Then
  If Hour(CDate(zxtime("zm2s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm2s")) <= Time And Time < CDate(zxtime("zm2x")) Then
  If Hour(CDate(zxtime("zm2x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm2x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm2x")) <= Time And Time < CDate(zxtime("zm3s")) Then
  If Hour(CDate(zxtime("zm3s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm3s")) <= Time And Time < CDate(zxtime("zm3x")) Then
  If Hour(CDate(zxtime("zm3x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm3x")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm3x")) <= Time And Time < CDate(zxtime("zm4s")) Then
  If Hour(CDate(zxtime("zm4s")) - Time) > 0 Then
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4s")) - Time) + 61
  Else
   Label1.Caption = "距上课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4s")) - Time) + 1
  End If
 End If
 If CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zm4x")) Then
  If Hour(CDate(zxtime("zm4x")) - Time) > 0 Then
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4x")) - Time) + 61
  Else
   Label1.Caption = "距下课还有"
   Label2.Caption = Minute(CDate(zxtime("zm4x")) - Time) + 1
  End If
 End If
 If Val(allclass("ct")) > 4 Then
  If CDate(zxtime("zm4x")) <= Time And Time < CDate(zxtime("zf5s")) Then
   If Hour(CDate(zxtime("zf5s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zf5s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zf5s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zf5s")) <= Time And Time < CDate(zxtime("zf5x")) Then
   If Hour(CDate(zxtime("zf5x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zf5x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zf5x")) - Time) + 1
   End If
  End If
 End If
 If Val(allclass("ct")) > 5 Then
  If CDate(zxtime("zf5x")) <= Time And Time < CDate(zxtime("zf6s")) Then
   If Hour(CDate(zxtime("zf6s")) - Time) > 0 Then
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zf6s")) - Time) + 61
   Else
    Label1.Caption = "距上课还有"
    Label2.Caption = Minute(CDate(zxtime("zf6s")) - Time) + 1
   End If
  End If
  If CDate(zxtime("zf6s")) <= Time And Time < CDate(zxtime("zf6x")) Then
   If Hour(CDate(zxtime("zf6x")) - Time) > 0 Then
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zf6x")) - Time) + 61
   Else
    Label1.Caption = "距下课还有"
    Label2.Caption = Minute(CDate(zxtime("zf6x")) - Time) + 1
   End If
  End If
 End If
End If
End Sub

Public Function allclass(s$) As String
Dim classlist As String
Open "data\z" & Weekday(Date, 2) & ".txt" For Input As #1
Line Input #1, classlist
Close #1
allclass = Trim(Mid(classlist, InStr(classlist, CStr(s)) + 2, 3))
End Function
