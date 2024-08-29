VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7395
   LinkTopic       =   "Form4"
   ScaleHeight     =   735
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
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
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   6750
      Picture         =   "Form4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ìì"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "ÊýÑ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "888"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "¾àÃ÷Äê»¹ÓÐ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
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
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim speed As Integer
Dim bg As String

Private Sub Command1_Click()
kbcout = 4
kbtext = Label3.Caption
kbcolr = Me.BackColor
Form16.Show
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width
Me.Top = Form2.Height * 2
Open "data\f4color.txt" For Input As #1
Line Input #1, bg
Close #1
Me.BackColor = bg
speed = 1
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Label3.Caption = allclass("c2")
allkbh = allkbh + Me.Height
Open "data\djstext.txt" For Input As #1
Line Input #1, djstext
Close #1
Open "data\djstime.txt" For Input As #1
Line Input #1, djstime
Close #1
Label1.Caption = "¾à" & djstext & "»¹ÓÐ"
Label2.Caption = CDate(djstime) - Date
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
 Form5.Show
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Open "data\djstext.txt" For Input As #1
Line Input #1, djstext
Close #1
Open "data\djstime.txt" For Input As #1
Line Input #1, djstime
Close #1
Label1.Caption = "¾à" & djstext & "»¹ÓÐ"
Label2.Caption = CDate(djstime) - Date
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

