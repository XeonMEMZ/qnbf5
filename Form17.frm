VERSION 5.00
Begin VB.Form Form17 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form17"
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   LinkTopic       =   "Form17"
   ScaleHeight     =   1470
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ò»ÖÜ¿Î±í"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "¿¼ÊÔÄ£Ê½"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ëø¶¨ÆÁÄ»"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ëæ»úµãÃû"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   2040
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   1830
      Picture         =   "Form17.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Ð¡¹¤¾ß"
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim speed As Integer
Dim sj As Integer

Private Sub Command1_Click()
If Timer4.Enabled = True Then
 Timer4.Enabled = False
Else
 Timer5.Enabled = True
End If
End Sub

Private Sub Command2_Click()
If fdczt = False Then
 Open "data\namec.txt" For Input As #1
 Line Input #1, cname
 Close #1
 Randomize
 sj = Int(Rnd * (CInt(cname) - 1 + 1) + 1)
 fdctime = 1
 fdctext = "Ëæ»úµãÃû:" & namelist(sj)
 Form18.Show
End If
End Sub

Private Sub Command3_Click()
Unload Form19
ksmsmode = False
Form19.Show
End Sub

Private Sub Command4_Click()
Unload Form19
ksmsmode = True
Form19.Show
End Sub

Private Sub Command5_Click()
Form20.Show
End Sub

Private Sub Form_Load()
Me.Height = 735
Me.Left = Screen.Width
Me.Top = allkbh
speed = 1
allkbh = allkbh + Me.Height
Unload Form1
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
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If speed > 0 Then
 If Me.Left - speed >= Screen.Width - Me.Width + (780 / 3) Then
  Me.Left = Me.Left - speed
  Me.Height = Me.Height + speed
  speed = speed + 2
 Else
  Me.Left = Me.Left - speed
  Me.Height = Me.Height + speed - 2
  speed = speed - 4
 End If
Else
 speed = 1
 Me.Left = Screen.Width - Me.Width
 Me.Height = 1470
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
  Me.Height = Me.Height - speed
  speed = speed + 2
 Else
  Me.Left = Me.Left + speed
  Me.Height = Me.Height - speed + 2
  speed = speed - 1.5
 End If
Else
 speed = 1
 Me.Left = Screen.Width - Me.Width + 780
 Me.Height = 735
 Timer5.Enabled = False
End If
End Sub

Public Function namelist(n%) As String
Dim name As String
Open "data\name.txt" For Input As #1
Line Input #1, name
Close #1
If n < 10 Then
 namelist = Trim(Mid(name, InStr(name, CStr(n)) + 1, 4))
ElseIf n >= 10 Then
 namelist = Trim(Mid(name, InStr(name, CStr(n)) + 2, 4))
End If
End Function
