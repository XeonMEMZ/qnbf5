VERSION 5.00
Begin VB.Form Form19 
   BackColor       =   &H004D5E2B&
   BorderStyle     =   0  'None
   Caption         =   "Form19"
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9720
   LinkTopic       =   "Form19"
   Picture         =   "Form19.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "¹Ø±Õ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   4
      Top             =   5640
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
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
      Height          =   765
      Left            =   1200
      TabIndex        =   3
      Text            =   "¿ÆÄ¿"
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
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
      Height          =   765
      Left            =   3720
      TabIndex        =   2
      Text            =   "Ê±¼ä"
      Top             =   3840
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
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
      Height          =   765
      Left            =   1200
      TabIndex        =   1
      Text            =   "¿ÆÄ¿"
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
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
      Height          =   765
      Left            =   3720
      TabIndex        =   0
      Text            =   "Ê±¼ä"
      Top             =   4680
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   99.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2265
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   8055
   End
End
Attribute VB_Name = "Form19"
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
Dim tmd As Integer

Private Sub Command1_Click()
allkbh = 0
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
Unload Form9
Unload Form10
Unload Form11
Unload Form12
Unload Form13
Unload Form14
Unload Form17
Me.Left = 0
Me.Top = 0
Me.Width = Screen.Width
Me.Height = Screen.Height
Command1.Left = Me.Width - 1535
Command1.Top = Me.Height - 1535
Label1.Left = (Screen.Width - Label1.Width) / 2
Label1.Top = (Screen.Height - Label1.Height) / 2 - 1500
Text1.Left = Label1.Left + 360
Text1.Top = Label1.Top + 2520
Text2.Left = Label1.Left + 2880
Text2.Top = Label1.Top + 2520
Text3.Left = Label1.Left + 360
Text3.Top = Label1.Top + 3360
Text4.Left = Label1.Left + 2880
Text4.Top = Label1.Top + 3360
If Len(Time) = 7 Then
 Label1.Caption = "0" & Time
Else
 Label1.Caption = Time
End If
If ksmsmode = True Then
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
Else
 Me.Picture = LoadPicture()
 Me.BackColor = RGB(255, 255, 255)
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 50, LWA_ALPHA
 Label1.Visible = False
 Text1.Visible = False
 Text2.Visible = False
 Text3.Visible = False
 Text4.Visible = False
End If
End Sub

Private Sub Timer1_Timer()
If Len(Time) = 7 Then
 Label1.Caption = "0" & Time
Else
 Label1.Caption = Time
End If
End Sub
