VERSION 5.00
Begin VB.Form Form15 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Èí¼þÉèÖÃ"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "¹Ø±ÕÈí¼þ"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "È·¶¨"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0FF&
      Caption         =   "¹ØÓÚÈí¼þ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÐÞ¸ÄÃûµ¥"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ÐÞ¸ÄUÅÌÅÌ·û"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ÐÞ¸Äµ¹¼ÆÊ±"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ÐÞ¸Ä×÷Ï¢Ê±¼ä"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ÐÞ¸ÄÒ»ÖÜ¿Î±í"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim speed As Integer

Private Sub Command1_Click()
Form22.Show
End Sub

Private Sub Command2_Click()
Form23.Show
End Sub

Private Sub Command3_Click()
Form24.Show
End Sub

Private Sub Command4_Click()
Form25.Show
End Sub

Private Sub Command5_Click()
Call Shell("cmd /c start data\name.txt")
Call Shell("cmd /c start data\namec.txt")
End Sub

Private Sub Command6_Click()
Form21.Show
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Not Form2.Left > Screen.Width + 100 Then
 Form2.Left = Form2.Left + speed
 Form3.Left = Form3.Left + speed
 Form4.Left = Form4.Left + speed
 Form5.Left = Form5.Left + speed
 Form6.Left = Form6.Left + speed
 Form7.Left = Form7.Left + speed
 Form8.Left = Form8.Left + speed
 Form9.Left = Form9.Left + speed
 Form10.Left = Form10.Left + speed
 Form11.Left = Form11.Left + speed
 Form12.Left = Form12.Left + speed
 Form13.Left = Form13.Left + speed
 Form14.Left = Form14.Left + speed
 Form17.Left = Form17.Left + speed
 speed = speed + 5
Else
 Unload Form1
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
 Unload Form16
 Unload Form17
 Unload Form18
 Unload Form19
 Unload Form20
 Unload Me
End If
End Sub
