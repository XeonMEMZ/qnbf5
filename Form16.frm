VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form16 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÐÞ¸Ä¿Î±íÑÕÉ«"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "È¡Ïû"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
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
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   2850
      Picture         =   "Form16.frx":6988A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00979E10&
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
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00979E10&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   2700
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÐÞ¸Ä¿Î±íÑÕÉ«"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2610
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bg As String
Dim kbcount As Integer

Private Sub Command1_Click()
 CommonDialog1.CancelError = True
 On Error GoTo cuowu
 CommonDialog1.ShowColor
 Text1.BackColor = CommonDialog1.Color
 Shape1.BackColor = CommonDialog1.Color
cuowu:
 Exit Sub
End Sub

Private Sub Command2_Click()
If kbcount = 3 Then
 Form3.BackColor = Text1.BackColor
ElseIf kbcount = 4 Then
 Form4.BackColor = Text1.BackColor
ElseIf kbcount = 5 Then
 Form5.BackColor = Text1.BackColor
ElseIf kbcount = 6 Then
 Form6.BackColor = Text1.BackColor
ElseIf kbcount = 7 Then
 Form7.BackColor = Text1.BackColor
ElseIf kbcount = 8 Then
 Form8.BackColor = Text1.BackColor
ElseIf kbcount = 9 Then
 Form9.BackColor = Text1.BackColor
ElseIf kbcount = 10 Then
 Form10.BackColor = Text1.BackColor
ElseIf kbcount = 11 Then
 Form11.BackColor = Text1.BackColor
ElseIf kbcount = 12 Then
 Form12.BackColor = Text1.BackColor
ElseIf kbcount = 13 Then
 Form13.BackColor = Text1.BackColor
ElseIf kbcount = 14 Then
 Form14.BackColor = Text1.BackColor
End If
bg = Text1.BackColor
Open "data\f" & kbcount & "color.txt" For Output As #1
Print #1, bg
Close #1
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
kbcount = kbcout
Text1.Text = kbtext
Text1.BackColor = kbcolr
Shape1.BackColor = kbcolr
End Sub
