VERSION 5.00
Begin VB.Form Form21 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于软件"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9015
   Icon            =   "Form21.frx":0000
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   9015
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   0
      Picture         =   "Form21.frx":6988A
      ScaleHeight     =   2100
      ScaleWidth      =   9000
      TabIndex        =   0
      Top             =   0
      Width           =   9030
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bilibili : 软硬工作室"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   1080
      TabIndex        =   3
      Top             =   2500
      Width           =   2805
   End
   Begin VB.Image Image2 
      Height          =   810
      Left            =   5640
      Picture         =   "Form21.frx":81DB3
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Powered by"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4200
      TabIndex        =   2
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "Form21.frx":8F00F
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Call Shell("cmd /c start https://space.bilibili.com/404100570")
End Sub
