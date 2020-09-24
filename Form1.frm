VERSION 5.00
Object = "{7CC4F088-D6CE-42B8-8C31-661DD3725862}#36.0#0"; "ZaidTrans.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Translate Over the Internet"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":1762
      Left            =   2160
      List            =   "Form1.frx":1775
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin ZaidTrans.ZaidTranslator ZaidTranslator1 
      Height          =   405
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Translate"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   4095
      Left            =   3480
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":17D9
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Translator Between"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Translated Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text to Translate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False

ZaidTranslator1.Set_Language Combo1.ListIndex

Text2.Text = ZaidTranslator1.Translate(Text1.Text)

Text2.RightToLeft = ZaidTranslator1.Get_RightToLeft
Text2.Alignment = ZaidTranslator1.Get_Alignment

Command1.Enabled = True
End Sub

Private Sub Command2_Click()
ZaidTranslator1.About
End Sub

Private Sub Form_Load()
ZaidTranslator1.Set_Language English_Arabic
Combo1.ListIndex = 0
End Sub
