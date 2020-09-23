VERSION 5.00
Begin VB.Form MyAppFrm 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Height          =   435
      Left            =   1680
      Picture         =   "SimpleHelpFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Help"
      Top             =   105
      Width           =   435
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   435
      Left            =   1155
      TabIndex        =   2
      Top             =   105
      Width           =   435
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   435
      Left            =   630
      TabIndex        =   1
      Top             =   105
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   435
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Your VB Application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   315
      TabIndex        =   3
      Top             =   2310
      Width           =   6945
   End
End
Attribute VB_Name = "MyAppFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command4_Click()
SHelp.Show
End Sub
