VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple Help"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7740
   Icon            =   "SimpleHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OKbut 
      Caption         =   "OK"
      Height          =   435
      Left            =   105
      TabIndex        =   4
      Top             =   3150
      Width           =   960
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1470
      Top             =   3150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SimpleHelp.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SimpleHelp.frx":0EE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Description"
      Height          =   2955
      Left            =   3465
      TabIndex        =   1
      Top             =   105
      Width           =   4215
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   4005
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contents"
      Height          =   2955
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3270
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2640
         Left            =   105
         TabIndex        =   3
         Top             =   210
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   4657
         _Version        =   393217
         Indentation     =   706
         LabelEdit       =   1
         Style           =   5
         HotTracking     =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "SHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------\
'Author: Richard E. Gagnon.                                |
'URL:    http://members.cox.net/reg501/                    |
'Email:  reg501@cox.net                                    |
'Copyright © 2007 Richard E. Gagnon. All Rights Reserved.  |
'----------------------------------------------------------/
Private Const SC = 234      'Content Node Start
Private Const FC = 235      'Content Node End
Private Const SS1 = 224     'Subject1 Title Start
Private Const FS1 = 229     'Subject1 Title End
Private Const ST1 = 244     'Subject1 Text Start
Private Const FT1 = 246     'Subject1 Text End
Private Const SS2 = 251     'Subject2 Title Start
Private Const FS2 = 252     'Subject2 Title End
Private Const ST2 = 236     'Subject2 Text Start
Private Const FT2 = 237     'Subject2 Text End
Private HelpSpace() As Byte 'Open File Array
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
Dim FEH As Long
Dim Node1 As Node, Node2 As Node, Node3 As Node
Dim Fnum As Integer
Dim HelpFileName As String
Dim sText As String
Dim S1 As Long, FL As Long
Dim I As Long, J As Long
Fnum = FreeFile
HelpFileName = App.Path & "\SimpleHelp.shf"
Open HelpFileName For Binary Access Read As Fnum
FL = FileLen(HelpFileName)
ReDim HelpSpace(1 To FL)
Get Fnum, , HelpSpace
Close Fnum
TreeView1.Nodes.Clear
For I = 1 To FL
    Select Case HelpSpace(I)
        Case SC: S1 = I + 1
        Case FC
            sText = ""
            For J = S1 To I - 1: sText = sText & Chr(HelpSpace(J)): Next J
            Set Node1 = TreeView1.Nodes.Add(, , , sText, 1, 1)
        Case SS1: S1 = I + 1
        Case FS1
            sText = ""
            For J = S1 To I - 1: sText = sText & Chr(HelpSpace(J)): Next J
            Set Node2 = TreeView1.Nodes.Add(Node1.Index, tvwChild, , sText, 2, 2)
        Case ST1: S1 = I + 1
        Case FT1: Node2.Tag = Str(S1) & "," & Str(I - 1)
        Case SS2: S1 = I + 1
        Case FS2
            sText = ""
            For J = S1 To I - 1: sText = sText & Chr(HelpSpace(J)): Next J
            Set Node3 = TreeView1.Nodes.Add(Node2.Index, tvwChild, , sText, 2, 2)
        Case ST2: S1 = I + 1
        Case FT2
            Node3.Tag = Str(S1) & "," & Str(I - 1)
            Node2.Image = 1: Node2.SelectedImage = 1
    End Select
Next I
FEH = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 2 Or 1)
End Sub

Private Sub OKbut_Click()
Unload Me
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Tag <> "" Then
    Dim I As Long
    Dim TXT As String
    Dim mTag() As String
    mTag = Split(Node.Tag, ",")
    For I = Val(mTag(0)) To Val(mTag(1))
        TXT = TXT & Chr(HelpSpace(I))
    Next I
    Text1.Text = TXT
Else
    Text1.Text = ""
End If
End Sub
