VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SHComposer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Help Composer"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   Icon            =   "SHComposer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   840
      Top             =   4830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox Tbar 
      BorderStyle     =   0  'None
      Height          =   510
      Index           =   0
      Left            =   210
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   491
      TabIndex        =   4
      Top             =   210
      Width           =   7365
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   4
         Left            =   2520
         Picture         =   "SHComposer.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Save As"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   0
         Left            =   0
         Picture         =   "SHComposer.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exit"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   1
         Left            =   630
         Picture         =   "SHComposer.frx":163E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "New"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   2
         Left            =   1260
         Picture         =   "SHComposer.frx":1DA8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Open"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   3
         Left            =   1890
         Picture         =   "SHComposer.frx":2512
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Simple Help Composer"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3255
         TabIndex        =   9
         Top             =   0
         Width           =   4110
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   210
      Top             =   4725
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
            Picture         =   "SHComposer.frx":2C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHComposer.frx":33F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Description"
      Height          =   3270
      Left            =   3465
      TabIndex        =   1
      Top             =   840
      Width           =   4215
      Begin VB.CheckBox Instruct 
         Caption         =   "Instructions"
         Height          =   225
         Left            =   1470
         TabIndex        =   11
         Top             =   2940
         Value           =   1  'Checked
         Width           =   1380
      End
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
         Index           =   0
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
      Height          =   3270
      Left            =   105
      TabIndex        =   0
      Top             =   840
      Width           =   3270
      Begin MSComctlLib.TreeView TV1 
         Height          =   2955
         Left            =   105
         TabIndex        =   3
         Top             =   210
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   5212
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   706
         LabelEdit       =   1
         Style           =   5
         FullRowSelect   =   -1  'True
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   645
      Index           =   0
      Left            =   105
      Top             =   105
      Width           =   7575
   End
   Begin VB.Menu NodeEdit 
      Caption         =   "EditNode"
      Visible         =   0   'False
      Begin VB.Menu DoNode 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu DoNode 
         Caption         =   "Delete"
         Index           =   1
      End
      Begin VB.Menu DoNode 
         Caption         =   "Insert"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu DoNode 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu DoNode 
         Caption         =   "Rename"
         Index           =   4
      End
   End
End
Attribute VB_Name = "SHComposer"
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
Private FL As Long          'Open File Length
Private OpenFileName As String

Private Sub cmdBut_Click(Index As Integer)
Select Case Index
    Case 0: Unload Me
    Case 1: GoNew
    Case 2: GoOpen
    Case 3: GoSave False
    Case 4: GoSave True
End Select
End Sub
Private Sub SaveFile(fName As String)
Dim Fnum As Integer
Fnum = FreeFile
Open fName For Output As Fnum
Dim Node As Node
For Each Node In TV1.Nodes
    If Node.Tag = 0 Then
        Print #Fnum, Chr(SC) & Node.Text & Chr(FC)
        If Node.Children > 0 Then
            Dim Child As Node
            Set Child = TV1.Nodes(Node.Index).Child
            Do Until Child Is Nothing
                Print #Fnum, Chr(SS1) & Child.Text & Chr(FS1)
                Print #Fnum, Chr(ST1) & Text1(Child.Tag).Text & Chr(FT1)
                If Child.Children > 0 Then
                    Dim GrandChild As Node
                    Set GrandChild = Child.Child
                    Do Until GrandChild Is Nothing
                        Print #Fnum, Chr(SS2) & GrandChild.Text & Chr(FS2)
                        Print #Fnum, Chr(ST2) & Text1(GrandChild.Tag).Text & Chr(FT2)
                        Set GrandChild = GrandChild.Next
                    Loop
                End If
                Set Child = Child.Next
            Loop
        End If
        Write #Fnum, "----------------------------------------"
    End If
Next
Close Fnum
Set Child = Nothing         'Free up memory
Set GrandChild = Nothing    'Free up memory
End Sub
Private Sub GoSave(Stype As Boolean)
On Error GoTo OPNerr
If Stype Then
    CD1.Filter = "SHF files (*.shf)|*.shf"
    CD1.DialogTitle = "SAVE FILE AS"
    CD1.FileName = OpenFileName
    CD1.Flags = cdlOFNFileMustExist
    CD1.ShowSave
    OpenFileName = CD1.FileName
    Me.Caption = OpenFileName
End If
If OpenFileName = "" Then
    If Err <> 32755 Then MsgBox "No filename. Please enter a Filename" & vbCrLf & "Try using 'Save As' or 'New'", vbInformation, "No Filename"
Else
    SaveFile OpenFileName
    Dim Fnum As Integer
    FL = FileLen(OpenFileName)
    If FL > 0 Then
        Fnum = FreeFile
        ReDim HelpSpace(1 To FL)
        Open OpenFileName For Binary Access Read As Fnum
        Get Fnum, , HelpSpace
        Close Fnum
    End If
End If
Exit Sub
OPNerr:
If Err <> 32755 Then MsgBox (Error & vbCr & vbCr & "Error Number: " & Str(Err)), vbCritical, "! ERROR !"
End Sub
Private Sub GoOpen()
Dim Fnum As Integer
Dim Node1 As Node, Node2 As Node, Node3 As Node
Dim sText As String
Dim S1 As Long
Dim I As Long, J As Long, K As Long
Dim FZ As Byte
Dim TT As String
On Error GoTo OPNerr
CD1.Filter = "SHF files (*.shf)|*.shf"
CD1.DialogTitle = "OPEN FILE"
CD1.FileName = ""
CD1.Flags = cdlOFNFileMustExist
CD1.ShowOpen
OpenFileName = CD1.FileName
If OpenFileName <> "" Then
    Me.Caption = OpenFileName
    FL = FileLen(OpenFileName)
    ClearAll
    'Max file size = 2,147,483,647 bytes
    If FL > 0 Then
        Me.MousePointer = 11
        Fnum = FreeFile
        ReDim HelpSpace(1 To FL)
        Open OpenFileName For Binary Access Read As Fnum
        Get Fnum, , HelpSpace
        Close Fnum
        For I = 1 To FL
            Select Case HelpSpace(I)
                Case SC: S1 = I + 1
                Case FC
                    sText = ""
                    For J = S1 To I - 1: sText = sText & Chr(HelpSpace(J)): Next J
                    Set Node1 = TV1.Nodes.Add(, , , sText, 1, 1)
                    Node1.Tag = 0
                Case SS1: S1 = I + 1
                Case FS1
                    sText = ""
                    For J = S1 To I - 1: sText = sText & Chr(HelpSpace(J)): Next J
                    Set Node2 = TV1.Nodes.Add(Node1.Index, tvwChild, , sText, 2, 2)
                    Node2.Tag = Text1.UBound + 1
                    Load Text1(Node2.Tag)
                    FormatTextBox (Node2.Tag)
                Case ST1: S1 = I + 1
                Case FT1:
                    TT = ""
                    For K = S1 To I - 1: TT = TT & Chr(HelpSpace(K)): Next K
                    Text1(Node2.Tag).Text = TT
                 Case SS2: S1 = I + 1
                Case FS2
                    sText = ""
                    For J = S1 To I - 1: sText = sText & Chr(HelpSpace(J)): Next J
                    Set Node3 = TV1.Nodes.Add(Node2.Index, tvwChild, , sText, 2, 2)
                    Node3.Tag = Text1.UBound + 1
                    Load Text1(Node3.Tag)
                    FormatTextBox (Node3.Tag)
                Case ST2: S1 = I + 1
                Case FT2:
                    Node2.Image = 1
                    Node2.SelectedImage = 1
                    TT = ""
                    For K = S1 To I - 1: TT = TT & Chr(HelpSpace(K)): Next K
                    Text1(Node3.Tag).Text = TT
            End Select
        Next I
        Me.MousePointer = 0
    End If
End If
Set Node1 = Nothing  'Free up memory
Set Node2 = Nothing  'Free up memory
Set Node3 = Nothing  'Free up memory
Exit Sub
OPNerr:
If Err <> 32755 Then MsgBox (Error & vbCr & vbCr & "Error Number: " & Str(Err)), vbCritical, "! ERROR !"
End Sub
Private Sub ClearAll()
TV1.Nodes.Clear
Dim control As TextBox
For Each control In Text1
    If control.Index > 0 Then Unload Text1(control.Index)
Next
End Sub
Private Sub GoNew()
On Error GoTo OPNerr
Dim Confirm As Integer
If OpenFileName = "" Then
    Confirm = vbYes
Else
    Confirm = MsgBox("Are you sure you want to create a new Help File?", vbQuestion + vbYesNo, "New File")
End If
If Confirm = vbYes Then
    CD1.Filter = "SHF files (*.shf)|*.shf"
    CD1.DialogTitle = "NEW FILE"
    CD1.FileName = ""
    CD1.Flags = cdlOFNFileMustExist
    CD1.ShowSave
    OpenFileName = CD1.FileName
    If OpenFileName <> "" Then
        Me.Caption = OpenFileName
        ClearAll
    End If
End If
Exit Sub
OPNerr:
If Err <> 32755 Then MsgBox (Error & vbCr & vbCr & "Error Number: " & Str(Err)), vbCritical, "! ERROR !"
End Sub
Private Sub DoNode_Click(Index As Integer)
Dim Node As Node
Dim I As Integer
Select Case Index
    Case 0  'Add Node
        If TV1.SelectedItem Is Nothing Then
            Set Node = TV1.Nodes.Add(, , , "Content", 1, 1)
            Node.Tag = 0
        Else
            Set Node = TV1.Nodes.Add(TV1.SelectedItem.Index, tvwChild, , "Subject", 2, 2)
            Node.Tag = Text1.UBound + 1
            Load Text1(Node.Tag)
            FormatTextBox (Node.Tag)
            TV1.SelectedItem.Expanded = True
            If Node.Parent.Tag > 0 Then
                Node.Parent.Image = 1
                Node.Parent.SelectedImage = 1
            End If
        End If
    Case 1  'Delete Node(s)
        If Not TV1.SelectedItem Is Nothing Then
            Set Node = TV1.Nodes(TV1.SelectedItem.Index)
            UnloadTextBoxes Node
            If Not Node.Parent Is Nothing Then
                If Node.Parent.Tag > 0 And Node.Parent.Children = 1 Then
                    Node.Parent.Image = 2
                    Node.Parent.SelectedImage = 2
                End If
            End If
            TV1.Nodes.Remove Node.Index
            If Not TV1.SelectedItem Is Nothing Then
                If TV1.SelectedItem.Tag > 0 Then
                    Text1(TV1.SelectedItem.Tag).ZOrder vbBringToFront
                Else
                    Text1(0).ZOrder vbBringToFront
                End If
            End If
        End If
    Case 2  'Insert Node
         'Set Node1 = TV1.Nodes.Add(1, , , "TBD", 1, 1)
    Case 4  'Rename Node
        Dim TT As String
        TT = InputBox("Enter Label Name", "Label Name", TV1.SelectedItem.Text)
        If TT <> "" Then TV1.SelectedItem.Text = TT
        TV1_NodeClick TV1.Nodes(TV1.SelectedItem.Index)
End Select
Set Node = Nothing  'Free up memory
End Sub

Private Sub Form_Load()
ShowInstructions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Fnum As Integer
Dim TempFname As String
Dim TempSpace() As Byte
Dim I As Integer
Dim TFL As Long
Dim control As TextBox
TempFname = App.Path & "\SHTEMP.TMP"
SaveFile TempFname
TFL = FileLen(TempFname)
If OpenFileName <> "" Then
    I = vbNo
    If TFL > 0 Then
        Fnum = FreeFile
        Open TempFname For Binary Access Read As Fnum
        ReDim TempSpace(1 To TFL)
        Get Fnum, , TempSpace
        Close Fnum
        If TFL <> FL Then
            I = MsgBox("Save changes to '" & OpenFileName & "'?", vbYesNo + vbQuestion, "COMPOSER")
        Else
            Dim FS As Long
            Dim Fequal As Boolean
            Fequal = True
            For FS = LBound(TempSpace) To UBound(TempSpace)
                If TempSpace(FS) <> HelpSpace(FS) Then Fequal = False
            Next FS
            If Not Fequal Then
                I = MsgBox("Save changes to '" & OpenFileName & "'?", vbYesNo + vbQuestion, "COMPOSER")
            End If
        End If
    Else
        If TFL <> FL Then I = MsgBox("Save changes to '" & OpenFileName & "'?", vbYesNo + vbQuestion, "COMPOSER")
    End If
    If I = vbYes Then GoSave False
Else
    If TFL > 0 Then GoSave True
End If
Kill TempFname
For Each control In Text1
    If control.Index > 0 Then Unload Text1(control.Index)
Next
End Sub

Private Sub Instruct_Click()
If Instruct.Value = 0 Then
    Text1(0).Text = ""
    If Not TV1.SelectedItem Is Nothing Then
        TV1_NodeClick TV1.Nodes(TV1.SelectedItem.Index)
    End If
Else
    ShowInstructions
End If
End Sub

Private Sub TV1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If TV1.HitTest(x, y) Is Nothing Then
        DoNode(0).Caption = "Add Content"
        DoNode(0).Visible = True
        DoNode(1).Visible = False
        DoNode(3).Visible = False
        DoNode(4).Visible = False
        TV1.SelectedItem = Nothing
    Else
        TV1.HitTest(x, y).Selected = True
        DoNode(1).Visible = True
        DoNode(3).Visible = True
        DoNode(4).Visible = True
        If TV1.SelectedItem.Tag = 0 Then
            DoNode(0).Visible = True
            DoNode(0).Caption = "Add Subject"
            DoNode(1).Caption = "Delete Content"
        Else
            If TV1.SelectedItem.Parent.Tag = 0 Then
                DoNode(0).Caption = "Add Subject"
                DoNode(1).Caption = "Delete Subject"
                DoNode(0).Visible = True
            Else
                DoNode(0).Visible = False
                DoNode(1).Caption = "Delete Subject"
            End If
        End If
    End If
    Me.PopupMenu NodeEdit
End If
End Sub

Private Sub TV1_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Tag > 0 Then
    Text1(Node.Tag).ZOrder vbBringToFront
Else
    Text1(0).ZOrder vbBringToFront
End If
End Sub
Private Sub FormatTextBox(Index As Integer)
With Text1(Index)
    Set .Container = Frame2
    .Text = ""
    .ZOrder vbSendToBack
    .Locked = False
    .Visible = True
    .Width = Text1(0).Width
    .Height = Text1(0).Height
    .Top = Text1(0).Top
    .Left = Text1(0).Left
End With
End Sub
Private Sub UnloadTextBoxes(ByVal n As Node)
'Destroy all Textboxes in deleted node recursively
If n.Tag > 0 Then Unload Text1(n.Tag)
Set n = n.Child
Do Until n Is Nothing
    UnloadTextBoxes n
    Set n = n.Next
Loop
End Sub

Private Sub ShowInstructions()
Text1(0).Text = "To Add new Content, right click the Contents window." & vbCrLf & vbCrLf & _
                "To Add a subject to the Content right click the Contents Node." & vbCrLf & vbCrLf & _
                "To enter Help text, click the subject Node and start typing in the Description window" & vbCrLf & vbCrLf & _
                "When finished Composing the file, Click the save button"
Text1(0).ZOrder vbBringToFront
End Sub
