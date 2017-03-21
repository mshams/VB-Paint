VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paint By M.SH"
   ClientHeight    =   5070
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6855
   Icon            =   "paint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3720
      Left            =   960
      MousePointer    =   99  'Custom
      ScaleHeight     =   3660
      ScaleWidth      =   5715
      TabIndex        =   17
      Top             =   120
      Width           =   5775
      Begin MSComDlg.CommonDialog dlg 
         Left            =   5160
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "*.Bmp"
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   6615
      Begin VB.CommandButton cmdcus 
         Caption         =   "Custom"
         Height          =   735
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   19
         Left            =   4200
         TabIndex        =   39
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   18
         Left            =   4200
         TabIndex        =   38
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   17
         Left            =   3840
         TabIndex        =   37
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   16
         Left            =   3840
         TabIndex        =   36
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   15
         Left            =   3480
         TabIndex        =   35
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   14
         Left            =   3480
         TabIndex        =   34
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   13
         Left            =   3120
         TabIndex        =   33
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   3120
         TabIndex        =   32
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   2760
         TabIndex        =   31
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   2760
         TabIndex        =   30
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   2400
         TabIndex        =   29
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   2400
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   2040
         TabIndex        =   27
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00004040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   2040
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   25
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   21
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame Fra 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.Label l2 
         AutoSize        =   -1  'True
         Caption         =   "y = "
         Height          =   195
         Left            =   5640
         TabIndex        =   42
         Top             =   600
         Width           =   255
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "x = "
         Height          =   195
         Left            =   5640
         TabIndex        =   41
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape c 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   1
         Left            =   360
         Top             =   480
         Width           =   375
      End
      Begin VB.Shape c 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   2
         Left            =   120
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   735
      Begin VB.OptionButton f 
         Caption         =   "No fill"
         Height          =   255
         Index           =   1
         Left            =   70
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1215
         Width           =   600
      End
      Begin VB.OptionButton f 
         Caption         =   "Fill"
         Height          =   255
         Index           =   0
         Left            =   70
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   975
         Width           =   600
      End
      Begin VB.OptionButton s 
         Caption         =   "Size 3"
         Height          =   255
         Index           =   2
         Left            =   70
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   660
         Width           =   600
      End
      Begin VB.OptionButton s 
         Caption         =   "Size 2"
         Height          =   255
         Index           =   1
         Left            =   70
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   420
         Width           =   600
      End
      Begin VB.OptionButton s 
         Caption         =   "Size 1"
         Height          =   255
         Index           =   0
         Left            =   70
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   600
      End
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   11
      Left            =   480
      MaskColor       =   &H00008080&
      Picture         =   "paint.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   10
      Left            =   120
      MaskColor       =   &H00008080&
      Picture         =   "paint.frx":047C
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   9
      Left            =   480
      MaskColor       =   &H00008080&
      Picture         =   "paint.frx":05EE
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   8
      Left            =   120
      MaskColor       =   &H00008080&
      Picture         =   "paint.frx":0760
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   7
      Left            =   480
      MaskColor       =   &H00008080&
      Picture         =   "paint.frx":08D2
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   375
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   6
      Left            =   120
      MaskColor       =   &H00008080&
      Picture         =   "paint.frx":0A50
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1200
      Width           =   375
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   5
      Left            =   480
      Picture         =   "paint.frx":0BCE
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   4
      Left            =   120
      Picture         =   "paint.frx":0D4C
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   3
      Left            =   480
      Picture         =   "paint.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   2
      Left            =   120
      Picture         =   "paint.frx":1048
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton op 
      Caption         =   "§"
      Height          =   375
      Index           =   1
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton op 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "paint.frx":11C6
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.Menu mnufil 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuopn 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusav 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusas 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuext 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabt 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "fp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------
'-             VB Programming Project        -
'-                                           -
'-  Project Name  : Paint                    -
'-  Coded By      : Mohammad Shams Javi      -
'-  Mail          : M.Shams.J@Gmail.com      -
'---------------------------------------------

Private Sub cmdcus_Click()
dlg.Action = 3
c(1).FillColor = dlg.Color
End Sub

Private Sub Fra_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
c(Button).FillColor = Fra(index).BackColor
End Sub

Private Sub mnuabt_Click()
'About Me
MsgBox "             VB Programming Project        " + vbCrLf + _
 vbCrLf + _
"  Project Name  : Paint" + vbCrLf + _
"  Coded By        : Mohammad Shams Javi" + vbCrLf + _
"  Mail                 : M.Shams.J@Gmail.com" + vbCrLf + vbCrLf + _
vbTab + "Copyright (c) 1384/8/27", , "About Me"

End Sub

Private Sub mnuext_Click()
Dim result As Integer

If PicChanged Then
    result = MsgBox("Do you want to save changes ?", 51)
    If result = vbYes Then
        Call mnusav_Click
        Exit Sub
    ElseIf result = vbCancel Then Exit Sub
    End If
End If
End
End Sub

Private Sub mnunew_Click()
Dim result As Integer

If PicChanged Then
    result = MsgBox("Do you want to save changes ?", 51)
    If result = vbYes Then
        Call mnusav_Click
        Exit Sub
    ElseIf result = vbCancel Then Exit Sub
    End If
End If

pic.Picture = LoadPicture("")
pic.Cls       'set new settings
fname = ""
PicChanged = False

End Sub

Private Sub mnuopn_Click()
Dim result As Integer

If PicChanged Then
    result = MsgBox("Do you want to save changes ?", 51)
    If result = vbYes Then
        Call mnusav_Click
        Exit Sub
    ElseIf result = vbCancel Then Exit Sub
    End If
End If

dlg.Filter = "Bitmap|*.Bmp;*.dib|Jpeg|*.jpg|All Graphic Formats|*.bmp;*.gif;*.jpg;*.wmf;*.ico;*.cur"
dlg.ShowOpen
If dlg.FileName <> "" And Dir(dlg.FileName) <> "" Then
    pic.Cls
    PicChanged = False
    fname = dlg.FileName
    pic.Picture = LoadPicture(fname)
End If
End Sub

Private Sub mnusas_Click()
dlg.Filter = "Bitmap|*.Bmp;*.dib|Icon|*.ico"
dlg.ShowSave    'save as menu
If dlg.FileName <> "" Then
    fname = dlg.FileName
    SavePicture pic.Image, fname
End If

End Sub

Private Sub mnusav_Click()
If fname = "" Then      'save Menu
    Call mnusas_Click
Else
    SavePicture pic.Image, fname
End If
End Sub

Private Sub op_Click(index As Integer)
SetEn (index)
SetCur (index)
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicChanged = True
p1.X = X: p1.Y = Y  'set star point
p2 = p1
pic.DrawWidth = GE(s) + 1  'set draw width

Select Case GE(op)      'which start point selected
    Case 0: Call draw(0, 0, Button, X, Y)
    Case 6: Call draw(6, GE(s), Button, X, Y)
    Case 4: Call FillCol(Button, X, Y)
    Case 2      'text output
        pic.CurrentX = X: fp.pic.CurrentY = Y
        pic.ForeColor = c(1).FillColor
        pic.Print InputBox("Please enter your text :")
    Case 3, 7 To 9: pic.PSet (X, Y)
    Case 5
        pic.ForeColor = vbWhite
        pic.PSet (X, Y)
End Select

End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

l1.Caption = "x = " + Str(X)
l2.Caption = "y = " + Str(Y)

If Button = 1 Or Button = 2 Then Call draw(GE(op), GE(s), Button, X, Y)

p2.X = X: p2.Y = Y

End Sub

Function GE(ctrl As Object) As Byte
Dim c As Control
For Each c In ctrl  'get enabled option index
    If c.Value = True Then GE = c.index
Next
End Function
