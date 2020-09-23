VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm5x5edit 
   BackColor       =   &H00955800&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "5x5 Editor"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   2625
   ForeColor       =   &H00000000&
   Icon            =   "frm5x5edit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   2625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00955800&
      Caption         =   "Stage Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   90
      TabIndex        =   26
      Top             =   3015
      Width           =   2490
      Begin VB.TextBox txtDescription 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   90
         MaxLength       =   15
         TabIndex        =   27
         Text            =   "Stage"
         Top             =   315
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00955800&
      ForeColor       =   &H00FFFFFF&
      Height          =   1905
      Left            =   675
      TabIndex        =   0
      Top             =   1035
      Width           =   1860
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   25
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   24
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   23
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   22
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   21
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   20
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   19
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   18
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   17
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   16
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   11
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   15
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   12
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   14
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   13
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   13
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   14
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   12
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   15
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   11
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   16
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   10
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   17
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   9
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   18
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   8
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   19
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   7
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   20
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   6
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   21
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   5
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   22
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   4
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   23
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   3
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   24
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   2
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   25
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   1
         Top             =   1485
         Width           =   300
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.pcs|*.pcs"
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   2610
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   38
      Top             =   3780
      Width           =   2535
   End
   Begin VB.Image Image12 
      Height          =   420
      Left            =   -45
      Picture         =   "frm5x5edit.frx":1272
      Stretch         =   -1  'True
      Top             =   3690
      Width           =   2790
   End
   Begin VB.Label lblRow1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   37
      Top             =   1260
      Width           =   555
   End
   Begin VB.Label lblRow2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   36
      Top             =   1575
      Width           =   555
   End
   Begin VB.Label lblRow3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   35
      Top             =   1890
      Width           =   555
   End
   Begin VB.Label lblRow4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   34
      Top             =   2205
      Width           =   555
   End
   Begin VB.Label lblRow5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   33
      Top             =   2520
      Width           =   555
   End
   Begin VB.Line Line3 
      X1              =   630
      X2              =   0
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line4 
      X1              =   585
      X2              =   0
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblCol1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1005
      Left            =   855
      TabIndex        =   32
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1005
      Left            =   1170
      TabIndex        =   31
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1005
      Left            =   1485
      TabIndex        =   30
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1005
      Left            =   1800
      TabIndex        =   29
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1005
      Left            =   2115
      TabIndex        =   28
      Top             =   0
      Width           =   330
   End
   Begin VB.Line Line1 
      X1              =   855
      X2              =   855
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Line Line2 
      X1              =   2430
      X2              =   2430
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Image Image6 
      Height          =   1005
      Left            =   900
      Picture         =   "frm5x5edit.frx":19BC
      Stretch         =   -1  'True
      Top             =   45
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   1005
      Left            =   1215
      Picture         =   "frm5x5edit.frx":213D
      Stretch         =   -1  'True
      Top             =   45
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   1005
      Left            =   1530
      Picture         =   "frm5x5edit.frx":28BE
      Stretch         =   -1  'True
      Top             =   45
      Width           =   240
   End
   Begin VB.Image Image9 
      Height          =   1005
      Left            =   1845
      Picture         =   "frm5x5edit.frx":303F
      Stretch         =   -1  'True
      Top             =   45
      Width           =   240
   End
   Begin VB.Image Image10 
      Height          =   1005
      Left            =   2160
      Picture         =   "frm5x5edit.frx":37C0
      Stretch         =   -1  'True
      Top             =   45
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   0
      Picture         =   "frm5x5edit.frx":3F41
      Stretch         =   -1  'True
      Top             =   1305
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   0
      Picture         =   "frm5x5edit.frx":42C5
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   645
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   0
      Picture         =   "frm5x5edit.frx":4649
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   0
      Picture         =   "frm5x5edit.frx":49CD
      Stretch         =   -1  'True
      Top             =   2250
      Width           =   645
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   0
      Picture         =   "frm5x5edit.frx":4D51
      Stretch         =   -1  'True
      Top             =   2565
      Width           =   645
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "&Open.."
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save As.."
         Shortcut        =   ^S
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu edit 
      Caption         =   "E&dit"
      Begin VB.Menu clear 
         Caption         =   "Clear All"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu protect 
         Caption         =   "Protect"
      End
   End
End
Attribute VB_Name = "frm5x5edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clear_Click()
Dim PicNum As Integer
PicNum = 1

Do While PicNum < 26
    Square(PicNum).Picture = LoadPicture()
    PicEDIT(PicNum) = False
    PicNum = PicNum + 1
Loop

txtDescription.Text = "Stage"

lblRow1.Caption = "0"
lblRow2.Caption = "0"
lblRow3.Caption = "0"
lblRow4.Caption = "0"
lblRow5.Caption = "0"
lblCol1.Caption = "0"
lblCOl2.Caption = "0"
lblCol3.Caption = "0"
lblCol4.Caption = "0"
lblCol5.Caption = "0"

End Sub

Private Sub exit_Click()

frmMain.Show
Unload Me

End Sub

Private Sub Form_Load()
Protected = False
Dim tmp As Integer
tmp = 1

Do While tmp < 26
    PicEDIT(tmp) = False
tmp = tmp + 1
Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub open_Click()

On Error GoTo errorhandler:

Dim tmp As Integer
Dim Tmp2 As Boolean
tmp = 1

CommonDialog1.ShowOpen
Open CommonDialog1.FileName For Input As #1
Input #1, StageSize, Pic(1), Pic(2), Pic(3), Pic(4), Pic(5), Pic(6), Pic(7), Pic(8), Pic(9), Pic(10), Pic(11), Pic(12), Pic(13), Pic(14), Pic(15), Pic(16), Pic(17), Pic(18), Pic(19), Pic(20), Pic(21), Pic(22), Pic(23), Pic(24), Pic(25)
Input #1, Row(1), Row(2), Row(3), Row(4), Row(5), Col(1), Col(2), Col(3), Col(4), Col(5), Description, Tmp2
Close #1

If StageSize = "5x5" And Tmp2 = False Then
    
    Do While tmp < 26
        If Pic(tmp) = True Then
            Square(tmp).Picture = LoadPicture(App.Path & "\true.gif")
            PicEDIT(tmp) = True
        ElseIf Pic(tmp) = False Then
            Square(tmp).Picture = LoadPicture()
            PicEDIT(tmp) = False
        End If
    tmp = tmp + 1
    Loop
    
    lblRow1.Caption = Row(1)
    lblRow2.Caption = Row(2)
    lblRow3.Caption = Row(3)
    lblRow4.Caption = Row(4)
    lblRow5.Caption = Row(5)
    
    lblCol1.Caption = Col(1)
    lblCOl2.Caption = Col(2)
    lblCol3.Caption = Col(3)
    lblCol4.Caption = Col(4)
    lblCol5.Caption = Col(5)
    
    txtDescription.Text = Description
    lblStatus.Caption = "Load OK! -- " & CommonDialog1.FileTitle
ElseIf Tmp2 = True Then
    lblStatus.Caption = "ERROR! This file is protected!"
Else
    lblStatus.Caption = "ERROR! Grid Size not correct!"
End If

errorhandler:
    Select Case Err
    Case Is = 75
    End Select

End Sub

Private Sub protect_Click()

If protect.Checked = False Then
    protect.Checked = True
    MsgBox "Warning: You and/or anyone else will not be able to open this file for editing once it is saved!"
    Protected = True
ElseIf protect.Checked = True Then
    protect.Checked = False
    Protected = False
End If

End Sub

Private Sub save_Click()
On Error GoTo errorhandler
CommonDialog1.ShowSave

Open CommonDialog1.FileName For Output As #1
Write #1, "5x5", PicEDIT(1), PicEDIT(2), PicEDIT(3), PicEDIT(4), PicEDIT(5), PicEDIT(6), PicEDIT(7), PicEDIT(8), PicEDIT(9), PicEDIT(10), PicEDIT(11), PicEDIT(12), PicEDIT(13), PicEDIT(14), PicEDIT(15), PicEDIT(16), PicEDIT(17), PicEDIT(18), PicEDIT(19), PicEDIT(20), PicEDIT(21), PicEDIT(22), PicEDIT(23), PicEDIT(24), PicEDIT(25)
Write #1, lblRow1.Caption, lblRow2.Caption, lblRow3.Caption, lblRow4.Caption, lblRow5.Caption, lblCol1.Caption, lblCOl2.Caption, lblCol3.Caption, lblCol4.Caption, lblCol5.Caption, txtDescription.Text, Protected
Close #1

lblStatus.Caption = "Stage Saved Successfully!"

errorhandler:
    Select Case Err
    Case Is = 75
    End Select
End Sub

Private Sub Square_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    PicEDIT(Index) = True
    Square(Index).Picture = LoadPicture(App.Path & "\true.gif")
ElseIf Button = 2 Then
    PicEDIT(Index) = False
    Square(Index).Picture = LoadPicture()
End If

Call CalcRow1
Call CalcRow2
Call CalcRow3
Call CalcRow4
Call CalcRow5
Call CalcCol1
Call CalcCol2
Call CalcCol3
Call CalcCol4
Call CalcCol5
End Sub

Private Sub CalcRow1()

Dim RowCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 1
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 6
    If PicEDIT(PicCount) = True And LastEdit = True Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        PicCount = PicCount + 1
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    End If
Loop

tmp = 1
Do Until tmp = 6
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then lblRow1.Caption = "0" Else lblRow1.Caption = RowEdit
End Sub

Private Sub CalcRow2()

Dim RowCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 6
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 11
    If PicEDIT(PicCount) = True And LastEdit = True Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        PicCount = PicCount + 1
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    End If
Loop

tmp = 1
Do Until tmp = 6
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then lblRow2.Caption = "0" Else lblRow2.Caption = RowEdit
End Sub

Private Sub CalcRow3()

Dim RowCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 11
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 16
    If PicEDIT(PicCount) = True And LastEdit = True Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        PicCount = PicCount + 1
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    End If
Loop

tmp = 1
Do Until tmp = 6
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then lblRow3.Caption = "0" Else lblRow3.Caption = RowEdit
End Sub

Private Sub CalcRow4()

Dim RowCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 16
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 21
    If PicEDIT(PicCount) = True And LastEdit = True Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        PicCount = PicCount + 1
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    End If
Loop

tmp = 1
Do Until tmp = 6
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then lblRow4.Caption = "0" Else lblRow4.Caption = RowEdit
End Sub

Private Sub CalcRow5()
Dim RowCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 21
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 26
    If PicEDIT(PicCount) = True And LastEdit = True Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        PicCount = PicCount + 1
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    End If
Loop

tmp = 1
Do Until tmp = 6
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then lblRow5.Caption = "0" Else lblRow5.Caption = RowEdit
End Sub

Private Sub CalcCol1()

Dim ColCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 1
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 22
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 5
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    End If
Loop

tmp = 1
Do Until tmp = 6
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then lblCol1.Caption = 0 Else lblCol1.Caption = ColEdit
End Sub

Private Sub CalcCol2()

Dim ColCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 2
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 23
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 5
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    End If
Loop

tmp = 1
Do Until tmp = 6
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then lblCOl2.Caption = 0 Else lblCOl2.Caption = ColEdit
End Sub

Private Sub CalcCol3()

Dim ColCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 3
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 24
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 5
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    End If
Loop

tmp = 1
Do Until tmp = 6
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then lblCol3.Caption = 0 Else lblCol3.Caption = ColEdit
End Sub

Private Sub CalcCol4()

Dim ColCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 4
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 25
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 5
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    End If
Loop

tmp = 1
Do Until tmp = 6
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then lblCol4.Caption = 0 Else lblCol4.Caption = ColEdit

End Sub

Private Sub CalcCol5()

Dim ColCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 5
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 26
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 5
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    End If
Loop

tmp = 1
Do Until tmp = 6
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then lblCol5.Caption = 0 Else lblCol5.Caption = ColEdit

End Sub
