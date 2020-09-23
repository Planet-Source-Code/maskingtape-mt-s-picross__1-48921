VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00955800&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MT's Picross"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4830
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3735
      Top             =   3510
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.pcs|*.pcs"
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1530
      TabIndex        =   18
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Shape Shape9 
      Height          =   465
      Left            =   1485
      Top             =   3555
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Line Line13 
      X1              =   60
      X2              =   258
      Y1              =   207
      Y2              =   207
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   90
      TabIndex        =   17
      Top             =   1620
      Width           =   4650
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      X1              =   135
      X2              =   135
      Y1              =   126
      Y2              =   120
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   111
      X2              =   135
      Y1              =   126
      Y2              =   126
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Stages have a 15x15 grid!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Top             =   1665
      Width           =   2625
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00000000&
      Height          =   510
      Left            =   45
      Top             =   1575
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2610
      TabIndex        =   15
      Top             =   3150
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   405
      TabIndex        =   14
      Top             =   3150
      Width           =   1815
   End
   Begin VB.Shape Shape7 
      Height          =   465
      Left            =   2565
      Top             =   3105
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Shape Shape6 
      Height          =   465
      Left            =   360
      Top             =   3105
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Height          =   555
      Left            =   90
      TabIndex        =   13
      Top             =   2520
      Width           =   4650
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   90
      TabIndex        =   12
      Top             =   2115
      Width           =   4650
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   90
      TabIndex        =   11
      Top             =   1215
      Width           =   4650
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   90
      TabIndex        =   10
      Top             =   810
      Width           =   4650
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   90
      TabIndex        =   9
      Top             =   405
      Width           =   4650
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000000&
      Height          =   645
      Left            =   45
      Top             =   2475
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      Height          =   420
      Left            =   45
      Top             =   2070
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   420
      Left            =   45
      Top             =   1170
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      Height          =   420
      Left            =   45
      Top             =   765
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   420
      Left            =   45
      Top             =   360
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   161
      X2              =   161
      Y1              =   180
      Y2              =   195
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Download a new stage      off the Internet!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   2610
      TabIndex        =   8
      Top             =   2490
      Width           =   2115
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   114
      X2              =   161
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Image Image7 
      Height          =   450
      Left            =   120
      Picture         =   "frmMain.frx":1272
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
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
      Height          =   225
      Left            =   4410
      TabIndex        =   7
      Top             =   0
      Width           =   435
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Editors"
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
      Height          =   225
      Left            =   3780
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Game"
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
      Height          =   225
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   450
      Left            =   2595
      Picture         =   "frmMain.frx":168A
      Top             =   3120
      Width           =   1800
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   405
      Picture         =   "frmMain.frx":1B6A
      Top             =   3120
      Width           =   1800
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   -60
      Picture         =   "frmMain.frx":204B
      Top             =   2070
      Width           =   1500
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   15
      Picture         =   "frmMain.frx":2413
      Top             =   1170
      Width           =   1500
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   -135
      Picture         =   "frmMain.frx":279F
      Top             =   765
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   -180
      Picture         =   "frmMain.frx":2B65
      Top             =   360
      Width           =   1500
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   99
      X2              =   162
      Y1              =   157
      Y2              =   157
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Load a custom stage!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2610
      TabIndex        =   4
      Top             =   2115
      Width           =   2085
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   105
      X2              =   129
      Y1              =   97
      Y2              =   97
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Stages have a 10x10 grid."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2070
      TabIndex        =   3
      Top             =   1275
      Width           =   2625
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   99
      X2              =   141
      Y1              =   70
      Y2              =   70
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Stages have a 5x5 grid."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2295
      TabIndex        =   2
      Top             =   870
      Width           =   2355
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Learn to play the game!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2340
      TabIndex        =   1
      Top             =   450
      Width           =   2310
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   96
      X2              =   147
      Y1              =   43
      Y2              =   43
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   147
      X2              =   147
      Y1              =   43
      Y2              =   34
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   141
      X2              =   141
      Y1              =   70
      Y2              =   61
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   129
      X2              =   129
      Y1              =   97
      Y2              =   91
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   162
      X2              =   162
      Y1              =   157
      Y2              =   148
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MT's Picross! --- Version 0.5"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   2265
   End
   Begin VB.Image Image11 
      Height          =   330
      Left            =   0
      Picture         =   "frmMain.frx":2EFA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2265
   End
   Begin VB.Image Image8 
      Height          =   450
      Left            =   105
      Picture         =   "frmMain.frx":3866
      Top             =   1620
      Width           =   1500
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   1530
      Picture         =   "frmMain.frx":3C47
      Top             =   3555
      Width           =   1800
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu starteasy 
         Caption         =   "Start Easy Level"
      End
      Begin VB.Menu starthard 
         Caption         =   "Start Hard Level"
      End
      Begin VB.Menu extrahard 
         Caption         =   "Start Extra Hard Level"
      End
      Begin VB.Menu loadcustom 
         Caption         =   "Load Custom Level"
      End
      Begin VB.Menu netlevel 
         Caption         =   "Load Net Level"
      End
      Begin VB.Menu break1 
         Caption         =   "-"
      End
      Begin VB.Menu associate 
         Caption         =   "Associate .pcs Files"
      End
      Begin VB.Menu break2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu easyedit 
         Caption         =   "5 x 5 Editor"
      End
      Begin VB.Menu hardedit 
         Caption         =   "10 x 10 Editor"
      End
      Begin VB.Menu extrahardedit 
         Caption         =   "15 x 15 Editor"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu loadtutorial 
         Caption         =   "Load Tutorial"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub associate_Click()
rc = MsgBox("This will associate .pcs files with MT's Picross. This will allow you to load a stage by double clicking on the stage file. Continue?", vbYesNo, "Associate?")
If rc = vbYes Then
    Call SaveString(HKEY_CLASSES_ROOT, ".pcs", "", "pcsfile")
    Call SaveString(HKEY_CLASSES_ROOT, ".pcs", "Content Type", "text/plain")
    Call SaveString(HKEY_CLASSES_ROOT, "pcsfile", "", "MT's Picross Stage")
    Call SaveDWord(HKEY_CLASSES_ROOT, "pcsfile", "EditFlags", "0000")
    Call SaveString(HKEY_CLASSES_ROOT, "pcsfile\DefaultIcon", "", App.Path & "\" & App.EXEName & ".exe,0")
    Call SaveString(HKEY_CLASSES_ROOT, "pcsfile\Shell", "", "")
    Call SaveString(HKEY_CLASSES_ROOT, "pcsfile\Shell\Open", "", "")
    Call SaveString(HKEY_CLASSES_ROOT, "pcsfile\Shell\Open\Command", "", App.Path & "\" & App.EXEName & ".exe %1")
    MsgBox "Association Complete!"
Else
End If
End Sub

Private Sub easyedit_Click()

frm5x5edit.Show
Me.Hide

End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub extrahard_Click()
frm15x15select.Show
Me.Hide
End Sub

Private Sub extrahardedit_Click()
frm15x15edit.Show
Me.Hide
End Sub

Private Sub Form_Load()

If Command$ <> "%1" And Command$ <> "" And Done = False Then
    Open Command$ For Input As #1
    Input #1, StageSize
    Close #1

    If StageSize = "5x5" Then
        Stage = Command$
        frm5x5.Show
        Me.Hide
        Custom = True
    ElseIf StageSize = "10x10" Then
        Stage = Command$
        frm10x10.Show
        Me.Hide
        Custom = True
    ElseIf StageSize = "15x15" Then
        Stage = Command$
        frm15x15.Show
        Me.Hide
        Custom = True
    End If
End If
       
End Sub

Private Sub hardedit_Click()

frm10x10edit.Show
Me.Hide

End Sub

Private Sub Label10_Click()
frmTutorial.Show
Me.Hide
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = True
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Shape8.Visible = False
Shape9.Visible = False
End Sub

Private Sub Label11_Click()
frm5x5Select.Show
Me.Hide
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Shape2.Visible = True
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Shape8.Visible = False
Shape9.Visible = False
End Sub

Private Sub Label12_Click()
frm10x10select.Show
Me.Hide
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = True
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Shape8.Visible = False
Shape9.Visible = False
End Sub

Private Sub Label13_Click()

On Error GoTo errorhandler

CommonDialog1.ShowOpen

Open CommonDialog1.FileName For Input As #1
Input #1, StageSize
Close #1

If StageSize = "5x5" Then
    Stage = CommonDialog1.FileName
    frm5x5.Show
    Me.Hide
    Custom = True
ElseIf StageSize = "10x10" Then
    Stage = CommonDialog1.FileName
    frm10x10.Show
    Me.Hide
    Custom = True
ElseIf StageSize = "15x15" Then
    Stage = CommonDialog1.FileName
    frm15x15.Show
    Me.Hide
    Custom = True
End If

errorhandler:
    Select Case Err
    Case Is = 75
    End Select

End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = True
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Shape8.Visible = False
Shape9.Visible = False
End Sub

Private Sub Label14_Click()
frmNetLevels.Show
Me.Hide
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = True
Shape6.Visible = False
Shape7.Visible = False
Shape8.Visible = False
Shape9.Visible = False
End Sub

Private Sub Label15_Click()
frm5x5edit.Show
Me.Hide
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = True
Shape7.Visible = False
Shape8.Visible = False
Shape9.Visible = False
End Sub

Private Sub Label16_Click()
frm10x10edit.Show
Me.Hide
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = True
Shape8.Visible = False
Shape9.Visible = False
End Sub

Private Sub Label18_Click()
frm15x15select.Show
Me.Hide
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Shape8.Visible = True
Shape9.Visible = False
End Sub

Private Sub Label19_Click()
frm15x15edit.Show
Me.Hide
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Shape8.Visible = False
Shape9.Visible = True
End Sub

Private Sub Label7_Click()
Me.PopupMenu file
End Sub

Private Sub Label8_Click()
Me.PopupMenu edit
End Sub

Private Sub Label9_Click()
Me.PopupMenu help
End Sub

Private Sub loadcustom_Click()
On Error GoTo errorhandler

CommonDialog1.ShowOpen

Open CommonDialog1.FileName For Input As #1
Input #1, StageSize
Close #1

If StageSize = "5x5" Then
    Stage = CommonDialog1.FileName
    frm5x5.Show
    Me.Hide
    Custom = True
ElseIf StageSize = "10x10" Then
    Stage = CommonDialog1.FileName
    frm10x10.Show
    Me.Hide
    Custom = True
End If

errorhandler:
    Select Case Err
    Case Is = 75
    End Select

End Sub

Private Sub loadtutorial_Click()

frmTutorial.Show
Me.Hide

End Sub

Private Sub netlevel_Click()
frmNetLevels.Show
Me.Hide
End Sub

Private Sub starteasy_Click()

frm5x5Select.Show
Me.Hide

End Sub

Private Sub starthard_Click()

frm10x10select.Show
Me.Hide

End Sub
