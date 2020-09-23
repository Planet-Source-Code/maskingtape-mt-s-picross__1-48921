VERSION 5.00
Begin VB.Form frmTutorial 
   BackColor       =   &H00955800&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MT's Picross: Tutorial"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5520
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmTutorial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2925
      Top             =   90
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00955800&
      Caption         =   "Instructions"
      ForeColor       =   &H00FFFFFF&
      Height          =   2310
      Left            =   2970
      TabIndex        =   26
      Top             =   180
      Width           =   2490
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Continue - ->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   495
         TabIndex        =   28
         Top             =   1890
         Width           =   1680
      End
      Begin VB.Label lblInstructions 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to MT's Picross! Let's Get started!"
         ForeColor       =   &H00FFFFFF&
         Height          =   2265
         Left            =   135
         TabIndex        =   27
         Top             =   270
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00955800&
      Height          =   1905
      Left            =   675
      TabIndex        =   0
      Top             =   945
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
         BackColor       =   &H80000005&
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
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   3015
      X2              =   3015
      Y1              =   3195
      Y2              =   2970
   End
   Begin VB.Image imgExit 
      Height          =   450
      Left            =   4860
      Picture         =   "frmTutorial.frx":1272
      Top             =   2475
      Width           =   600
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
      TabIndex        =   40
      Top             =   1125
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
      TabIndex        =   39
      Top             =   1440
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
      TabIndex        =   38
      Top             =   1755
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
      TabIndex        =   37
      Top             =   2070
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
      TabIndex        =   36
      Top             =   2385
      Width           =   555
   End
   Begin VB.Line Line3 
      X1              =   630
      X2              =   0
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Line Line4 
      X1              =   585
      X2              =   0
      Y1              =   2745
      Y2              =   2745
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   2340
      X2              =   2340
      Y1              =   3195
      Y2              =   2970
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "XXX"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2430
      TabIndex        =   35
      Top             =   2925
      Width           =   555
   End
   Begin VB.Image Image21 
      Height          =   285
      Left            =   0
      Picture         =   "frmTutorial.frx":1908
      Stretch         =   -1  'True
      Top             =   2970
      Width           =   5535
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
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
      Top             =   2970
      Width           =   2175
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      X1              =   -90
      X2              =   5580
      Y1              =   2925
      Y2              =   2925
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
      Left            =   810
      TabIndex        =   33
      Top             =   -45
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
      Left            =   1125
      TabIndex        =   32
      Top             =   -45
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
      Left            =   1440
      TabIndex        =   31
      Top             =   -45
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
      Left            =   1755
      TabIndex        =   30
      Top             =   -45
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
      Left            =   2070
      TabIndex        =   29
      Top             =   -45
      Width           =   330
   End
   Begin VB.Line Line1 
      X1              =   810
      X2              =   810
      Y1              =   -90
      Y2              =   945
   End
   Begin VB.Line Line2 
      X1              =   2385
      X2              =   2385
      Y1              =   -90
      Y2              =   945
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2700
      X2              =   2700
      Y1              =   2340
      Y2              =   2835
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   1815
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   1035
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   870
      Left            =   675
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Image Image10 
      Height          =   1005
      Left            =   2115
      Picture         =   "frmTutorial.frx":2052
      Stretch         =   -1  'True
      Top             =   -90
      Width           =   240
   End
   Begin VB.Image Image9 
      Height          =   1005
      Left            =   1800
      Picture         =   "frmTutorial.frx":27D3
      Stretch         =   -1  'True
      Top             =   -90
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   1005
      Left            =   1485
      Picture         =   "frmTutorial.frx":2F54
      Stretch         =   -1  'True
      Top             =   -90
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   1005
      Left            =   1170
      Picture         =   "frmTutorial.frx":36D5
      Stretch         =   -1  'True
      Top             =   -90
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   1005
      Left            =   855
      Picture         =   "frmTutorial.frx":3E56
      Stretch         =   -1  'True
      Top             =   -90
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   0
      Picture         =   "frmTutorial.frx":45D7
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   0
      Picture         =   "frmTutorial.frx":495B
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   645
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   0
      Picture         =   "frmTutorial.frx":4CDF
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   0
      Picture         =   "frmTutorial.frx":5063
      Stretch         =   -1  'True
      Top             =   2115
      Width           =   645
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   0
      Picture         =   "frmTutorial.frx":53E7
      Stretch         =   -1  'True
      Top             =   2430
      Width           =   645
   End
End
Attribute VB_Name = "frmTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgExit_Click()

frmMain.Show
Unload Me

End Sub

Private Sub Form_Load()

Tutorial = 1
Time = 300
lblTime.Caption = Time

Open App.Path & "\stages\tutorial.pcs" For Input As #1
Input #1, StageSize, Pic(1), Pic(2), Pic(3), Pic(4), Pic(5), Pic(6), Pic(7), Pic(8), Pic(9), Pic(10), Pic(11), Pic(12), Pic(13), Pic(14), Pic(15), Pic(16), Pic(17), Pic(18), Pic(19), Pic(20), Pic(21), Pic(22), Pic(23), Pic(24), Pic(25)
Input #1, Row(1), Row(2), Row(3), Row(4), Row(5), Col(1), Col(2), Col(3), Col(4), Col(5), Description
Close #1

lblRow1.Caption = Row(1)
lblRow2.Caption = Row(2)
lblRow3.Caption = Row(3)
lblRow4.Caption = Row(4)
lblRow5.Caption = Row(5)

lblCol1.Caption = Col(1)
lblCol2.Caption = Col(2)
lblCol3.Caption = Col(3)
lblCol4.Caption = Col(4)
lblCol5.Caption = Col(5)

End Sub

Private Sub Label1_Click()
Tutorial = Tutorial + 1
If Tutorial = 1 Then
    Shape1.Visible = True
    lblInstructions.Caption = "The numbers above the grid tells you how many boxes you have to draw in the downward direction. For example:"
End If

If Tutorial = 2 Then
    Shape1.Width = 330
    Shape1.Left = 810
    lblInstructions.Caption = "The first column has a 3 over it. That means there are 3 blocks that you have to click on in that column."
    Square(11).Picture = LoadPicture(App.Path & "\true.gif")
    Square(16).Picture = LoadPicture(App.Path & "\true.gif")
    Square(21).Picture = LoadPicture(App.Path & "\true.gif")
End If

If Tutorial = 3 Then
    Shape1.Visible = False
    Shape2.Visible = True
    lblInstructions.Caption = "The numbers on the left side of the grid tells you how many boxes you have to draw in a left to right direction. For Example:"
End If

If Tutorial = 4 Then
    Shape2.Top = 2070
    Shape2.Height = 375
    lblInstructions.Caption = "The fourth row from the top has a 5 next to it. That means there are 5 blocks you have to click on in that row."
    Square(16).Picture = LoadPicture(App.Path & "\true.gif")
    Square(17).Picture = LoadPicture(App.Path & "\true.gif")
    Square(18).Picture = LoadPicture(App.Path & "\true.gif")
    Square(19).Picture = LoadPicture(App.Path & "\true.gif")
    Square(20).Picture = LoadPicture(App.Path & "\true.gif")
End If

If Tutorial = 5 Then
    Shape2.Top = 1485
    Shape2.Height = 375
    lblInstructions.Caption = "This row has a 1 and a 1 next to it, so that means you click 1 block and then another block seperated by at least 1 space."
    Square(7).Picture = LoadPicture(App.Path & "\true.gif")
    Square(9).Picture = LoadPicture(App.Path & "\true.gif")
End If

If Tutorial = 6 Then
    Shape2.Visible = False
    Line5.Visible = True
    lblInstructions.Caption = "This is the time limit. For each stage you only have a certain amount of time to compete it."
    Timer1.Enabled = True
End If

If Tutorial = 7 Then
    lblInstructions.Caption = "If you click on a bad square your time will be reduced by a certain amount. If your time goes down to zero, the game is over. But don't panic! Take your time!"
    Square(1).Picture = LoadPicture(App.Path & "\false.gif")
    lblStatus.Caption = "Miss! -40 seconds!"
    Time = Time - 40
End If

If Tutorial = 8 Then
    Line5.Visible = False
    lblStatus.Caption = ""
    Timer1.Enabled = False
    lblInstructions.Caption = "That's it! You now know everything you need to start playing. Good Luck!"
    Label1.Enabled = False
End If

End Sub

Private Sub Timer1_Timer()

Time = Time - 1
lblTime.Caption = Time

End Sub
