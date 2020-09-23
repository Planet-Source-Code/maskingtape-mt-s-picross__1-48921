VERSION 5.00
Begin VB.Form frm5x5 
   BackColor       =   &H00955800&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MT's Picross: 5x5 PlayField"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   2790
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm5x5.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   186
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   180
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00955800&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   675
      TabIndex        =   0
      Top             =   990
      Width           =   1815
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   25
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   33
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   24
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   32
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   23
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   31
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   22
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   30
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   21
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   29
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   20
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   28
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   19
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   27
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   18
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   26
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   17
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   25
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   16
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   16
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   15
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   15
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   14
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   14
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   12
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   12
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   11
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   11
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   10
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   9
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   8
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   7
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   6
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   5
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   4
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   3
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   2
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   1
         Top             =   225
         Width           =   300
      End
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   186
      Y1              =   198
      Y2              =   198
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   147
      X2              =   147
      Y1              =   216
      Y2              =   201
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
      Left            =   0
      TabIndex        =   37
      Top             =   3015
      Width           =   2175
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
      Left            =   2205
      TabIndex        =   36
      Top             =   3015
      Width           =   555
   End
   Begin VB.Image Image12 
      Height          =   375
      Left            =   0
      Picture         =   "frm5x5.frx":1272
      Stretch         =   -1  'True
      Top             =   2970
      Width           =   2790
   End
   Begin VB.Image imgExit 
      Height          =   450
      Left            =   90
      Picture         =   "frm5x5.frx":19BC
      Top             =   360
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Line Line4 
      X1              =   39
      X2              =   0
      Y1              =   189
      Y2              =   189
   End
   Begin VB.Line Line3 
      X1              =   42
      X2              =   0
      Y1              =   81
      Y2              =   81
   End
   Begin VB.Line Line2 
      X1              =   159
      X2              =   159
      Y1              =   0
      Y2              =   69
   End
   Begin VB.Line Line1 
      X1              =   54
      X2              =   54
      Y1              =   0
      Y2              =   69
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
      TabIndex        =   35
      Top             =   45
      Width           =   330
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
      TabIndex        =   34
      Top             =   2475
      Width           =   555
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
      TabIndex        =   24
      Top             =   45
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
      TabIndex        =   23
      Top             =   45
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
      TabIndex        =   22
      Top             =   45
      Width           =   330
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
      TabIndex        =   21
      Top             =   45
      Width           =   330
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
      TabIndex        =   20
      Top             =   2160
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
      TabIndex        =   19
      Top             =   1845
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
      TabIndex        =   18
      Top             =   1530
      Width           =   555
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
      TabIndex        =   17
      Top             =   1215
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   0
      Picture         =   "frm5x5.frx":2052
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   0
      Picture         =   "frm5x5.frx":23D6
      Stretch         =   -1  'True
      Top             =   1575
      Width           =   645
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   0
      Picture         =   "frm5x5.frx":275A
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   0
      Picture         =   "frm5x5.frx":2ADE
      Stretch         =   -1  'True
      Top             =   2205
      Width           =   645
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   0
      Picture         =   "frm5x5.frx":2E62
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   645
   End
   Begin VB.Image Image6 
      Height          =   1005
      Left            =   855
      Picture         =   "frm5x5.frx":31E6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   1005
      Left            =   1170
      Picture         =   "frm5x5.frx":3967
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   1005
      Left            =   1485
      Picture         =   "frm5x5.frx":40E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image9 
      Height          =   1005
      Left            =   1800
      Picture         =   "frm5x5.frx":4869
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image10 
      Height          =   1005
      Left            =   2115
      Picture         =   "frm5x5.frx":4FEA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Menu game 
      Caption         =   "&Game"
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frm5x5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub exit_Click()

If Custom = True Then
    Custom = False
    frmMain.Show
    Unload Me
ElseIf NetGame = True Then
    NetGame = False
    frmNetLevels.Show
    Kill Stage
    Unload Me
Else
    Call EasyRefresh
    frm5x5Select.Show
    Unload Me
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 1 And KeyCode = vbKeyB Then Call Win
End Sub

Private Sub Form_Load()

IsTrue = 0
Time = 600
lblTime.Caption = Time

Open Stage For Input As #1
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

Call CalcTrue

rc = MsgBox("Would you like a starting hint?", vbYesNo, "Hint?")
If rc = vbYes Then
    Call StartHint
Else
End If

End Sub

Private Sub imgExit_Click()

If Custom = True Then
    Custom = False
    frmMain.Show
    Unload Me
ElseIf NetGame = True Then
    NetGame = False
    frmNetLevels.Show
    Kill Stage
    Unload Me
Else
    Call EasyRefresh
    frm5x5Select.Show
    Unload Me
End If

End Sub

Private Sub Square_Click(Index As Integer)

If Pic(Index) = True Then
    Square(Index).Picture = LoadPicture(App.Path & "\true.gif")
    lblStatus.Caption = ""
    IsTrue = IsTrue + 1
    If IsTrue = NumTrue Then
        Timer1.Enabled = False
        lblStatus.Caption = "You Win! It's " & Description & "!"
        Frame1.Enabled = False
        Call Win
    End If
ElseIf Pic(Index) = False Then
    Square(Index).Picture = LoadPicture(App.Path & "\false.gif")
    Time = Time - 120
    lblStatus.Caption = "Miss! -120 seconds!"
    lblTime.Caption = Time
    If Time < 0 Then
        Frame1.Enabled = False
        Timer1.Enabled = False
        lblStatus.Caption = "Time's Up! Game Over!"
        lblTime.Caption = "0"
        imgExit.Visible = True
    End If
End If

Square(Index).Enabled = False

End Sub

Private Sub CalcTrue()
Dim PicNum As Integer
NumTrue = 0
PicNum = 1

Do While PicNum < 26
    If Pic(PicNum) = True Then NumTrue = NumTrue + 1
    PicNum = PicNum + 1
Loop

End Sub

Private Sub Win()

Dim PicNum As Integer

PicNum = 1

Do While PicNum < 26
    If Pic(PicNum) = False Then
        Square(PicNum).BackColor = &H955800
        Square(PicNum).Picture = LoadPicture()
        Square(PicNum).BorderStyle = 0
    ElseIf Pic(PicNum) = True Then
        Square(PicNum).Picture = LoadPicture(App.Path & "\true.gif")
    End If
    Sleep 100
    PicNum = PicNum + 1
    DoEvents
Loop

If Custom = False And NetGame = False Then
    EasyStagePass(StageNum) = True
    If EasyStageTime(StageNum) < Time Then
        EasyStageTime(StageNum) = Time
    ElseIf EasyStageTime(StageNum) > Time Then
    End If

    Open file For Output As #1
    Write #1, EasyStagePass(1), EasyStageTime(1), EasyStagePass(2), EasyStageTime(2), EasyStagePass(3), EasyStageTime(3), EasyStagePass(4), EasyStageTime(4), EasyStagePass(5), EasyStageTime(5), EasyStagePass(6), EasyStageTime(6), EasyStagePass(7), EasyStageTime(7), EasyStagePass(8), EasyStageTime(8), EasyStagePass(9), EasyStageTime(9), EasyStagePass(10), EasyStageTime(10)
    Close #1
End If

imgExit.Visible = True

End Sub

Private Sub Timer1_Timer()

If Time = 0 Then
    Frame1.Enabled = False
    Timer1.Enabled = False
    lblTime.Caption = "Time's Up! Game Over!"
    imgExit.Visible = True
Else
    Time = Time - 1
    lblTime.Caption = Time
End If

End Sub

Private Sub StartHint()

Dim Random As Integer
Dim Random2 As Integer
Dim tmp As Integer
Dim Tmp2 As Integer
Dim Tmp3 As Integer
Dim Tmp4 As Integer
Randomize

Random = Int((5 * Rnd) + 1)
Random2 = Int((5 * Rnd) + 1)

If Random = 1 Then tmp = 1: Tmp3 = 5
If Random = 2 Then tmp = 6: Tmp3 = 10
If Random = 3 Then tmp = 11: Tmp3 = 15
If Random = 4 Then tmp = 16: Tmp3 = 20
If Random = 5 Then tmp = 21: Tmp3 = 25

If Random2 = 1 Then Tmp2 = 1: Tmp4 = 21
If Random2 = 2 Then Tmp2 = 2: Tmp4 = 22
If Random2 = 3 Then Tmp2 = 3: Tmp4 = 23
If Random2 = 4 Then Tmp2 = 4: Tmp4 = 24
If Random2 = 5 Then Tmp2 = 5: Tmp4 = 25

Do While tmp <= Tmp3
    If Pic(tmp) = True And Square(tmp).Enabled = True Then
        Square(tmp).Picture = LoadPicture(App.Path & "\true.gif")
        Square(tmp).Enabled = False
        IsTrue = IsTrue + 1
    ElseIf Pic(tmp) = False Then
        Square(tmp).Picture = LoadPicture(App.Path & "\false.gif")
        Square(tmp).Enabled = False
    End If
tmp = tmp + 1
Loop

Do While Tmp2 <= Tmp4
    If Pic(Tmp2) = True And Square(Tmp2).Enabled = True Then
        Square(Tmp2).Picture = LoadPicture(App.Path & "\true.gif")
        Square(Tmp2).Enabled = False
        IsTrue = IsTrue + 1
    ElseIf Pic(Tmp2) = False Then
        Square(Tmp2).Picture = LoadPicture(App.Path & "\false.gif")
        Square(Tmp2).Enabled = False
    End If
Tmp2 = Tmp2 + 5
Loop

End Sub
