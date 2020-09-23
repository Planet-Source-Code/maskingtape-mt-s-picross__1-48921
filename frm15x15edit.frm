VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm15x15edit 
   BackColor       =   &H00955800&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MT's Picross: 15x15 Editor"
   ClientHeight    =   8205
   ClientLeft      =   150
   ClientTop       =   465
   ClientWidth     =   6780
   Icon            =   "frm15x15edit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   6780
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
      Left            =   2205
      TabIndex        =   226
      Top             =   7065
      Width           =   2490
      Begin VB.TextBox txtDescription 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   90
         MaxLength       =   25
         TabIndex        =   227
         Text            =   "Stage"
         Top             =   315
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00955800&
      Height          =   5070
      Left            =   1800
      TabIndex        =   0
      Top             =   1935
      Width           =   4920
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   150
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   225
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   149
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   224
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   148
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   223
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   147
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   222
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   146
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   221
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   145
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   220
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   144
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   219
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   143
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   218
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   142
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   217
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   141
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   216
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   140
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   215
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   139
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   214
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   138
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   213
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   137
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   212
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   136
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   211
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   135
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   210
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   134
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   209
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   133
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   208
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   132
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   207
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   131
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   206
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   130
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   205
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   129
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   204
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   128
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   203
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   127
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   202
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   126
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   201
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   125
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   200
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   124
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   199
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   123
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   198
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   122
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   197
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   121
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   196
         Top             =   2730
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   120
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   195
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   119
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   194
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   118
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   193
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   117
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   192
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   116
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   191
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   115
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   190
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   114
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   189
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   113
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   188
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   112
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   187
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   111
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   186
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   110
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   185
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   108
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   184
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   109
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   183
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   107
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   182
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   106
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   181
         Top             =   2415
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   105
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   180
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   104
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   179
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   103
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   178
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   102
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   177
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   101
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   176
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   25
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   175
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   24
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   174
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   23
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   173
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   22
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   172
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   21
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   171
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   20
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   170
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   19
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   169
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   18
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   168
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   17
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   167
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   16
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   166
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   15
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   165
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   14
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   164
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   13
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   163
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   12
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   162
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   11
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   161
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   160
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   159
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   158
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   157
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   156
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   155
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   154
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   153
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   152
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   151
         Top             =   210
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   26
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   150
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   27
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   149
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   28
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   148
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   29
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   147
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   30
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   146
         Top             =   525
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   31
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   145
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   32
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   144
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   33
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   143
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   34
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   142
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   35
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   141
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   36
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   140
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   37
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   139
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   38
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   138
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   39
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   137
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   40
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   136
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   41
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   135
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   42
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   134
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   43
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   133
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   44
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   132
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   45
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   131
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   46
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   130
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   47
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   129
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   48
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   128
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   49
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   127
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   50
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   126
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   51
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   125
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   52
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   124
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   53
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   123
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   54
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   122
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   55
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   121
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   56
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   120
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   57
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   119
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   58
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   118
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   59
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   117
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   60
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   116
         Top             =   1155
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   61
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   115
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   62
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   114
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   63
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   113
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   64
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   112
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   65
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   111
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   66
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   110
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   67
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   109
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   68
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   108
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   69
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   107
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   70
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   106
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   71
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   105
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   72
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   104
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   73
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   103
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   74
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   102
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   75
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   101
         Top             =   1470
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   76
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   100
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   77
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   99
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   78
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   98
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   79
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   97
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   80
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   96
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   81
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   95
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   82
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   94
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   83
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   93
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   84
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   92
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   85
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   91
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   86
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   90
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   87
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   89
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   88
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   88
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   89
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   87
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   90
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   86
         Top             =   1785
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   91
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   85
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   92
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   84
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   93
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   83
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   94
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   82
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   95
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   81
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   96
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   80
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   97
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   79
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   98
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   78
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   99
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   77
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   100
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   76
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   151
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   75
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   152
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   74
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   153
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   73
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   154
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   72
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   155
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   71
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   156
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   70
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   157
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   69
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   158
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   68
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   159
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   67
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   160
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   66
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   161
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   65
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   162
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   64
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   163
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   63
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   164
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   62
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   165
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   61
         Top             =   3360
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   166
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   60
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   167
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   59
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   168
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   58
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   169
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   57
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   170
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   56
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   171
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   55
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   172
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   54
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   173
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   53
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   174
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   52
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   175
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   51
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   176
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   50
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   177
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   49
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   178
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   48
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   179
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   47
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   180
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   46
         Top             =   3675
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   181
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   45
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   182
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   44
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   183
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   43
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   184
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   42
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   185
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   41
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   186
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   40
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   187
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   39
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   188
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   38
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   189
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   37
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   190
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   36
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   191
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   35
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   192
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   34
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   193
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   33
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   194
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   32
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   195
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   31
         Top             =   3990
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   196
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   30
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   197
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   29
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   198
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   28
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   199
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   27
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   200
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   26
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   201
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   25
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   202
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   24
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   203
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   23
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   204
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   22
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   205
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   21
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   206
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   20
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   207
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   19
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   208
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   18
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   209
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   17
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   210
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   16
         Top             =   4305
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   211
         Left            =   105
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   15
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   212
         Left            =   420
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   14
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   213
         Left            =   735
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   13
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   214
         Left            =   1050
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   12
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   215
         Left            =   1365
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   11
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   216
         Left            =   1680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   10
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   217
         Left            =   1995
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   9
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   218
         Left            =   2310
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   8
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   219
         Left            =   2625
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   7
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   220
         Left            =   2940
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   6
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   221
         Left            =   3255
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   5
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   222
         Left            =   3570
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   4
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   223
         Left            =   3885
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   3
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   224
         Left            =   4200
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   2
         Top             =   4620
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   225
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   1
         Top             =   4620
         Width           =   300
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   420
      Top             =   630
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.pcs|*.pcs"
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
      Left            =   0
      TabIndex        =   258
      Top             =   7920
      Width           =   6765
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      X1              =   -405
      X2              =   6930
      Y1              =   7875
      Y2              =   7875
   End
   Begin VB.Image Image31 
      Height          =   285
      Left            =   -90
      Picture         =   "frm15x15edit.frx":1272
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   7065
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   -330
      Y1              =   6930
      Y2              =   6930
   End
   Begin VB.Line Line1 
      X1              =   1845
      X2              =   -540
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label lblRow10 
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
      Left            =   -330
      TabIndex        =   257
      Top             =   4950
      Width           =   1935
   End
   Begin VB.Label lblRow9 
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
      Left            =   -330
      TabIndex        =   256
      Top             =   4635
      Width           =   1935
   End
   Begin VB.Label lblRow8 
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
      Left            =   -330
      TabIndex        =   255
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblRow7 
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
      Left            =   -330
      TabIndex        =   254
      Top             =   4005
      Width           =   1935
   End
   Begin VB.Label lblRow6 
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
      Left            =   -330
      TabIndex        =   253
      Top             =   3690
      Width           =   1935
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
      Left            =   -330
      TabIndex        =   252
      Top             =   2115
      Width           =   1935
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
      Left            =   -330
      TabIndex        =   251
      Top             =   2430
      Width           =   1935
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
      Left            =   -330
      TabIndex        =   250
      Top             =   2745
      Width           =   1935
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
      Left            =   -330
      TabIndex        =   249
      Top             =   3060
      Width           =   1935
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
      Left            =   -330
      TabIndex        =   248
      Top             =   3375
      Width           =   1935
   End
   Begin VB.Label lblRow11 
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
      Left            =   -330
      TabIndex        =   247
      Top             =   5265
      Width           =   1935
   End
   Begin VB.Label lblRow12 
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
      Left            =   -330
      TabIndex        =   246
      Top             =   5580
      Width           =   1935
   End
   Begin VB.Label lblRow13 
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
      Left            =   -330
      TabIndex        =   245
      Top             =   5895
      Width           =   1935
   End
   Begin VB.Label lblRow14 
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
      Left            =   -330
      TabIndex        =   244
      Top             =   6210
      Width           =   1935
   End
   Begin VB.Label lblRow15 
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
      Left            =   -330
      TabIndex        =   243
      Top             =   6525
      Width           =   1935
   End
   Begin VB.Line Line4 
      X1              =   6615
      X2              =   6615
      Y1              =   -90
      Y2              =   1935
   End
   Begin VB.Line Line3 
      X1              =   1890
      X2              =   1890
      Y1              =   -90
      Y2              =   1920
   End
   Begin VB.Label lblCol10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2145
      Left            =   4725
      TabIndex        =   242
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   4410
      TabIndex        =   241
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   4095
      TabIndex        =   240
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   3780
      TabIndex        =   239
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   3465
      TabIndex        =   238
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   3150
      TabIndex        =   237
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   2835
      TabIndex        =   236
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2100
      Left            =   2520
      TabIndex        =   235
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCOl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   2205
      TabIndex        =   234
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2190
      Left            =   1890
      TabIndex        =   233
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2145
      Left            =   5040
      TabIndex        =   232
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2100
      Left            =   5355
      TabIndex        =   231
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2100
      Left            =   5670
      TabIndex        =   230
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2100
      Left            =   5985
      TabIndex        =   229
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCol15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   6300
      TabIndex        =   228
      Top             =   15
      Width           =   330
   End
   Begin VB.Image Image11 
      Height          =   2040
      Left            =   1935
      Picture         =   "frm15x15edit.frx":19BC
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   2040
      Left            =   2250
      Picture         =   "frm15x15edit.frx":2328
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   2040
      Left            =   2565
      Picture         =   "frm15x15edit.frx":2C94
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   2040
      Left            =   2880
      Picture         =   "frm15x15edit.frx":3600
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   2040
      Left            =   3195
      Picture         =   "frm15x15edit.frx":3F6C
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   2040
      Left            =   3510
      Picture         =   "frm15x15edit.frx":48D8
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   2040
      Left            =   3825
      Picture         =   "frm15x15edit.frx":5244
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   2040
      Left            =   4140
      Picture         =   "frm15x15edit.frx":5BB0
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   2040
      Left            =   4455
      Picture         =   "frm15x15edit.frx":651C
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image9 
      Height          =   2040
      Left            =   4770
      Picture         =   "frm15x15edit.frx":6E88
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image10 
      Height          =   2040
      Left            =   5085
      Picture         =   "frm15x15edit.frx":77F4
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image12 
      Height          =   2040
      Left            =   5400
      Picture         =   "frm15x15edit.frx":8160
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image13 
      Height          =   2040
      Left            =   5715
      Picture         =   "frm15x15edit.frx":8ACC
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image14 
      Height          =   2040
      Left            =   6030
      Picture         =   "frm15x15edit.frx":9438
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image15 
      Height          =   2040
      Left            =   6345
      Picture         =   "frm15x15edit.frx":9DA4
      Stretch         =   -1  'True
      Top             =   15
      Width           =   240
   End
   Begin VB.Image Image16 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":A710
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2040
   End
   Begin VB.Image Image17 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":AF80
      Stretch         =   -1  'True
      Top             =   2475
      Width           =   2040
   End
   Begin VB.Image Image18 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":B7F0
      Stretch         =   -1  'True
      Top             =   2790
      Width           =   2040
   End
   Begin VB.Image Image19 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":C060
      Stretch         =   -1  'True
      Top             =   3105
      Width           =   2040
   End
   Begin VB.Image Image20 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":C8D0
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   2040
   End
   Begin VB.Image Image21 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":D140
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   2040
   End
   Begin VB.Image Image22 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":D9B0
      Stretch         =   -1  'True
      Top             =   4050
      Width           =   2040
   End
   Begin VB.Image Image23 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":E220
      Stretch         =   -1  'True
      Top             =   4365
      Width           =   2040
   End
   Begin VB.Image Image24 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":EA90
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   2040
   End
   Begin VB.Image Image25 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":F300
      Stretch         =   -1  'True
      Top             =   4995
      Width           =   2040
   End
   Begin VB.Image Image26 
      Height          =   285
      Left            =   -270
      Picture         =   "frm15x15edit.frx":FB70
      Stretch         =   -1  'True
      Top             =   5310
      Width           =   2040
   End
   Begin VB.Image Image27 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":103E0
      Stretch         =   -1  'True
      Top             =   5625
      Width           =   2040
   End
   Begin VB.Image Image28 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":10C50
      Stretch         =   -1  'True
      Top             =   5940
      Width           =   2040
   End
   Begin VB.Image Image29 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":114C0
      Stretch         =   -1  'True
      Top             =   6255
      Width           =   2040
   End
   Begin VB.Image Image30 
      Height          =   285
      Left            =   -285
      Picture         =   "frm15x15edit.frx":11D30
      Stretch         =   -1  'True
      Top             =   6570
      Width           =   2040
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "Open.."
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "Save As.."
         Shortcut        =   ^S
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Index           =   0
      Begin VB.Menu clear 
         Caption         =   "Clear All"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu protect 
         Caption         =   "Protect"
      End
   End
End
Attribute VB_Name = "frm15x15edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clear_Click()

Dim PicNum As Integer
PicNum = 1

Do While PicNum < 226
    Square(PicNum).Picture = LoadPicture()
    PicEDIT(PicNum) = False
    PicNum = PicNum + 1
Loop

lblRow1.Caption = "0"
lblRow2.Caption = "0"
lblRow3.Caption = "0"
lblRow4.Caption = "0"
lblRow5.Caption = "0"
lblRow6.Caption = "0"
lblRow7.Caption = "0"
lblRow8.Caption = "0"
lblRow9.Caption = "0"
lblRow10.Caption = "0"
lblRow11.Caption = "0"
lblRow12.Caption = "0"
lblRow13.Caption = "0"
lblRow14.Caption = "0"
lblRow15.Caption = "0"

lblCol1.Caption = "0"
lblCOl2.Caption = "0"
lblCol3.Caption = "0"
lblCol4.Caption = "0"
lblCol5.Caption = "0"
lblCol6.Caption = "0"
lblCol7.Caption = "0"
lblCol8.Caption = "0"
lblCol9.Caption = "0"
lblCol10.Caption = "0"
lblCol11.Caption = "0"
lblCol12.Caption = "0"
lblCol13.Caption = "0"
lblCol14.Caption = "0"
lblCol15.Caption = "0"

txtDescription.Text = "Stage"

End Sub

Private Sub exit_Click()

frmMain.Show
Unload Me

End Sub

Private Sub Form_Load()

Protected = False
Dim tmp As Integer
tmp = 1

Do While tmp < 226
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
Dim Tmp2
tmp = 1

CommonDialog1.ShowOpen
Open CommonDialog1.FileName For Input As #1
Input #1, StageSize, Pic(1), Pic(2), Pic(3), Pic(4), Pic(5), Pic(6), Pic(7), Pic(8), Pic(9), Pic(10), Pic(11), Pic(12), Pic(13), Pic(14), Pic(15), Pic(16), Pic(17), Pic(18), Pic(19), Pic(20), Pic(21), Pic(22), Pic(23), Pic(24), Pic(25), Pic(26), Pic(27), Pic(28), Pic(29), Pic(30), Pic(31), Pic(32), Pic(33), Pic(34), Pic(35), Pic(36), Pic(37), Pic(38), Pic(39), Pic(40), Pic(41), Pic(42), Pic(43), Pic(44), Pic(45), Pic(46), Pic(47), Pic(48), Pic(49), Pic(50)
Input #1, Pic(51), Pic(52), Pic(53), Pic(54), Pic(55), Pic(56), Pic(57), Pic(58), Pic(59), Pic(60), Pic(61), Pic(62), Pic(63), Pic(64), Pic(65), Pic(66), Pic(67), Pic(68), Pic(69), Pic(70), Pic(71), Pic(72), Pic(73), Pic(74), Pic(75), Pic(76), Pic(77), Pic(78), Pic(79), Pic(80), Pic(81), Pic(82), Pic(83), Pic(84), Pic(85), Pic(86), Pic(87), Pic(88), Pic(89), Pic(90), Pic(91), Pic(92), Pic(93), Pic(94), Pic(95), Pic(96), Pic(97), Pic(98), Pic(99), Pic(100)
Input #1, Pic(101), Pic(102), Pic(103), Pic(104), Pic(105), Pic(106), Pic(107), Pic(108), Pic(109), Pic(110), Pic(111), Pic(112), Pic(113), Pic(114), Pic(115), Pic(116), Pic(117), Pic(118), Pic(119), Pic(120), Pic(121), Pic(122), Pic(123), Pic(124), Pic(125), Pic(126), Pic(127), Pic(128), Pic(129), Pic(130), Pic(131), Pic(132), Pic(133), Pic(134), Pic(135), Pic(136), Pic(137), Pic(138), Pic(139), Pic(140), Pic(141), Pic(142), Pic(143), Pic(144), Pic(145), Pic(146), Pic(147), Pic(148), Pic(149), Pic(150)
Input #1, Pic(151), Pic(152), Pic(153), Pic(154), Pic(155), Pic(156), Pic(157), Pic(158), Pic(159), Pic(160), Pic(161), Pic(162), Pic(163), Pic(164), Pic(165), Pic(166), Pic(167), Pic(168), Pic(169), Pic(170), Pic(171), Pic(172), Pic(173), Pic(174), Pic(175), Pic(176), Pic(177), Pic(178), Pic(179), Pic(180), Pic(181), Pic(182), Pic(183), Pic(184), Pic(185), Pic(186), Pic(187), Pic(188), Pic(189), Pic(190), Pic(191), Pic(192), Pic(193), Pic(194), Pic(195), Pic(196), Pic(197), Pic(198), Pic(199), Pic(200)
Input #1, Pic(201), Pic(202), Pic(203), Pic(204), Pic(205), Pic(206), Pic(207), Pic(208), Pic(209), Pic(210), Pic(211), Pic(212), Pic(213), Pic(214), Pic(215), Pic(216), Pic(217), Pic(218), Pic(219), Pic(220), Pic(221), Pic(222), Pic(223), Pic(224), Pic(225)
Input #1, Row(1), Row(2), Row(3), Row(4), Row(5), Row(6), Row(7), Row(8), Row(9), Row(10), Row(11), Row(12), Row(13), Row(14), Row(15), Col(1), Col(2), Col(3), Col(4), Col(5), Col(6), Col(7), Col(8), Col(9), Col(10), Col(11), Col(12), Col(13), Col(14), Col(15), Description, Tmp2
Close #1

If StageSize = "15x15" And Tmp2 = False Then
    
    Do While tmp < 226
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
    lblRow6.Caption = Row(6)
    lblRow7.Caption = Row(7)
    lblRow8.Caption = Row(8)
    lblRow9.Caption = Row(9)
    lblRow10.Caption = Row(10)
    lblRow11.Caption = Row(11)
    lblRow12.Caption = Row(12)
    lblRow13.Caption = Row(13)
    lblRow14.Caption = Row(14)
    lblRow15.Caption = Row(15)
     
    lblCol1.Caption = Col(1)
    lblCOl2.Caption = Col(2)
    lblCol3.Caption = Col(3)
    lblCol4.Caption = Col(4)
    lblCol5.Caption = Col(5)
    lblCol6.Caption = Col(6)
    lblCol7.Caption = Col(7)
    lblCol8.Caption = Col(8)
    lblCol9.Caption = Col(9)
    lblCol10.Caption = Col(10)
    lblCol11.Caption = Col(11)
    lblCol12.Caption = Col(12)
    lblCol13.Caption = Col(13)
    lblCol14.Caption = Col(14)
    lblCol15.Caption = Col(15)

    
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
Write #1, "15x15", PicEDIT(1), PicEDIT(2), PicEDIT(3), PicEDIT(4), PicEDIT(5), PicEDIT(6), PicEDIT(7), PicEDIT(8), PicEDIT(9), PicEDIT(10), PicEDIT(11), PicEDIT(12), PicEDIT(13), PicEDIT(14), PicEDIT(15), PicEDIT(16), PicEDIT(17), PicEDIT(18), PicEDIT(19), PicEDIT(20), PicEDIT(21), PicEDIT(22), PicEDIT(23), PicEDIT(24), PicEDIT(25), PicEDIT(26), PicEDIT(27), PicEDIT(28), PicEDIT(29), PicEDIT(30), PicEDIT(31), PicEDIT(32), PicEDIT(33), PicEDIT(34), PicEDIT(35), PicEDIT(36), PicEDIT(37), PicEDIT(38), PicEDIT(39), PicEDIT(40), PicEDIT(41), PicEDIT(42), PicEDIT(43), PicEDIT(44), PicEDIT(45), PicEDIT(46), PicEDIT(47), PicEDIT(48), PicEDIT(49), PicEDIT(50)
Write #1, PicEDIT(51), PicEDIT(52), PicEDIT(53), PicEDIT(54), PicEDIT(55), PicEDIT(56), PicEDIT(57), PicEDIT(58), PicEDIT(59), PicEDIT(60), PicEDIT(61), PicEDIT(62), PicEDIT(63), PicEDIT(64), PicEDIT(65), PicEDIT(66), PicEDIT(67), PicEDIT(68), PicEDIT(69), PicEDIT(70), PicEDIT(71), PicEDIT(72), PicEDIT(73), PicEDIT(74), PicEDIT(75), PicEDIT(76), PicEDIT(77), PicEDIT(78), PicEDIT(79), PicEDIT(80), PicEDIT(81), PicEDIT(82), PicEDIT(83), PicEDIT(84), PicEDIT(85), PicEDIT(86), PicEDIT(87), PicEDIT(88), PicEDIT(89), PicEDIT(90), PicEDIT(91), PicEDIT(92), PicEDIT(93), PicEDIT(94), PicEDIT(95), PicEDIT(96), PicEDIT(97), PicEDIT(98), PicEDIT(99), PicEDIT(100)
Write #1, PicEDIT(101), PicEDIT(102), PicEDIT(103), PicEDIT(104), PicEDIT(105), PicEDIT(106), PicEDIT(107), PicEDIT(108), PicEDIT(109), PicEDIT(110), PicEDIT(111), PicEDIT(112), PicEDIT(113), PicEDIT(114), PicEDIT(115), PicEDIT(116), PicEDIT(117), PicEDIT(118), PicEDIT(119), PicEDIT(120), PicEDIT(121), PicEDIT(122), PicEDIT(123), PicEDIT(124), PicEDIT(125), PicEDIT(126), PicEDIT(127), PicEDIT(128), PicEDIT(129), PicEDIT(130), PicEDIT(131), PicEDIT(132), PicEDIT(133), PicEDIT(134), PicEDIT(135), PicEDIT(136), PicEDIT(137), PicEDIT(138), PicEDIT(139), PicEDIT(140), PicEDIT(141), PicEDIT(142), PicEDIT(143), PicEDIT(144), PicEDIT(145), PicEDIT(146), PicEDIT(147), PicEDIT(148), PicEDIT(149), PicEDIT(150)
Write #1, PicEDIT(151), PicEDIT(152), PicEDIT(153), PicEDIT(154), PicEDIT(155), PicEDIT(156), PicEDIT(157), PicEDIT(158), PicEDIT(159), PicEDIT(160), PicEDIT(161), PicEDIT(162), PicEDIT(163), PicEDIT(164), PicEDIT(165), PicEDIT(166), PicEDIT(167), PicEDIT(168), PicEDIT(169), PicEDIT(170), PicEDIT(171), PicEDIT(172), PicEDIT(173), PicEDIT(174), PicEDIT(175), PicEDIT(176), PicEDIT(177), PicEDIT(178), PicEDIT(179), PicEDIT(180), PicEDIT(181), PicEDIT(182), PicEDIT(183), PicEDIT(184), PicEDIT(185), PicEDIT(186), PicEDIT(187), PicEDIT(188), PicEDIT(189), PicEDIT(190), PicEDIT(191), PicEDIT(192), PicEDIT(193), PicEDIT(194), PicEDIT(195), PicEDIT(196), PicEDIT(197), PicEDIT(198), PicEDIT(199), PicEDIT(200)
Write #1, PicEDIT(201), PicEDIT(202), PicEDIT(203), PicEDIT(204), PicEDIT(205), PicEDIT(206), PicEDIT(207), PicEDIT(208), PicEDIT(209), PicEDIT(210), PicEDIT(211), PicEDIT(212), PicEDIT(213), PicEDIT(214), PicEDIT(215), PicEDIT(216), PicEDIT(217), PicEDIT(218), PicEDIT(219), PicEDIT(220), PicEDIT(221), PicEDIT(222), PicEDIT(223), PicEDIT(224), PicEDIT(225)
Write #1, lblRow1.Caption, lblRow2.Caption, lblRow3.Caption, lblRow4.Caption, lblRow5.Caption, lblRow6.Caption, lblRow7.Caption, lblRow8.Caption, lblRow9.Caption, lblRow10.Caption, lblRow11.Caption, lblRow12.Caption, lblRow13.Caption, lblRow14.Caption, lblRow15.Caption, lblCol1.Caption, lblCOl2.Caption, lblCol3.Caption, lblCol4.Caption, lblCol5.Caption, lblCol6.Caption, lblCol7.Caption, lblCol8.Caption, lblCol9.Caption, lblCol10.Caption, lblCol11.Caption, lblCol12.Caption, lblCol13.Caption, lblCol14.Caption, lblCol15.Caption, txtDescription.Text, Protected
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
Call CalcRow6
Call CalcRow7
Call CalcRow8
Call CalcRow9
Call CalcRow10
Call CalcRow11
Call CalcRow12
Call CalcRow13
Call CalcRow14
Call CalcRow15
Call CalcCol1
Call CalcCol2
Call CalcCol3
Call CalcCol4
Call CalcCol5
Call CalcCol6
Call CalcCol7
Call CalcCol8
Call CalcCol9
Call CalcCol10
Call CalcCol11
Call CalcCol12
Call CalcCol13
Call CalcCol14
Call CalcCol15
End Sub

Private Sub CalcRow1()

Dim RowCount(1 To 15) As Integer
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
Do Until tmp = 16
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

Dim RowCount(1 To 15) As Integer
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

Do While PicCount < 31
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
Do Until tmp = 16
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

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 31
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 46
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
Do Until tmp = 16
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

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 46
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 61
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
Do Until tmp = 16
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

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 61
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 76
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
Do Until tmp = 16
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

Private Sub CalcRow6()

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 76
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 91
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
Do Until tmp = 16
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

If RowEdit = "" Then lblRow6.Caption = "0" Else lblRow6.Caption = RowEdit

End Sub

Private Sub CalcRow7()

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 91
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 106
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
Do Until tmp = 16
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

If RowEdit = "" Then lblRow7.Caption = "0" Else lblRow7.Caption = RowEdit

End Sub

Private Sub CalcRow8()

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 106
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 121
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
Do Until tmp = 16
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

If RowEdit = "" Then lblRow8.Caption = "0" Else lblRow8.Caption = RowEdit

End Sub

Private Sub CalcRow9()

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 121
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 136
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
Do Until tmp = 16
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

If RowEdit = "" Then lblRow9.Caption = "0" Else lblRow9.Caption = RowEdit

End Sub

Private Sub CalcRow10()

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 136
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 151
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
Do Until tmp = 16
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

If RowEdit = "" Then lblRow10.Caption = "0" Else lblRow10.Caption = RowEdit

End Sub

Private Sub CalcRow11()

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 151
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 166
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
Do Until tmp = 16
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

If RowEdit = "" Then lblRow11.Caption = "0" Else lblRow11.Caption = RowEdit

End Sub

Private Sub CalcRow12()

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 166
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 181
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
Do Until tmp = 16
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

If RowEdit = "" Then lblRow12.Caption = "0" Else lblRow12.Caption = RowEdit

End Sub

Private Sub CalcRow13()

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 181
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 196
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
Do Until tmp = 16
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

If RowEdit = "" Then lblRow13.Caption = "0" Else lblRow13.Caption = RowEdit

End Sub

Private Sub CalcRow14()

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 196
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 211
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
Do Until tmp = 16
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

If RowEdit = "" Then lblRow14.Caption = "0" Else lblRow14.Caption = RowEdit

End Sub

Private Sub CalcRow15()

Dim RowCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 211
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 226
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
Do Until tmp = 16
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

If RowEdit = "" Then lblRow15.Caption = "0" Else lblRow15.Caption = RowEdit

End Sub

Private Sub CalcCol1()

Dim ColCount(1 To 15) As Integer
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

Do While PicCount < 212
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

Dim ColCount(1 To 15) As Integer
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

Do While PicCount < 213
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

Dim ColCount(1 To 15) As Integer
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

Do While PicCount < 214
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

Dim ColCount(1 To 15) As Integer
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

Do While PicCount < 215
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

Dim ColCount(1 To 15) As Integer
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

Do While PicCount < 216
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

Private Sub CalcCol6()

Dim ColCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 6
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 217
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

If ColEdit = "" Then lblCol6.Caption = 0 Else lblCol6.Caption = ColEdit
End Sub

Private Sub CalcCol7()

Dim ColCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 7
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 218
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

If ColEdit = "" Then lblCol7.Caption = 0 Else lblCol7.Caption = ColEdit

End Sub

Private Sub CalcCol8()

Dim ColCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 8
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 219
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

If ColEdit = "" Then lblCol8.Caption = 0 Else lblCol8.Caption = ColEdit

End Sub

Private Sub CalcCol9()

Dim ColCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 9
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 220
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

If ColEdit = "" Then lblCol9.Caption = 0 Else lblCol9.Caption = ColEdit

End Sub

Private Sub CalcCol10()

Dim ColCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 10
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 221
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

If ColEdit = "" Then lblCol10.Caption = 0 Else lblCol10.Caption = ColEdit

End Sub

Private Sub CalcCol11()

Dim ColCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 11
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 222
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

If ColEdit = "" Then lblCol11.Caption = 0 Else lblCol11.Caption = ColEdit

End Sub

Private Sub CalcCol12()

Dim ColCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 12
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 223
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

If ColEdit = "" Then lblCol12.Caption = 0 Else lblCol12.Caption = ColEdit

End Sub

Private Sub CalcCol13()

Dim ColCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 13
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 224
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

If ColEdit = "" Then lblCol13.Caption = 0 Else lblCol13.Caption = ColEdit

End Sub

Private Sub CalcCol14()

Dim ColCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 14
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 225
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

If ColEdit = "" Then lblCol14.Caption = 0 Else lblCol14.Caption = ColEdit

End Sub

Private Sub CalcCol15()

Dim ColCount(1 To 15) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 15
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 226
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 15
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 15
    End If
Loop

tmp = 1
Do Until tmp = 16
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

If ColEdit = "" Then lblCol15.Caption = 0 Else lblCol15.Caption = ColEdit

End Sub


