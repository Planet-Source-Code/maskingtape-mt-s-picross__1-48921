VERSION 5.00
Begin VB.Form frm15x15 
   BackColor       =   &H00955800&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MT's Picross: 15x15 Playfield"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7215
   ControlBox      =   0   'False
   Icon            =   "frm15x15.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1470
      Top             =   315
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00955800&
      Height          =   5070
      Left            =   2100
      TabIndex        =   0
      Top             =   1995
      Width           =   4920
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   225
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   225
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
         TabIndex        =   224
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
         TabIndex        =   223
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
         TabIndex        =   222
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
         TabIndex        =   221
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
         TabIndex        =   220
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
         TabIndex        =   219
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
         TabIndex        =   218
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
         TabIndex        =   217
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
         TabIndex        =   216
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
         TabIndex        =   215
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
         TabIndex        =   214
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
         TabIndex        =   213
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
         TabIndex        =   212
         Top             =   4620
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
         TabIndex        =   211
         Top             =   4620
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
         TabIndex        =   210
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
         TabIndex        =   209
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
         TabIndex        =   208
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
         TabIndex        =   207
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
         TabIndex        =   206
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
         TabIndex        =   205
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
         TabIndex        =   204
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
         TabIndex        =   203
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
         TabIndex        =   202
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
         TabIndex        =   201
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
         TabIndex        =   200
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
         TabIndex        =   199
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
         TabIndex        =   198
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
         TabIndex        =   197
         Top             =   4305
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
         TabIndex        =   196
         Top             =   4305
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
         TabIndex        =   195
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
         TabIndex        =   194
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
         TabIndex        =   193
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
         TabIndex        =   192
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
         TabIndex        =   191
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
         TabIndex        =   190
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
         TabIndex        =   189
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
         TabIndex        =   188
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
         TabIndex        =   187
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
         TabIndex        =   186
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
         TabIndex        =   185
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
         TabIndex        =   184
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
         TabIndex        =   183
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
         TabIndex        =   182
         Top             =   3990
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
         TabIndex        =   181
         Top             =   3990
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
         TabIndex        =   180
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
         TabIndex        =   179
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
         TabIndex        =   178
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
         TabIndex        =   177
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
         TabIndex        =   176
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
         TabIndex        =   175
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
         TabIndex        =   174
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
         TabIndex        =   173
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
         TabIndex        =   172
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
         TabIndex        =   171
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
         TabIndex        =   170
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
         TabIndex        =   169
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
         TabIndex        =   168
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
         TabIndex        =   167
         Top             =   3675
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
         TabIndex        =   166
         Top             =   3675
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
         TabIndex        =   165
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
         TabIndex        =   164
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
         TabIndex        =   163
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
         TabIndex        =   162
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
         TabIndex        =   161
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
         TabIndex        =   160
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
         TabIndex        =   159
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
         TabIndex        =   158
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
         TabIndex        =   157
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
         TabIndex        =   156
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
         TabIndex        =   155
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
         TabIndex        =   154
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
         TabIndex        =   153
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
         TabIndex        =   152
         Top             =   3360
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
         TabIndex        =   151
         Top             =   3360
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
         TabIndex        =   150
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
         TabIndex        =   149
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
         TabIndex        =   148
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
         TabIndex        =   147
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
         TabIndex        =   146
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
         TabIndex        =   145
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
         TabIndex        =   144
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
         TabIndex        =   143
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
         TabIndex        =   142
         Top             =   2100
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
         TabIndex        =   141
         Top             =   2100
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
         TabIndex        =   140
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
         TabIndex        =   139
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
         TabIndex        =   138
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
         TabIndex        =   137
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         TabIndex        =   134
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
         TabIndex        =   133
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   127
         Top             =   1785
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
         TabIndex        =   126
         Top             =   1785
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   117
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         Index           =   62
         Left            =   420
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
         Index           =   61
         Left            =   105
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
         Index           =   60
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   110
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
         TabIndex        =   109
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
         TabIndex        =   108
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
         Top             =   1155
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
         TabIndex        =   96
         Top             =   1155
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
         Top             =   840
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
         TabIndex        =   81
         Top             =   840
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
         Top             =   525
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
         TabIndex        =   76
         Top             =   525
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
         Top             =   210
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
         TabIndex        =   61
         Top             =   210
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   525
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
         TabIndex        =   51
         Top             =   525
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
         Top             =   2100
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
         TabIndex        =   46
         Top             =   2100
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   2415
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
         TabIndex        =   31
         Top             =   2415
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   2730
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
         TabIndex        =   16
         Top             =   2730
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   3045
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   150
         Left            =   4515
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   1
         Top             =   3045
         Width           =   300
      End
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   6570
      X2              =   6570
      Y1              =   7515
      Y2              =   7290
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
      TabIndex        =   257
      Top             =   7290
      Width           =   6450
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
      Left            =   6615
      TabIndex        =   256
      Top             =   7290
      Width           =   555
   End
   Begin VB.Image imgExit 
      Height          =   450
      Left            =   720
      Picture         =   "frm15x15.frx":1272
      Top             =   810
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Image31 
      Height          =   285
      Left            =   0
      Picture         =   "frm15x15.frx":1908
      Stretch         =   -1  'True
      Top             =   7290
      Width           =   7245
   End
   Begin VB.Label lblCol15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   6615
      TabIndex        =   255
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   6255
      TabIndex        =   254
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   5985
      TabIndex        =   253
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   5670
      TabIndex        =   252
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   5355
      TabIndex        =   251
      Top             =   0
      Width           =   330
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
      Left            =   0
      TabIndex        =   250
      Top             =   6615
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
      Left            =   0
      TabIndex        =   249
      Top             =   6300
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
      Left            =   0
      TabIndex        =   248
      Top             =   5985
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
      Left            =   0
      TabIndex        =   247
      Top             =   5670
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
      Left            =   0
      TabIndex        =   246
      Top             =   5355
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
      Left            =   0
      TabIndex        =   245
      Top             =   3465
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
      Left            =   0
      TabIndex        =   244
      Top             =   3150
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
      Left            =   0
      TabIndex        =   243
      Top             =   2835
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
      Left            =   0
      TabIndex        =   242
      Top             =   2520
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
      Left            =   0
      TabIndex        =   241
      Top             =   2205
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
      Left            =   0
      TabIndex        =   240
      Top             =   3780
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
      Left            =   0
      TabIndex        =   239
      Top             =   4095
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
      Left            =   0
      TabIndex        =   238
      Top             =   4410
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
      Left            =   0
      TabIndex        =   237
      Top             =   4725
      Width           =   1935
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
      Left            =   0
      TabIndex        =   236
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   -210
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Line Line2 
      X1              =   2115
      X2              =   0
      Y1              =   7020
      Y2              =   7020
   End
   Begin VB.Label lblCol1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   2205
      TabIndex        =   235
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCOl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   2520
      TabIndex        =   234
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   2835
      TabIndex        =   233
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   3150
      TabIndex        =   232
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   3465
      TabIndex        =   231
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   3780
      TabIndex        =   230
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   4095
      TabIndex        =   229
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   4410
      TabIndex        =   228
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   4725
      TabIndex        =   227
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1920
      Left            =   5040
      TabIndex        =   226
      Top             =   0
      Width           =   330
   End
   Begin VB.Line Line3 
      X1              =   2205
      X2              =   2205
      Y1              =   -105
      Y2              =   2115
   End
   Begin VB.Line Line4 
      X1              =   6930
      X2              =   6930
      Y1              =   -105
      Y2              =   2070
   End
   Begin VB.Image Image11 
      Height          =   2040
      Left            =   2250
      Picture         =   "frm15x15.frx":2052
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   2040
      Left            =   2565
      Picture         =   "frm15x15.frx":29BE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   2040
      Left            =   2880
      Picture         =   "frm15x15.frx":332A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   2040
      Left            =   3195
      Picture         =   "frm15x15.frx":3C96
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   2040
      Left            =   3510
      Picture         =   "frm15x15.frx":4602
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   2040
      Left            =   3825
      Picture         =   "frm15x15.frx":4F6E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image15 
      Height          =   2040
      Left            =   6660
      Picture         =   "frm15x15.frx":58DA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image14 
      Height          =   2040
      Left            =   6345
      Picture         =   "frm15x15.frx":6246
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image13 
      Height          =   2040
      Left            =   6030
      Picture         =   "frm15x15.frx":6BB2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image12 
      Height          =   2040
      Left            =   5715
      Picture         =   "frm15x15.frx":751E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image10 
      Height          =   2040
      Left            =   5400
      Picture         =   "frm15x15.frx":7E8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image9 
      Height          =   2040
      Left            =   5085
      Picture         =   "frm15x15.frx":87F6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   2040
      Left            =   4770
      Picture         =   "frm15x15.frx":9162
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   2040
      Left            =   4455
      Picture         =   "frm15x15.frx":9ACE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   2040
      Left            =   4140
      Picture         =   "frm15x15.frx":A43A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image16 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":ADA6
      Stretch         =   -1  'True
      Top             =   2250
      Width           =   2040
   End
   Begin VB.Image Image17 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":B616
      Stretch         =   -1  'True
      Top             =   2565
      Width           =   2040
   End
   Begin VB.Image Image18 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":BE86
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2040
   End
   Begin VB.Image Image19 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":C6F6
      Stretch         =   -1  'True
      Top             =   3195
      Width           =   2040
   End
   Begin VB.Image Image20 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":CF66
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   2040
   End
   Begin VB.Image Image21 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":D7D6
      Stretch         =   -1  'True
      Top             =   3825
      Width           =   2040
   End
   Begin VB.Image Image22 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":E046
      Stretch         =   -1  'True
      Top             =   4140
      Width           =   2040
   End
   Begin VB.Image Image23 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":E8B6
      Stretch         =   -1  'True
      Top             =   4455
      Width           =   2040
   End
   Begin VB.Image Image24 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":F126
      Stretch         =   -1  'True
      Top             =   4770
      Width           =   2040
   End
   Begin VB.Image Image25 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":F996
      Stretch         =   -1  'True
      Top             =   5085
      Width           =   2040
   End
   Begin VB.Image Image26 
      Height          =   285
      Left            =   90
      Picture         =   "frm15x15.frx":10206
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2040
   End
   Begin VB.Image Image27 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":10A76
      Stretch         =   -1  'True
      Top             =   5715
      Width           =   2040
   End
   Begin VB.Image Image28 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":112E6
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   2040
   End
   Begin VB.Image Image29 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":11B56
      Stretch         =   -1  'True
      Top             =   6345
      Width           =   2040
   End
   Begin VB.Image Image30 
      Height          =   285
      Left            =   45
      Picture         =   "frm15x15.frx":123C6
      Stretch         =   -1  'True
      Top             =   6660
      Width           =   2040
   End
   Begin VB.Menu game 
      Caption         =   "&Game"
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frm15x15"
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
    Call ExtraHardRefresh
    frm15x15select.Show
    Unload Me
End If

End Sub

Private Sub Form_Load()

IsTrue = 0
Time = 800
lblTime.Caption = Time

Open Stage For Input As #1
Input #1, StageSize, Pic(1), Pic(2), Pic(3), Pic(4), Pic(5), Pic(6), Pic(7), Pic(8), Pic(9), Pic(10), Pic(11), Pic(12), Pic(13), Pic(14), Pic(15), Pic(16), Pic(17), Pic(18), Pic(19), Pic(20), Pic(21), Pic(22), Pic(23), Pic(24), Pic(25), Pic(26), Pic(27), Pic(28), Pic(29), Pic(30), Pic(31), Pic(32), Pic(33), Pic(34), Pic(35), Pic(36), Pic(37), Pic(38), Pic(39), Pic(40), Pic(41), Pic(42), Pic(43), Pic(44), Pic(45), Pic(46), Pic(47), Pic(48), Pic(49), Pic(50)
Input #1, Pic(51), Pic(52), Pic(53), Pic(54), Pic(55), Pic(56), Pic(57), Pic(58), Pic(59), Pic(60), Pic(61), Pic(62), Pic(63), Pic(64), Pic(65), Pic(66), Pic(67), Pic(68), Pic(69), Pic(70), Pic(71), Pic(72), Pic(73), Pic(74), Pic(75), Pic(76), Pic(77), Pic(78), Pic(79), Pic(80), Pic(81), Pic(82), Pic(83), Pic(84), Pic(85), Pic(86), Pic(87), Pic(88), Pic(89), Pic(90), Pic(91), Pic(92), Pic(93), Pic(94), Pic(95), Pic(96), Pic(97), Pic(98), Pic(99), Pic(100)
Input #1, Pic(101), Pic(102), Pic(103), Pic(104), Pic(105), Pic(106), Pic(107), Pic(108), Pic(109), Pic(110), Pic(111), Pic(112), Pic(113), Pic(114), Pic(115), Pic(116), Pic(117), Pic(118), Pic(119), Pic(120), Pic(121), Pic(122), Pic(123), Pic(124), Pic(125), Pic(126), Pic(127), Pic(128), Pic(129), Pic(130), Pic(131), Pic(132), Pic(133), Pic(134), Pic(135), Pic(136), Pic(137), Pic(138), Pic(139), Pic(140), Pic(141), Pic(142), Pic(143), Pic(144), Pic(145), Pic(146), Pic(147), Pic(148), Pic(149), Pic(150)
Input #1, Pic(151), Pic(152), Pic(153), Pic(154), Pic(155), Pic(156), Pic(157), Pic(158), Pic(159), Pic(160), Pic(161), Pic(162), Pic(163), Pic(164), Pic(165), Pic(166), Pic(167), Pic(168), Pic(169), Pic(170), Pic(171), Pic(172), Pic(173), Pic(174), Pic(175), Pic(176), Pic(177), Pic(178), Pic(179), Pic(180), Pic(181), Pic(182), Pic(183), Pic(184), Pic(185), Pic(186), Pic(187), Pic(188), Pic(189), Pic(190), Pic(191), Pic(192), Pic(193), Pic(194), Pic(195), Pic(196), Pic(197), Pic(198), Pic(199), Pic(200)
Input #1, Pic(201), Pic(202), Pic(203), Pic(204), Pic(205), Pic(206), Pic(207), Pic(208), Pic(209), Pic(210), Pic(211), Pic(212), Pic(213), Pic(214), Pic(215), Pic(216), Pic(217), Pic(218), Pic(219), Pic(220), Pic(221), Pic(222), Pic(223), Pic(224), Pic(225)
Input #1, Row(1), Row(2), Row(3), Row(4), Row(5), Row(6), Row(7), Row(8), Row(9), Row(10), Row(11), Row(12), Row(13), Row(14), Row(15), Col(1), Col(2), Col(3), Col(4), Col(5), Col(6), Col(7), Col(8), Col(9), Col(10), Col(11), Col(12), Col(13), Col(14), Col(15), Description, Tmp2
Close #1

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

Call CalcTrue

rc = MsgBox("Would you like a starting hint?", vbYesNo, "Hint?")
If rc = vbYes Then
    Call StartHint
Else
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 1 And KeyCode = vbKeyB Then Call Win
End Sub

Private Sub Win()

Dim PicNum As Integer

PicNum = 1

Do While PicNum < 226
    If Pic(PicNum) = False Then
        Square(PicNum).BackColor = &H955800
        Square(PicNum).Picture = LoadPicture()
        Square(PicNum).BorderStyle = 0
    ElseIf Pic(PicNum) = True Then
        Square(PicNum).Picture = LoadPicture(App.Path & "\true.gif")
    End If
     
    Sleep 25
    PicNum = PicNum + 1
    DoEvents
Loop

If Custom = False And NetGame = False Then
    ExtraHardStagePass(StageNum) = True
    If ExtraHardStageTime(StageNum) < Time Then
        ExtraHardStageTime(StageNum) = Time
    ElseIf ExtraHardStageTime(StageNum) > Time Then
    End If

    Open file For Output As #1
    Write #1, ExtraHardStagePass(1), ExtraHardStageTime(1), ExtraHardStagePass(2), ExtraHardStageTime(2), ExtraHardStagePass(3), ExtraHardStageTime(3), ExtraHardStagePass(4), ExtraHardStageTime(4), ExtraHardStagePass(5), ExtraHardStageTime(5), ExtraHardStagePass(6), ExtraHardStageTime(6), ExtraHardStagePass(7), ExtraHardStageTime(7), ExtraHardStagePass(8), ExtraHardStageTime(8), ExtraHardStagePass(9), ExtraHardStageTime(9), ExtraHardStagePass(10), ExtraHardStageTime(10)
    Close #1
End If

imgExit.Visible = True

End Sub

Public Sub CalcTrue()
Dim PicNum As Integer
NumTrue = 0
PicNum = 1

Do While PicNum < 226
    If Pic(PicNum) = True Then NumTrue = NumTrue + 1
    PicNum = PicNum + 1
Loop

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
    Call ExtraHardRefresh
    frm15x15select.Show
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
        lblExit.Visible = True
    End If
End If

Square(Index).Enabled = False

End Sub

Private Sub Timer1_Timer()

If Time = 0 Then
    Frame1.Enabled = False
    Timer1.Enabled = False
    lblStatus.Caption = "Time's Up! Game Over!"
    lblExit.Visible = True
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

Random = Int((15 * Rnd) + 1)
Random2 = Int((15 * Rnd) + 1)

If Random = 1 Then tmp = 1: Tmp3 = 15
If Random = 2 Then tmp = 16: Tmp3 = 30
If Random = 3 Then tmp = 31: Tmp3 = 45
If Random = 4 Then tmp = 46: Tmp3 = 60
If Random = 5 Then tmp = 61: Tmp3 = 75
If Random = 6 Then tmp = 76: Tmp3 = 90
If Random = 7 Then tmp = 91: Tmp3 = 105
If Random = 8 Then tmp = 106: Tmp3 = 120
If Random = 9 Then tmp = 121: Tmp3 = 135
If Random = 10 Then tmp = 136: Tmp3 = 150
If Random = 11 Then tmp = 151: Tmp3 = 165
If Random = 12 Then tmp = 166: Tmp3 = 180
If Random = 13 Then tmp = 181: Tmp3 = 195
If Random = 14 Then tmp = 196: Tmp3 = 210
If Random = 15 Then tmp = 211: Tmp3 = 225


If Random2 = 1 Then Tmp2 = 1: Tmp4 = 211
If Random2 = 2 Then Tmp2 = 2: Tmp4 = 212
If Random2 = 3 Then Tmp2 = 3: Tmp4 = 213
If Random2 = 4 Then Tmp2 = 4: Tmp4 = 214
If Random2 = 5 Then Tmp2 = 5: Tmp4 = 215
If Random2 = 6 Then Tmp2 = 6: Tmp4 = 216
If Random2 = 7 Then Tmp2 = 7: Tmp4 = 217
If Random2 = 8 Then Tmp2 = 8: Tmp4 = 218
If Random2 = 9 Then Tmp2 = 9: Tmp4 = 219
If Random2 = 10 Then Tmp2 = 10: Tmp4 = 220
If Random2 = 11 Then Tmp2 = 11: Tmp4 = 221
If Random2 = 12 Then Tmp2 = 12: Tmp4 = 222
If Random2 = 13 Then Tmp2 = 13: Tmp4 = 223
If Random2 = 14 Then Tmp2 = 14: Tmp4 = 224
If Random2 = 15 Then Tmp2 = 15: Tmp4 = 225

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
Tmp2 = Tmp2 + 15
Loop

End Sub


