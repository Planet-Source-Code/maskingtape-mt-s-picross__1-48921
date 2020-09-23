VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmNetLevels 
   BackColor       =   &H00955800&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download Levels"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4965
   Icon            =   "frmNetLevels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   330
      Left            =   1350
      TabIndex        =   11
      Top             =   2205
      Width           =   960
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1905
      Left            =   90
      TabIndex        =   10
      Top             =   270
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   3360
      _Version        =   393217
      Indentation     =   531
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2430
      Top             =   810
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2430
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2430
      Top             =   405
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4365
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download List"
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   2205
      Width           =   1170
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
      Left            =   2475
      TabIndex        =   9
      Top             =   2340
      Width           =   2445
   End
   Begin VB.Image Image12 
      Height          =   420
      Left            =   2430
      Picture         =   "frmNetLevels.frx":1272
      Stretch         =   -1  'True
      Top             =   2250
      Width           =   2835
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   2415
      X2              =   2415
      Y1              =   0
      Y2              =   2625
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   2415
      X2              =   0
      Y1              =   2625
      Y2              =   2625
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Stage List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   840
      TabIndex        =   8
      Top             =   0
      Width           =   750
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   2835
      X2              =   4935
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2835
      X2              =   2835
      Y1              =   0
      Y2              =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Size:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2940
      TabIndex        =   7
      Top             =   315
      Width           =   750
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2940
      TabIndex        =   6
      Top             =   735
      Width           =   960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Created By:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2940
      TabIndex        =   5
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
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
      Left            =   3885
      TabIndex        =   4
      Top             =   315
      Width           =   960
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
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
      Left            =   3885
      TabIndex        =   3
      Top             =   735
      Width           =   1065
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
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
      Left            =   2835
      TabIndex        =   2
      Top             =   1365
      Width           =   2115
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Stage Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   1380
   End
   Begin VB.Image Image20 
      Height          =   1725
      Left            =   2880
      Picture         =   "frmNetLevels.frx":19BC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   2625
      Left            =   0
      Picture         =   "frmNetLevels.frx":2328
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2400
   End
End
Attribute VB_Name = "frmNetLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FileOK = True
Timer1.Enabled = True
URL = "http://bellsouthpwp.net/m/a/maskingtape/netpicross/5x5.txt"
Call DownloadFile
Call CheckFile(App.Path & "\5x5.txt")
Command1.Enabled = False
If FileOK = False Then
    Timer1.Enabled = False
    lblStatus.Caption = "File Error. Please try again."
    Command1.Enabled = True
    Exit Sub
ElseIf FileOK = True Then
    Open App.Path & "\5x5.txt" For Input As #1
    While Not EOF(1)
    Input #1, MyString$
    TreeView1.Nodes.Add "5x5", tvwChild, "I" & TreeView1.Nodes.Count + 2, MyString$
    DoEvents
    Wend
    Close #1
    Kill App.Path & "\5x5.txt"
    lblStatus.Caption = "Ready...."
    Timer1.Enabled = False
End If

FileOK = True
Timer1.Enabled = True
URL = "http://bellsouthpwp.net/m/a/maskingtape/netpicross/10x10.txt"
Call DownloadFile
Call CheckFile(App.Path & "\10x10.txt")
Command1.Enabled = False
If FileOK = False Then
    Timer1.Enabled = False
    lblStatus.Caption = "File Error. Please try again."
    Command1.Enabled = True
    Exit Sub
ElseIf FileOK = True Then
    Open App.Path & "\10x10.txt" For Input As #1
    While Not EOF(1)
    Input #1, MyString$
    TreeView1.Nodes.Add "10x10", tvwChild, "I" & TreeView1.Nodes.Count + 2, MyString$
    DoEvents
    Wend
    Close #1
    Kill App.Path & "\10x10.txt"
    lblStatus.Caption = "Ready...."
    Timer1.Enabled = False
End If

FileOK = True
Timer1.Enabled = True
URL = "http://bellsouthpwp.net/m/a/maskingtape/netpicross/15x15.txt"
Call DownloadFile
Call CheckFile(App.Path & "\15x15.txt")
Command1.Enabled = False
If FileOK = False Then
    Timer1.Enabled = False
    lblStatus.Caption = "File Error. Please try again."
    Command1.Enabled = True
    Exit Sub
ElseIf FileOK = True Then
    Open App.Path & "\15x15.txt" For Input As #1
    While Not EOF(1)
    Input #1, MyString$
    TreeView1.Nodes.Add "15x15", tvwChild, "I" & TreeView1.Nodes.Count + 2, MyString$
    DoEvents
    Wend
    Close #1
    Kill App.Path & "\15x15.txt"
    lblStatus.Caption = "Ready...."
    Timer1.Enabled = False
End If


End Sub

Private Sub Command2_Click()
FileOK = True
Timer3.Enabled = True
URL = "http://bellsouthpwp.net/m/a/maskingtape/netpicross/" & NetFile & ".pcs"
Call DownloadFile
Call CheckFile(App.Path & "\" & NetFile & ".pcs")
If FileOK = False Then
    lblStatus.Caption = "File Error. Please try again."
    Timer3.Enabled = False
    Exit Sub
End If

End Sub

Private Sub Form_Load()
TreeView1.Nodes.Add , , "5x5", "5 x 5 Stages"
TreeView1.Nodes.Add , , "10x10", "10 x 10 Stages"
TreeView1.Nodes.Add , , "15x15", "15 x 15 Stages"

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub List1_Click()
FileOK = True
If List1.Text = "" Or List1.Text = "Coming Soon" Then
    Command2.Enabled = False
Else
    Timer2.Enabled = True
    URL = "http://bellsouthpwp.net/m/a/maskingtape/netpicross/" & List1.Text & ".txt"
    Call DownloadFile
    Call CheckFile(App.Path & "\" & List1.Text & ".txt")
    If FileOK = False Then
        lblStatus.Caption = "File Error. Please try again."
        Label4.Caption = "ERROR"
        Label5.Caption = "ERROR"
        Label6.Caption = "ERROR"
        Timer2.Enabled = False
        Exit Sub
    End If
    Command2.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
Dim MyString As String

If Inet1.StillExecuting = True Then
    lblStatus.Caption = "Downloading List, Please Wait."
End If
End Sub

Private Sub DownloadFile()
On Error GoTo errorhandler
Dim myData() As Byte
If Inet1.StillExecuting = True Then Exit Sub
myData() = Inet1.OpenURL(URL, icByteArray)
For X = Len(URL) To 1 Step -1
If Left$(Right$(URL, X), 1) = "/" Then RealFile$ = Right$(URL, X - 1)
Next X
myFile$ = App.Path + "\" + RealFile$
Open myFile$ For Binary Access Write As #1
Put #1, , myData()
Close #1

Exit Sub

errorhandler:
    Timer1.Enabled = False
    lblStatus.Caption = "File error! Try again later."

End Sub

Private Sub Timer2_Timer()
Dim Grid As String
Dim Description As String
Dim Created As String

If Inet1.StillExecuting = True Then
    lblStatus.Caption = "Downloading Info, Please Wait."
Else
    Open App.Path & "\" & NetFile & ".txt" For Input As #1
    Input #1, Grid, Description, Created
    Close #1
    Kill App.Path & "\" & NetFile & ".txt"
    Label4.Caption = Grid
    Label5.Caption = Description
    Label6.Caption = Created
    lblStatus.Caption = "Ready...."
    Timer2.Enabled = False
End If

End Sub

Private Sub Timer3_Timer()
If Inet1.StillExecuting = True Then
    lblStatus.Caption = "Loading Stage, Please Wait."
Else
    Stage = App.Path & "\" & NetFile & ".pcs"
    Open App.Path & "\" & NetFile & ".pcs" For Input As #1
    Input #1, StageSize
    Close #1

    If StageSize = "5x5" Then
        frm5x5.Show
        NetGame = True
        Me.Hide
    ElseIf StageSize = "10x10" Then
        frm10x10.Show
        NetGame = True
        Me.Hide
    ElseIf StageSize = "15x15" Then
        frm15x15.Show
        NetGame = True
        Me.Hide
    End If
    Timer3.Enabled = False
End If

End Sub

Public Sub CheckFile(file As String)
On Error GoTo errorhandler

Dim tmp As String

Open file For Input As #1
Input #1, tmp
Close #1
Tmp2 = Len(tmp)

If InStr(1, tmp, "<html>") Then FileOK = False Else FileOK = True
Exit Sub

errorhandler:
    FileOK = False
    Close #1
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Text = "5 x 5 Stages" Or Node.Text = "10 x 10 Stages" Or Node.Text = "15 x 15 Stages" Then
    Command2.Enabled = False
Else
    NetFile = Node.Text
    FileOK = True
    Command2.Enabled = False
    Timer2.Enabled = True
    URL = "http://bellsouthpwp.net/m/a/maskingtape/netpicross/" & NetFile & ".txt"
    Call DownloadFile
    Call CheckFile(App.Path & "\" & NetFile & ".txt")
    If FileOK = False Then
        lblStatus.Caption = "File Error. Please try again."
        Label4.Caption = "ERROR"
        Label5.Caption = "ERROR"
        Label6.Caption = "ERROR"
        Timer2.Enabled = False
        Exit Sub
    End If
    Command2.Enabled = True
End If

End Sub
