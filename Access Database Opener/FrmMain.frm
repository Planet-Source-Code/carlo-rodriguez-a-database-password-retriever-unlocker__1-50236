VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UnLock Access Database Password"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Database FileName:"
      Height          =   885
      Left            =   60
      TabIndex        =   23
      Top             =   60
      Width           =   5925
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Browse"
         Height          =   525
         Left            =   120
         Picture         =   "FrmMain.frx":0442
         TabIndex        =   24
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblFileName 
         Caption         =   "Database FileName...."
         Height          =   405
         Left            =   1050
         TabIndex        =   25
         Top             =   330
         Width           =   4725
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database Table List:"
      Height          =   2235
      Left            =   3900
      TabIndex        =   21
      Top             =   990
      Width           =   2085
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         ItemData        =   "FrmMain.frx":0D0C
         Left            =   90
         List            =   "FrmMain.frx":0D0E
         TabIndex        =   22
         Top             =   330
         Width           =   1875
      End
   End
   Begin MSComDlg.CommonDialog cdiag 
      Left            =   3960
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Access Database | *.mdb"
   End
   Begin VB.Frame Frame4 
      Caption         =   "Progress Statistics:"
      Height          =   2235
      Left            =   60
      TabIndex        =   10
      Top             =   990
      Width           =   3825
      Begin VB.Label lblCurrPass 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1530
         TabIndex        =   20
         Top             =   1740
         Width           =   2145
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1740
         Width           =   1545
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1530
         TabIndex        =   18
         Top             =   1380
         Width           =   2145
      End
      Begin VB.Label lblLength 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1530
         TabIndex        =   17
         Top             =   1050
         Width           =   2145
      End
      Begin VB.Label lblTotalCombo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1530
         TabIndex        =   16
         Top             =   690
         Width           =   2145
      End
      Begin VB.Label lblCombSec 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1530
         TabIndex        =   15
         Top             =   360
         Width           =   2145
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Running Time:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Current String Length:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1050
         Width           =   1875
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Combinations:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   690
         Width           =   1485
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Combinations / Sec:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1425
      End
   End
   Begin VB.Timer tRuntime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3990
      Top             =   3510
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4530
      TabIndex        =   4
      Top             =   4020
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4530
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4530
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1935
      Left            =   60
      TabIndex        =   6
      Top             =   3270
      Width           =   3825
      Begin VB.TextBox txtComboLen 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1500
         Width           =   1095
      End
      Begin VB.TextBox txtStartCombo 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1500
         Width           =   1095
      End
      Begin VB.TextBox txtCharacterSet 
         Height          =   735
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "FrmMain.frx":0D10
         Top             =   510
         Width           =   2565
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Start Length:"
         Height          =   210
         Left            =   1440
         TabIndex        =   9
         Top             =   1260
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Starting String:"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Character Set:"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   1050
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WithEvents cBF As clsBF
Attribute cBF.VB_VarHelpID = -1
Dim lRunningTime As Long
Dim bCrack As Boolean
Dim db As Database
Dim dbFileName As String

Private Sub cBF_CombinationsPerSec(Combos As Long)
    lblCombSec.Caption = Format(Combos, "#,###")
    lblCurrPass.Caption = cBF.CurrentPassword
    lblLength.Caption = Len(cBF.CurrentPassword)
End Sub

Private Sub cBF_TotalCombinations(Combos As String)
    lblTotalCombo.Caption = Format(Combos, "#,###")
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdOpen_Click()
On Error GoTo errH
    cdiag.ShowOpen
    lblFileName.Caption = cdiag.FileName
    If lblFileName.Caption <> "" Then cmdStart.Enabled = True
    dbFileName = cdiag.FileName
    Exit Sub
errH:
lblFileName.Caption = ""
End Sub

Private Sub cmdStart_Click()

On Error GoTo errhandler
Dim sTmp As String
    List1.Clear
    lRunningTime = 0
    bCrack = True
    DisEnableControls

With cBF
    .CharacterSet = txtCharacterSet.Text
    .FirstPassword = txtStartCombo.Text
    If txtComboLen.Text <> "" Then .StartLength = CInt(txtComboLen.Text)
    .Initialize

        Do Until bCrack = False Or List1.ListCount > 0
            DoEvents
            Set db = Nothing
            Set db = OpenDatabase(dbFileName, False, False, ";pwd=" & sTmp)
            GetData
            If List1.ListCount = 0 Then
                sTmp = .BruteForce
            End If
        Loop

    If bCrack = True Then
        lblCurrPass.Caption = .CurrentPassword
        MsgBox "Cracked in " & lblTime.Caption & vbCr & "Password = " & .CurrentPassword, vbApplicationModal + vbInformation, Me.Caption
    End If

    bCrack = False
    DisEnableControls

End With

Exit Sub
errhandler:

If Err.Number = 3031 Or Err.Number = 91 Then
    Resume Next
Else
    MsgBox Err.Source
End If

End Sub
Sub DisEnableControls()

    tRuntime.Enabled = Not tRuntime.Enabled
    txtCharacterSet.Enabled = Not txtCharacterSet.Enabled
    txtStartCombo.Enabled = txtStartCombo.Enabled
    txtComboLen.Enabled = Not txtComboLen.Enabled
    cmdStart.Enabled = Not cmdStart.Enabled
    cmdStop.Enabled = Not cmdStop.Enabled
    cmdOpen.Enabled = Not cmdOpen.Enabled
    
End Sub

Function TimeConv(Sec As Long) As String
Dim iSeconds As Integer
Dim iMinurts As Integer
Dim iHours As Integer
Dim iDays As Integer
iSeconds = Sec Mod 60
iMinurts = Int(Sec / 60)
iHours = Int(iMinurts / 60)
iDays = Int(iHours / 24)
TimeConv = iDays & " Days " & iHours & ":" & iMinurts & ":" & iSeconds
End Function
Private Sub cmdStop_Click()
    bCrack = False
End Sub

Private Sub Command1_Click()
    cdiag.ShowOpen
End Sub

Private Sub Form_Load()
    Set cBF = New clsBF
End Sub

Private Sub tRuntime_Timer()
    lRunningTime = lRunningTime + 1
    lblTime.Caption = TimeConv(lRunningTime)
End Sub

Private Sub GetData()
List1.Clear
    For i = 0 To db.TableDefs.Count - 1
            List1.AddItem (db.TableDefs(i).Name)
    Next i
End Sub

Private Sub txtText_Change()

End Sub
