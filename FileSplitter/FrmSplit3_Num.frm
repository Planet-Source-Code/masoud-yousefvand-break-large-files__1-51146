VERSION 5.00
Begin VB.Form FrmSplit3_Num 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File splitter wizard"
   ClientHeight    =   4644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6924
   Icon            =   "FrmSplit3_Num.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6924
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3780
      Left            =   150
      TabIndex        =   3
      Top             =   144
      Width           =   6580
      Begin VB.TextBox txtNum 
         Height          =   288
         Left            =   2580
         TabIndex        =   4
         Text            =   "2"
         Top             =   1620
         Width           =   1092
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Split the file into "
         Height          =   192
         Left            =   1392
         TabIndex        =   6
         Top             =   1644
         Width           =   1128
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "pieces."
         Height          =   192
         Left            =   3780
         TabIndex        =   5
         Top             =   1644
         Width           =   528
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<     &Back"
      Height          =   372
      Left            =   1080
      TabIndex        =   0
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit Wizard"
      Height          =   372
      Left            =   5160
      TabIndex        =   2
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next     >>"
      Height          =   372
      Left            =   2868
      TabIndex        =   1
      Top             =   4080
      Width           =   1572
   End
End
Attribute VB_Name = "FrmSplit3_Num"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

    FrmSplit2.Show
    FrmSplit2.cmdBack.SetFocus
    Unload Me

End Sub

Private Sub cmdExit_Click()

    ExitWizard

End Sub

Private Sub cmdNext_Click()

    On Error GoTo Err
    
    If Not IsNumeric(txtNum.Text) Then
        MsgBox "You must enter a number (use 0-9).", vbOKOnly + vbCritical, "Invalid number"
        txtNum.SetFocus
        SendKeys "{HOME}+{END}"
        Exit Sub
    End If
    If Val(txtNum.Text) < 2 Then
        MsgBox "You must enter a number greater than 2", vbOKOnly + vbCritical, "Invalid number"
        txtNum.SetFocus
        SendKeys "{HOME}+{END}"
        Exit Sub
    End If
    If Val(txtNum.Text) > FileLen(OriginalFileNamePath) Then
        MsgBox "You must enter a number lesser than your file size in bytes.", vbOKOnly + vbCritical, "Invalid number"
        txtNum.SetFocus
        SendKeys "{HOME}+{END}"
        Exit Sub
    End If
    If Val(txtNum.Text) > MAX_NUM_OF_SPLITS Then
        MsgBox "You must enter a number between 2 and " & _
            CStr(MAX_NUM_OF_SPLITS) & _
                ".", vbOKOnly + vbCritical, "Big number"
        txtNum.SetFocus
        SendKeys "{HOME}+{END}"
        Exit Sub
    End If
    
    UserNumOfPieces = CLng(txtNum.Text)
    FrmSplit4.Show
    FrmSplit4.cmdNext.SetFocus
    Unload Me
    
    Exit Sub
Err:
    MsgBox "You must enter a valid number.", vbOKOnly + vbCritical, "Invalid number"
    txtNum.SetFocus
    SendKeys "{HOME}+{END}"

End Sub

Private Sub Form_Load()

    FrmSplit3Numor3Size = 31

End Sub
