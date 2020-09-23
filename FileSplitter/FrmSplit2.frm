VERSION 5.00
Begin VB.Form FrmSplit2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File splitter wizard"
   ClientHeight    =   4644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6924
   Icon            =   "FrmSplit2.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6924
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " How do you want to split this file ?  "
      Height          =   1572
      Left            =   150
      TabIndex        =   7
      Top             =   2160
      Width           =   6580
      Begin VB.OptionButton Option2 
         Caption         =   "Define the size of each particle."
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   6012
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Define the number of particles."
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   6012
      End
   End
   Begin VB.Frame FileInfo 
      Caption         =   " File Information "
      Height          =   1440
      Left            =   150
      TabIndex        =   3
      Top             =   240
      Width           =   6580
      Begin VB.Label Label3 
         Height          =   252
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   6012
      End
      Begin VB.Label Label2 
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   6012
      End
      Begin VB.Label Label1 
         Height          =   252
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   6012
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
      Left            =   2880
      TabIndex        =   1
      Top             =   4080
      Width           =   1572
   End
End
Attribute VB_Name = "FrmSplit2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()

    FrmSplit.Show
    FrmSplit.cmdBack.SetFocus
    Unload Me
    
End Sub

Private Sub cmdExit_Click()

    ExitWizard

End Sub

Private Sub cmdNext_Click()

    If Option1.Value Then
        FrmSplit3_Num.Show
        FrmSplit3_Num.txtNum.SetFocus
        SendKeys "{HOME}+{END}"
    End If
    If Option2.Value Then
        FrmSplit3_Size.Show
        FrmSplit3_Size.txtSize.SetFocus
        SendKeys "{HOME}+{END}"
    End If
    Unload Me

End Sub

Private Sub Form_Activate()

    Label1.Caption = "File Name : " & GetFileName(OriginalFileNamePath)
    Label2.Caption = "File Size : " & FileKbMbGb(FileLen(OriginalFileNamePath))
    Label3.Caption = "File Location : " & GetFilePath(OriginalFileNamePath) & "\"

End Sub

'Private Sub Form_Load()
'
'    Label1.Caption = "File Name : " & GetFileName(OriginalFileNamePath)
'    Label2.Caption = "File Size : " & FileKbMbGb(FileLen(OriginalFileNamePath))
'    Label3.Caption = "File Location : " & GetFilePath(OriginalFileNamePath) & "\"
'
'End Sub
