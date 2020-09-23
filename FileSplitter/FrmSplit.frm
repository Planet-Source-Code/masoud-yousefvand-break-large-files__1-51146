VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSplit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File splitter wizard"
   ClientHeight    =   4644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6924
   Icon            =   "FrmSplit.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6924
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3804
      Left            =   156
      TabIndex        =   3
      Top             =   132
      Width           =   6580
      Begin VB.TextBox txtOriginalFile 
         Enabled         =   0   'False
         Height          =   288
         Left            =   180
         TabIndex        =   5
         Top             =   1776
         Width           =   4932
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   288
         Left            =   5100
         TabIndex        =   4
         Top             =   1776
         Width           =   1212
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Where is the file ?"
         Height          =   192
         Left            =   250
         TabIndex        =   6
         Top             =   1296
         Width           =   1248
      End
   End
   Begin MSComDlg.CommonDialog db 
      Left            =   6432
      Top             =   96
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      Filter          =   "All files"
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<     &Back"
      Height          =   372
      Left            =   1080
      TabIndex        =   2
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit Wizard"
      Height          =   372
      Left            =   5160
      TabIndex        =   1
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next     >>"
      Height          =   372
      Left            =   2880
      TabIndex        =   0
      Top             =   4080
      Width           =   1572
   End
End
Attribute VB_Name = "FrmSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

    FrmWizard2.Show
    FrmWizard2.cmdBack.SetFocus
    Unload Me
    
End Sub

Private Sub cmdBrowse_Click()

    db.ShowOpen
    txtOriginalFile.Text = db.filename
    OriginalFileNamePath = db.filename
    cmdNext.SetFocus

End Sub

Private Sub cmdExit_Click()

    ExitWizard

End Sub

Private Sub cmdNext_Click()

    If txtOriginalFile.Text = "" Then
        cmdBrowse.SetFocus
        MsgBox "You must locate a file for splitting.     ", vbCritical, "Invalid file name"
        Exit Sub
    End If
    OriginalFileNamePath = txtOriginalFile.Text
    FrmSplit2.Show
    FrmSplit2.cmdNext.SetFocus
    Unload Me
    
End Sub
