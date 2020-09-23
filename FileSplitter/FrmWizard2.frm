VERSION 5.00
Begin VB.Form FrmWizard2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File splitter wizard"
   ClientHeight    =   4644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6924
   Icon            =   "FrmWizard2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6924
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   384
      Left            =   1950
      Picture         =   "FrmWizard2.frx":030A
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   7
      Top             =   2268
      Width           =   384
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   384
      Left            =   1950
      Picture         =   "FrmWizard2.frx":0614
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   6
      Top             =   1644
      Width           =   384
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<     &Back"
      Height          =   372
      Left            =   1080
      TabIndex        =   5
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit Wizard"
      Height          =   372
      Left            =   5160
      TabIndex        =   4
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next     >>"
      Height          =   372
      Left            =   2880
      TabIndex        =   3
      Top             =   4080
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      Caption         =   "  What do you want to do ?  "
      Height          =   2256
      Left            =   1176
      TabIndex        =   0
      Top             =   996
      Width           =   4356
      Begin VB.OptionButton OptionAssemble 
         Caption         =   " &Assemble the splitted pieces."
         Height          =   252
         Left            =   1350
         TabIndex        =   2
         Top             =   1356
         Width           =   2532
      End
      Begin VB.OptionButton OptionSplit 
         Caption         =   " &Split a file to some pieces."
         Height          =   372
         Left            =   1350
         TabIndex        =   1
         Top             =   720
         Value           =   -1  'True
         Width           =   2172
      End
   End
End
Attribute VB_Name = "FrmWizard2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBack_Click()

    FrmWizard1.Show
    Unload Me
    
End Sub

Private Sub cmdExit_Click()

    ExitWizard

End Sub

Private Sub cmdNext_Click()
 
    If OptionAssemble.Value Then
        FrmAssemble.Show
'        FrmAssemble.cmdAssemble.SetFocus
    End If
    If OptionSplit.Value Then
        FrmSplit.Show
        FrmSplit.cmdBrowse.SetFocus
    End If
    Unload Me
    
End Sub
