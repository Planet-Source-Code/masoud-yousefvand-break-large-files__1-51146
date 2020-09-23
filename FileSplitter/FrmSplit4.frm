VERSION 5.00
Begin VB.Form FrmSplit4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File splitter wizard"
   ClientHeight    =   4644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6924
   Icon            =   "FrmSplit4.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6924
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   624
      Left            =   150
      TabIndex        =   7
      Top             =   3264
      Width           =   6580
      Begin VB.Label lblDestinationFolder 
         Height          =   228
         Left            =   180
         TabIndex        =   8
         Top             =   252
         Width           =   6216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3024
      Left            =   150
      TabIndex        =   3
      Top             =   180
      Width           =   6580
      Begin VB.DriveListBox Drive1 
         Height          =   288
         Left            =   500
         TabIndex        =   5
         Top             =   708
         Width           =   2500
      End
      Begin VB.DirListBox Dir1 
         Height          =   1584
         Left            =   500
         TabIndex        =   6
         Top             =   1176
         Width           =   2496
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Where to store the splitted pieces ?"
         Height          =   192
         Left            =   500
         TabIndex        =   4
         Top             =   348
         Width           =   2496
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
Attribute VB_Name = "FrmSplit4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastGoodDrive As String

Private Sub cmdBack_Click()

    If FrmSplit3Numor3Size = 31 Then
        FrmSplit3_Num.Show
        FrmSplit3_Num.cmdBack.SetFocus
        Unload Me
    ElseIf FrmSplit3Numor3Size = 32 Then
        FrmSplit3_Size.Show
        FrmSplit3_Size.cmdBack.SetFocus
        Unload Me
    End If

End Sub

Private Sub cmdExit_Click()

    ExitWizard

End Sub

Private Sub cmdNext_Click()

    DestinationSplitFolder = Dir1.Path
    FrmSplit5.Show
    FrmSplit5.txtCollectionName.SetFocus
    SendKeys "{HOME}+{END}"
    Unload Me

End Sub

Private Sub Dir1_Change()

    lblDestinationFolder.Caption = "Splitted files will be store at  " & _
        Dir1.Path
    DestinationSplitFolder = Dir1.Path
    
End Sub

Private Sub Drive1_Change()

    On Error GoTo DriveErr
    
    Static Last As String
    
    Dir1.Path = Drive1.Drive
    lblDestinationFolder.Caption = "Splitted files will be store at  " & _
        Dir1.Path
    DestinationSplitFolder = Dir1.Path
    Exit Sub
    
DriveErr:
    Drive1.Drive = LastGoodDrive
    MsgBox "Drive is unavailable.      ", vbCritical + vbOKOnly, "Error"
    
End Sub

Private Sub Form_Load()

    LastGoodDrive = Drive1.Drive
    lblDestinationFolder.Caption = "Splitted files will be store at  " & Dir1.Path

End Sub
