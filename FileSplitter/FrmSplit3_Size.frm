VERSION 5.00
Begin VB.Form FrmSplit3_Size 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File splitter wizard"
   ClientHeight    =   4644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6924
   Icon            =   "FrmSplit3_Size.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6924
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2892
      TabIndex        =   1
      Top             =   4068
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      Height          =   3492
      Left            =   150
      TabIndex        =   3
      Top             =   240
      Width           =   6580
      Begin VB.CheckBox chkFloppy 
         Caption         =   "Fit to Floppy (1.44 MB)"
         Height          =   372
         Left            =   984
         TabIndex        =   7
         Top             =   1680
         Width           =   2424
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         ItemData        =   "FrmSplit3_Size.frx":030A
         Left            =   3264
         List            =   "FrmSplit3_Size.frx":031A
         TabIndex        =   5
         Text            =   " - Choose unit - "
         Top             =   948
         Width           =   1416
      End
      Begin VB.TextBox txtSize 
         Height          =   288
         Left            =   1584
         TabIndex        =   4
         Top             =   948
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Set to"
         Height          =   192
         Left            =   984
         TabIndex        =   6
         Top             =   984
         Width           =   408
      End
   End
End
Attribute VB_Name = "FrmSplit3_Size"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COMBOTEXT As String = " - Choose unit - "

Private Sub chkFloppy_Click()

    If chkFloppy.Value Then
        txtSize.Enabled = False
        txtSize.Text = "1.44"
        Combo1.Text = "MB"
        Combo1.Enabled = False
        Label1.Enabled = False
    Else
        txtSize.Enabled = True
        Combo1.Enabled = True
        Label1.Enabled = True
        Combo1.Text = COMBOTEXT
        txtSize.Text = ""
        txtSize.SetFocus
    End If

End Sub

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
    
    Dim Rate As Long
    Dim Num As Long
    
    If Not IsNumeric(txtSize.Text) Then
       MsgBox "You must enter a number (use 0-9).", vbOKOnly + vbCritical, "Invalid number"
        txtSize.SetFocus
        SendKeys "{HOME}+{END}"
        Exit Sub
    End If
    
    With Combo1
        Select Case .Text
            Case .List(0)
                Rate = KB
            Case .List(1)
                Rate = MB
            Case .List(2)
                Rate = GB
            Case .List(3)
                Rate = 1
            Case Else
                MsgBox "Choose one of the KB, MB, GB or (Bytes) from the list", vbOKOnly + vbExclamation, "Invalid data"
                Combo1.SetFocus
                Exit Sub
        End Select
    End With
    
    Num = Val(txtSize.Text)
    UserSizeOfPieces = Num * Rate 'Size of each piece in bytes.
    
    If chkFloppy.Value Then
        If FileLen(OriginalFileNamePath) > FloppySize Then
            UserSizeOfPieces = FloppySize
        Else
            MsgBox "Your file size is lesser than a floppy capacity !", vbOKOnly + vbInformation, "No need to FileSplitter"
            Exit Sub
        End If
        FrmSplit4.Show
        FrmSplit4.cmdNext.SetFocus
        Unload Me
        Exit Sub
    End If
    
    If UserSizeOfPieces > FileLen(OriginalFileNamePath) Then
        MsgBox "The size you entered is greater than your original file size.", vbOKOnly + vbCritical, "Invalid size"
        txtSize.SetFocus
        SendKeys "{HOME}+{END}"
        Exit Sub
    End If
    
    If UserSizeOfPieces > MAX_NUM_OF_SPLITS Then
        MsgBox "The size you entered is too small.", vbOKOnly + vbCritical, "Invalid size"
        txtSize.SetFocus
        SendKeys "{HOME}+{END}"
        Exit Sub
    End If
    
    FrmSplit4.Show
    FrmSplit4.cmdNext.SetFocus
    Unload Me
    
    Exit Sub
Err:
    MsgBox "Correct the number and choose a valid unit due to your file size.", vbOKOnly + vbCritical, "Invalid data"
    
End Sub

Private Sub Combo1_Change()

    If Combo1.Text = Combo1.List(0) Then Exit Sub
    If Combo1.Text = Combo1.List(1) Then Exit Sub
    If Combo1.Text = Combo1.List(2) Then Exit Sub
    If Combo1.Text = Combo1.List(3) Then Exit Sub
    If Combo1.Text = COMBOTEXT Then Exit Sub
    Combo1.Text = ""
    
End Sub

Private Sub Form_Load()

    FrmSplit3Numor3Size = 32

End Sub
