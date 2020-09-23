VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAssemble2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File splitter wizard"
   ClientHeight    =   4644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6924
   Icon            =   "FrmAssemble2.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6924
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   648
      Left            =   150
      TabIndex        =   5
      Top             =   2844
      Width           =   6580
      Begin VB.TextBox txtTargetName 
         Height          =   288
         Left            =   132
         TabIndex        =   10
         Top             =   240
         Width           =   1812
      End
      Begin VB.Label lblFilePath 
         Height          =   288
         Left            =   2040
         TabIndex        =   6
         Top             =   252
         Width           =   4296
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2568
      Left            =   150
      TabIndex        =   3
      Top             =   204
      Width           =   6580
      Begin VB.DriveListBox Drive1 
         Height          =   288
         Left            =   350
         TabIndex        =   8
         Top             =   780
         Width           =   2500
      End
      Begin VB.DirListBox Dir1 
         Height          =   1152
         Left            =   350
         TabIndex        =   7
         Top             =   1224
         Width           =   2484
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Where do you want to store the original file ?"
         Height          =   192
         Left            =   350
         TabIndex        =   4
         Top             =   324
         Width           =   3108
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
   Begin VB.CommandButton cmdAssemble 
      Caption         =   "&Assemble"
      Height          =   372
      Left            =   2880
      TabIndex        =   1
      Top             =   4080
      Width           =   1572
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   336
      Left            =   150
      TabIndex        =   9
      Top             =   3624
      Width           =   6580
      _ExtentX        =   11599
      _ExtentY        =   593
      _Version        =   393216
      Appearance      =   1
      Max             =   10000
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmAssemble2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastGoodDrive As String
Private AssembleFolder As String

Private Sub cmdAssemble_Click()

    On Error GoTo Err
    
    Dim SplitsArray() As String
    Dim fn As Long
    Dim FilePathName As String
    Dim TargetPathName As String
    Dim Header As FragmentHeader
    Dim NumberOfPieces As Long
    Dim i As Long
    Dim j As Long
    Dim answer As Integer
    Dim FileExist As Boolean
    
    cmdBack.Enabled = False
    cmdAssemble.Enabled = False
    
    FrmAssemble.File1.Refresh
    FileExist = False
    SelectedCollection = FrmAssemble.lb.List(FrmAssemble.lb.ListIndex)
    FilePathName = FragmentsPath & "\" & SelectedCollection & _
        "(1)" & "." & DEFAULT_EXTENSION
    
    fn = FreeFile
    Open FilePathName For Binary As #fn
    Get #fn, , Header
    Close #fn
    
    NumberOfPieces = Header.NumberOfFragments
    For i = 1 To NumberOfPieces
        FileExist = False
        For j = 0 To FrmAssemble.File1.ListCount - 1
            If FrmAssemble.File1.List(j) = SelectedCollection & "(" & _
                CStr(i) & ")" & "." & _
                DEFAULT_EXTENSION Then FileExist = True
        Next
        If Not FileExist Then GoTo MissingFile
    Next
    
    ReDim SplitsArray(1 To NumberOfPieces)
    
    For i = 1 To NumberOfPieces
        SplitsArray(i) = FragmentsPath & "\" & SelectedCollection & _
        "(" & CStr(i) & ")" & "." & DEFAULT_EXTENSION
    Next
    
    TargetPathName = AssembleFolder & txtTargetName.Text
    If Assembler(TargetPathName, SplitsArray) Then
        pb.Value = pb.Min
        Beep
        answer = MsgBox("Assembling completed successfully.Go to destination folder ?", _
            vbYesNo + vbQuestion, "")
        If answer = vbYes Then
            Shell "explorer.exe " & AssembleFolder, vbNormalFocus
        Else
            Me.Hide
            MsgBox "Your assembled file is stored at '" & _
                AssembleFolder & "'", _
                vbInformation + vbOKOnly, "FileSplitter"
        End If
        ExitWizard
        End
    Else
        MsgBox "Unknown error occured." & vbCrLf & _
            "FileSplitter cannot assemble your file.", _
            vbOKOnly + vbCritical, "Fatal error"
        Close
        Kill AssembleFolder & txtTargetName.Text
    End If
    
    cmdBack.Enabled = True
    cmdAssemble.Enabled = True
    pb.Value = pb.Min
    
    Exit Sub
    
MissingFile:
    MsgBox "Some file(s) is/are missing." & vbCrLf & _
        "FileSplitter cannot make " & Header.OriginalFileName, _
        vbCritical + vbOKOnly, "Missing file"
        
    cmdBack.Enabled = True
    cmdAssemble.Enabled = True
    Exit Sub
Err:
   MsgBox "Unknown error occured." & vbCrLf & _
        "FileSplitter cannot assemble your file.", _
        vbOKOnly + vbCritical, "Fatal error"
        
    cmdBack.Enabled = True
    cmdAssemble.Enabled = True

End Sub


Private Sub cmdBack_Click()

    FrmAssemble.Show
    FrmAssemble.cmdBack.SetFocus
    Unload Me

End Sub

Private Sub cmdExit_Click()

    On Error GoTo KillErr

    Dim assemblecancel As Integer
    Dim i As Long

    If cmdAssemble.Enabled Then
        ExitWizard
        End
    Else
        assemblecancel = MsgBox("Are you sure you want to cancel assembling ?", _
        vbYesNo + vbExclamation, "assembling in process...")
        If assemblecancel = vbNo Then Exit Sub
        AssembleCancealed = True
        Close
        Kill AssembleFolder & txtTargetName.Text
        pb.Value = pb.Min
        cmdBack.Enabled = True
        cmdAssemble.Enabled = True
    End If
    
    Exit Sub
    
KillErr:
    pb.Value = pb.Min
    cmdBack.Enabled = True
    cmdAssemble.Enabled = True

End Sub

Private Sub Dir1_Change()

    AssembleFolder = Dir1.Path
    AssembleFolder = Trim(AssembleFolder)
    If Not Right(AssembleFolder, 1) = "\" Then
        AssembleFolder = AssembleFolder & "\"
    End If
    lblFilePath.Caption = " will be saved at " & Dir1.Path
    
End Sub

Private Sub Drive1_Change()

    On Error GoTo DriveErr
    
    Static Last As String
    
    Dir1.Path = Drive1.Drive
    AssembleFolder = Dir1.Path
    AssembleFolder = Trim(AssembleFolder)
    If Not Right(AssembleFolder, 1) = "\" Then
        AssembleFolder = AssembleFolder & "\"
    End If
    lblFilePath.Caption = " will be saved at " & Dir1.Path
    Exit Sub
    
DriveErr:
    Drive1.Drive = LastGoodDrive
    MsgBox "Drive is unavailable.      ", vbCritical + vbOKOnly, "Error"
    
End Sub

Private Sub Form_Load()

    cmdBack.Enabled = True
    cmdAssemble.Enabled = True
    pb.Value = pb.Min
    lblFilePath.Caption = " will be saved at " & Dir1.Path
        
    txtTargetName.Text = Trim(AssembleFileName)
    
End Sub
