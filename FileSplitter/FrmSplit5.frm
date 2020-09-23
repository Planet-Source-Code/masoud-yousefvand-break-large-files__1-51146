VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSplit5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File splitter wizard"
   ClientHeight    =   4644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6924
   Icon            =   "FrmSplit5.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6924
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   696
      Left            =   150
      TabIndex        =   5
      Top             =   2688
      Width           =   6580
      Begin VB.TextBox txtCollectionName 
         Height          =   288
         Left            =   3072
         TabIndex        =   7
         Text            =   "Collection1"
         Top             =   264
         Width           =   3144
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Choose a name for this group of files :"
         Height          =   192
         Left            =   204
         TabIndex        =   6
         Top             =   288
         Width           =   2652
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   336
      Left            =   156
      TabIndex        =   4
      Top             =   3600
      Width           =   6580
      _ExtentX        =   11599
      _ExtentY        =   593
      _Version        =   393216
      Appearance      =   1
      Max             =   10000
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2364
      Left            =   150
      TabIndex        =   3
      Top             =   204
      Width           =   6580
      Begin VB.Label lblInfo 
         Height          =   1920
         Left            =   200
         TabIndex        =   8
         Top             =   276
         Width           =   6168
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<     &Back"
      Height          =   372
      Left            =   1092
      TabIndex        =   0
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit Wizard"
      Height          =   372
      Left            =   5172
      TabIndex        =   2
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "&Split"
      Height          =   372
      Left            =   2880
      TabIndex        =   1
      Top             =   4080
      Width           =   1572
   End
End
Attribute VB_Name = "FrmSplit5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

    FrmSplit4.Show
    FrmSplit4.cmdBack.SetFocus
    Unload Me

End Sub

Private Sub cmdExit_Click()

    On Error GoTo KillErr

    Dim splitcancel As Integer
    Dim i As Long

    If cmdSplit.Enabled Then
        ExitWizard
        End
    Else
        splitcancel = MsgBox("Are you sure you want to cancel splitting ?", _
        vbYesNo + vbExclamation, "Splitting in process...")
        If splitcancel = vbNo Then Exit Sub
        SplitCancealed = True
        Close
        For i = 1 To MAX_NUM_OF_SPLITS
            Kill DestinationSplitFolder & "\" & CollectionName & _
                "(" & CStr(i) & ")" & "." & DEFAULT_EXTENSION
        Next
        pb.Value = pb.Min
        cmdBack.Enabled = True
        cmdSplit.Enabled = True
    End If
    
    Exit Sub
    
KillErr:
    pb.Value = pb.Min
    cmdBack.Enabled = True
    cmdSplit.Enabled = True

End Sub

Private Sub cmdSplit_Click()

    On Error Resume Next

    Dim answer As Integer

    cmdBack.Enabled = False
    cmdSplit.Enabled = False

    If Split(OriginalFileNamePath, DestinationSplitFolder, _
        CollectionName, UserSizeOfPieces, _
        DEFAULT_EXTENSION) Then
            pb.Value = pb.Min
            cmdBack.Enabled = True
            cmdSplit.Enabled = True
            Beep
            answer = MsgBox("Splitting completed successfully.Go to destination folder ?", _
                vbYesNo + vbQuestion, "")
            If answer = vbYes Then
                Shell "explorer.exe " & DestinationSplitFolder, vbNormalFocus
            Else
                Me.Hide
                MsgBox "Your splitted files are stored at '" & _
                    DestinationSplitFolder & "'", _
                    vbInformation + vbOKOnly, "FileSplitter"
            End If
            ExitWizard
            End
    Else
        MsgBox "Unknown error occured." & vbCrLf & _
            "FileSplitter cannot fragment your file.", _
            vbOKOnly + vbCritical, "Fatal error"
    End If
    
End Sub

Private Sub Form_Load()

    Dim number As Long
    Dim fs As Long ' Original file size
    Dim sizeofall As Long
    Dim sizeoflast As Long
    Dim sizeinfo As String

    CollectionName = txtCollectionName.Text
    cmdBack.Enabled = True
    cmdSplit.Enabled = True

    fs = FileLen(OriginalFileNamePath)

    If FrmSplit3Numor3Size = 31 Then ' User has determined number of files
        number = UserNumOfPieces
        sizeofall = Fix(fs / number)
        If fs = sizeofall * number Then
            sizeinfo = "Each file size will be : " & FileKbMbGb(sizeofall + HEADER_SIZE) & vbCrLf
            UserSizeOfPieces = sizeofall
        Else
            sizeofall = Fix(fs / number)
            sizeoflast = fs - ((number - 1) * sizeofall)
            sizeinfo = CStr(number - 1) & " file(s) size will be : " & FileKbMbGb(sizeofall + HEADER_SIZE) & vbCrLf & _
                "and one file size will be : " & FileKbMbGb(sizeoflast + HEADER_SIZE) & vbCrLf
            UserSizeOfPieces = sizeofall + 1
        End If
    ElseIf FrmSplit3Numor3Size = 32 Then ' User has determined size
        sizeofall = UserSizeOfPieces - HEADER_SIZE
        number = Fix(fs / sizeofall)
        sizeoflast = fs - (number * sizeofall)
        If sizeoflast = 0 Then
            sizeinfo = "Each file size will be : " & FileKbMbGb(UserSizeOfPieces) & vbCrLf
        Else
            sizeinfo = CStr(number) & " file(s) size will be : " & FileKbMbGb(UserSizeOfPieces) & vbCrLf & _
                "and one file size will be : " & FileKbMbGb(sizeoflast + HEADER_SIZE) & vbCrLf
        End If
        UserSizeOfPieces = sizeofall
    End If

    sizeinfo = sizeinfo & vbCrLf & "NOTE : " & CStr(HEADER_SIZE) & " bytes will be use for necessary information in each file."

    lblInfo.Caption = "Original file name is : " & GetFileName(OriginalFileNamePath) & vbCrLf & _
        "Original file size is : " & FileKbMbGb(FileLen(OriginalFileNamePath)) & vbCrLf & _
        "Original file location is : " & GetFilePath(OriginalFileNamePath) & "\" & vbCrLf & _
        vbCrLf & _
        "This file will be splitted to " & CStr(number) & " pieces." & vbCrLf & sizeinfo

End Sub

Private Sub txtCollectionName_Change()

    CollectionName = txtCollectionName.Text

End Sub
