VERSION 5.00
Begin VB.Form FrmAssemble 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File splitter wizard"
   ClientHeight    =   4644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6924
   Icon            =   "FrmAssemble.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6924
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   264
      Left            =   228
      TabIndex        =   11
      Top             =   4140
      Visible         =   0   'False
      Width           =   624
   End
   Begin VB.Frame Frame2 
      Caption         =   " Collection info "
      Height          =   1080
      Left            =   150
      TabIndex        =   7
      Top             =   2880
      Width           =   6580
      Begin VB.Label lblInfo 
         Height          =   660
         Left            =   156
         TabIndex        =   8
         Top             =   324
         Width           =   6216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2496
      Left            =   150
      TabIndex        =   3
      Top             =   180
      Width           =   6580
      Begin VB.ListBox lb 
         Height          =   1008
         Left            =   3600
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1236
         Width           =   2424
      End
      Begin VB.DirListBox Dir1 
         Height          =   1152
         Left            =   500
         TabIndex        =   5
         Top             =   1176
         Width           =   2484
      End
      Begin VB.DriveListBox Drive1 
         Height          =   288
         Left            =   500
         TabIndex        =   4
         Top             =   708
         Width           =   2500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Available collections in this folder :"
         Height          =   192
         Left            =   3600
         TabIndex        =   9
         Top             =   732
         Width           =   2436
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Where are the splitted files ?"
         Height          =   192
         Left            =   504
         TabIndex        =   6
         Top             =   350
         Width           =   2004
      End
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
      Left            =   2904
      TabIndex        =   0
      Top             =   4080
      Width           =   1572
   End
End
Attribute VB_Name = "FrmAssemble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NO_COLLECTION As String = "No collection available"
Private LastGoodDrive As String

Private Sub cmdBack_Click()

    FrmWizard2.Show
    FrmWizard2.cmdBack.SetFocus
    Unload Me
    
End Sub

Private Sub cmdExit_Click()

    ExitWizard

End Sub

Private Sub cmdNext_Click()

    FrmAssemble2.Show
    FrmAssemble2.txtTargetName.SetFocus
    SendKeys "{HOME}+{END}"

    Me.Hide

End Sub

Private Sub Dir1_Change()

    FragmentsPath = Dir1.Path
    Call UpdateListBox
    
End Sub

Private Sub Drive1_Change()

    On Error GoTo DriveErr
    
    Static Last As String
    
    Dir1.Path = Drive1.Drive
    FragmentsPath = Dir1.Path
    Call UpdateListBox
    Exit Sub
    
DriveErr:
    Drive1.Drive = LastGoodDrive
    MsgBox "Drive is unavailable.      ", vbCritical + vbOKOnly, "Error"
    
End Sub

Private Sub Form_Activate()

    LastGoodDrive = Drive1.Drive
    FragmentsPath = Dir1.Path
    cmdNext.Enabled = False
    SelectedCollection = ""
    Call UpdateListBox

End Sub

Private Sub UpdateListBox()

    Dim i As Long
    Dim j As Long
    Dim colname As String
    
    lb.Clear
    
    File1.Refresh
    File1.Path = FragmentsPath
    For i = 0 To File1.ListCount - 1
        If FileExtension(File1.List(i)) = DEFAULT_EXTENSION Then
            colname = GetCollectionName(File1.List(i))
            For j = 0 To lb.ListCount - 1
                If lb.List(j) = colname Then
                    j = 0
                    GoTo next_i
                End If
            Next
            lb.AddItem colname
            If SelectedCollection <> "" Then cmdNext.Enabled = True
        End If
next_i:
    Next
    If lb.ListCount = 0 Then
        lb.AddItem NO_COLLECTION
        cmdNext.Enabled = False
    End If

End Sub

Private Function GetCollectionName(ByVal filename As String) As String

    Dim strTemp As String
    Dim i As Integer

    strTemp = StrReverse(filename)
    i = InStr(1, strTemp, "(", vbTextCompare)
    GetCollectionName = Left(filename, (Len(strTemp) - i))

End Function

Private Sub lb_Click()

    Dim colname As String
    Dim FilePathName As String
    Dim fn As Long
    Dim Header As FragmentHeader
    Dim info As String
    
    colname = lb.List(lb.ListIndex)
    If colname = NO_COLLECTION Then Exit Sub
    cmdNext.Enabled = True
    FilePathName = FragmentsPath & "\" & colname & _
        "(1)" & "." & DEFAULT_EXTENSION
    
    fn = FreeFile
    Open FilePathName For Binary As #fn
    Get #fn, , Header
    Close #fn
    AssembleFileName = Header.OriginalFileName
    lblInfo.Caption = "Original file name is : " & _
        Header.OriginalFileName & vbCrLf & _
        "Original file size is : " & _
        FileKbMbGb(Header.OriginalFileSize) & vbCrLf & _
        "Splitted on : " & Header.DateOfSplitting

End Sub
