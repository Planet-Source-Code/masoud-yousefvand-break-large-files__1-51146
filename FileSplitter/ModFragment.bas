Attribute VB_Name = "ModFragment"
Option Explicit

Private Const AUTHOR_COMMENTS As String = "*** By Dr. M. " & _
    "Yousefvand E-mail address : yousefvand@yahoo.com ***" '58 bytes

Public Type FragmentHeader 'Header size = 256 bytes
    UniqueIdentifier As Long ' 4 bytes
    FragmentNumber As Long ' 4 bytes
    FragmentSize As Long ' 4 bytes
    NumberOfFragments As Long ' 4 bytes
    OriginalFileSize As Long ' 4 bytes
    OriginalFileName As String * 100
    DateOfSplitting As String * 36
    AuthorComment As String * 100
End Type

Public Function Split(ByVal OriginalFilePathName As String, ByVal _
    SaveToFolder As String, ByVal CollectionName As String, ByVal _
    FragmentSize As Long, Optional FragmentExtension _
    As String = "zzz") As Boolean
    
    Dim i As Long
    Dim FileSize As Long
    Dim SplitNumber As Long
    Dim FragmentName As String
    Dim LastFragmentSize As Long
    Dim FragmentBuffer() As Byte
    Dim OriginalFileNumber As Long
    Dim FragmentFileNumber As Long
    Dim Header As FragmentHeader
    Dim intPB As Integer ' how much add to pb.value per each file
    Dim NumberOfSplits As Long
    
    SplitCancealed = False
    
    ReDim FragmentBuffer(1 To FragmentSize) As Byte
        
    On Error GoTo SplitErr
    
    FileSize = FileLen(OriginalFilePathName)
    SplitNumber = Fix(FileSize / FragmentSize)
    intPB = Fix(FrmSplit5.pb.Max / SplitNumber)

    LastFragmentSize = FileSize - (FragmentSize * SplitNumber)
    
    If LastFragmentSize = 0 Then
        NumberOfSplits = SplitNumber
    Else
        NumberOfSplits = SplitNumber + 1
    End If
    
    Header.FragmentSize = FragmentSize
    Header.NumberOfFragments = NumberOfSplits
    Header.DateOfSplitting = Date
    Header.AuthorComment = AUTHOR_COMMENTS
    Header.OriginalFileName = LTrim(GetFileName(OriginalFilePathName))
    Header.OriginalFileSize = FileSize
    Randomize
    Header.UniqueIdentifier = CLng(Round(Rnd * 1000000000, 0))
    
    OriginalFileNumber = FreeFile
    Open OriginalFilePathName For Binary As #OriginalFileNumber
    
    For i = 1 To SplitNumber
        Get #OriginalFileNumber, , FragmentBuffer
        FragmentName = SaveToFolder & "\" & CollectionName & "(" _
            & CStr(i) & ")" & "." & FragmentExtension
        FragmentFileNumber = FreeFile
        Open FragmentName For Binary As #FragmentFileNumber
        Header.FragmentNumber = i
        Put #FragmentFileNumber, , Header ' 256 bytes
        Put #FragmentFileNumber, , FragmentBuffer
        Close #FragmentFileNumber
        FrmSplit5.pb.Value = FrmSplit5.pb.Value + intPB
        DoEvents
        If SplitCancealed Then
            Split = False
            Exit Function
        End If
    Next
    
    If LastFragmentSize = 0 Then
        Close #OriginalFileNumber
        FrmSplit5.pb.Value = FrmSplit5.pb.Max
        Split = True
        Exit Function
    End If
    
    ReDim FragmentBuffer(1 To LastFragmentSize)
    
    FragmentName = SaveToFolder & "\" & CollectionName & "(" _
        & CStr(SplitNumber + 1) & ")" & "." & FragmentExtension
    FragmentFileNumber = FreeFile
    Get #OriginalFileNumber, , FragmentBuffer
    Open FragmentName For Binary As #FragmentFileNumber
    Header.FragmentNumber = SplitNumber + 1
    Put #FragmentFileNumber, , Header ' 256 bytes
    Put #FragmentFileNumber, , FragmentBuffer
    Close #FragmentFileNumber
    Close #OriginalFileNumber
    
    FrmSplit5.pb.Value = FrmSplit5.pb.Max
    Split = True
    Exit Function
    
SplitErr:
    Close
    Split = False
    
End Function

