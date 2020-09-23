Attribute VB_Name = "ModAssembler"
Option Explicit

'Private Const HEADER_SIZE As Long = 256 ' Bytes

Public Function Assembler(ByVal TargetFilePathName As String, _
    ByVal FragmentsPathFileArray As Variant) As Boolean
    
    Dim i As Long
    Dim j As Long
    Dim FileBuffer() As Byte
    Dim FragmentsNumber As Long
    Dim FragmentFileLen As Long
    Dim TargetFileNumber As Long
    Dim Header As FragmentHeader
    Dim HeaderX() As FragmentHeader
    Dim FragmentFileNumber As Long
    Dim intPB As Integer ' how much add to pb.value per each file

    On Error GoTo AssemblerErr

    If Not IsArray(FragmentsPathFileArray) Then GoTo AssemblerErr
    FragmentsNumber = UBound(FragmentsPathFileArray)
    
    ReDim HeaderX(1 To FragmentsNumber)
    
    intPB = Fix(FrmAssemble2.pb.Max / FragmentsNumber)
    
    TargetFileNumber = FreeFile
    Open TargetFilePathName For Binary As #TargetFileNumber
    
    For i = 1 To FragmentsNumber
        FragmentFileLen = FileLen(FragmentsPathFileArray(i))
        ReDim FileBuffer(1 To (FragmentFileLen - HEADER_SIZE))
        FragmentFileNumber = FreeFile
        Open FragmentsPathFileArray(i) For Binary As #FragmentFileNumber
        Get #FragmentFileNumber, , Header
        
        HeaderX(i).AuthorComment = Header.AuthorComment
        HeaderX(i).DateOfSplitting = Header.DateOfSplitting
        HeaderX(i).FragmentNumber = Header.FragmentNumber
        HeaderX(i).FragmentSize = Header.FragmentSize
        HeaderX(i).NumberOfFragments = Header.NumberOfFragments
        HeaderX(i).OriginalFileName = Header.OriginalFileName
        HeaderX(i).OriginalFileSize = Header.OriginalFileSize
        HeaderX(i).UniqueIdentifier = Header.UniqueIdentifier
        
        Get #FragmentFileNumber, , FileBuffer
        Put #TargetFileNumber, , FileBuffer
        Close #FragmentFileNumber
        DoEvents
        If AssembleCancealed Then
            Assembler = False
            Exit Function
        End If
        FrmAssemble2.pb.Value = FrmAssemble2.pb.Value + intPB
    Next
    
    Close #TargetFileNumber
    FrmAssemble2.pb.Value = FrmAssemble2.pb.Max
    
    For i = 1 To Fix(FragmentsNumber / 2) + 1
        For j = FragmentsNumber To Fix(FragmentsNumber / 2) Step -1
            If HeaderX(i).UniqueIdentifier <> HeaderX(j).UniqueIdentifier Then
                Assembler = False
                Exit Function
            End If
        DoEvents
        Next
    Next
        
    Assembler = True
    Exit Function

AssemblerErr:
    Close
    Assembler = False

End Function
