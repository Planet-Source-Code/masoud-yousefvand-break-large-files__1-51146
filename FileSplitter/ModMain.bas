Attribute VB_Name = "ModMain"
Option Explicit

Public Const GreetingMessage As String = "Dedicate to PSC users."

Public Const HEADER_SIZE As Long = 256 ' Bytes
Public Const MAX_NUM_OF_SPLITS As Long = 9999
Public Const DEFAULT_EXTENSION As String = "spt"

'Public UseFloppy As Boolean

Public FrmSplit3Numor3Size As Integer ' 31 for num and 32 for size
Public SplitCancealed As Boolean
Public AssembleCancealed As Boolean

Public FragmentsPath As String ' Where are fragments (assembler uses)
Public OriginalFileNamePath As String
Public DestinationSplitFolder As String
Public CollectionName As String
Public AssembleFileName As String ' transfer name from frmass1 to frmass2
Public SelectedCollection As String ' assembler uses this
Public UserSizeOfPieces As Long
Public UserNumOfPieces As Long
Public Const FloppySize As Long = 1457664
Public Const KB As Long = 1024
Public Const MB As Long = 1048576
Public Const GB As Long = 1073741824

Sub Main()

    On Error Resume Next
    
    FrmScreenSplash.Show

End Sub

Public Function GetFileName(ByVal FilePathName As String) As String

    Dim strTemp As String
    Dim i As Integer

    strTemp = StrReverse(FilePathName)
    i = InStr(1, strTemp, "\", vbTextCompare)
    GetFileName = Right(FilePathName, (i - 1))

End Function

Public Function GetFilePath(ByVal FilePathName As String) As String

    Dim strTemp As String
    Dim i As Integer

    strTemp = StrReverse(FilePathName)
    i = InStr(1, strTemp, "\", vbTextCompare)
    GetFilePath = Left(FilePathName, (Len(strTemp) - i))

End Function

Public Function FileKbMbGb(ByVal FileSize As Long) As String

    Dim k As Long
    Dim m As Long
    Dim g As Long
    Dim temp As Double
    Dim temp2 As Double
    
    If FileSize < KB Then
        FileKbMbGb = FileSize & " Bytes."
        Exit Function
    End If
    If FileSize < MB Then
        FileKbMbGb = CStr(Fix(FileSize / KB)) & " KB, and " & _
            CStr(FileSize Mod KB) & " Bytes."
        Exit Function
    End If
    If FileSize < GB Then
        temp = FileSize Mod MB
        FileKbMbGb = CStr(Fix(FileSize / MB)) & " MB, and " & _
            CStr(Fix(temp / KB)) & " KB, and " & CStr(temp Mod KB) & _
            " Bytes."
        Exit Function
    End If
    If FileSize > GB Then
        temp = FileSize Mod GB
        temp2 = temp Mod MB
        FileKbMbGb = CStr(Fix(FileSize / GB)) & " GB, and " & _
        CStr(Fix(temp / MB)) & " MB, and " & CStr(Fix(temp2 / KB)) & _
        " KB, and " & CStr(temp2 Mod KB) & " Bytes."
        Exit Function
    End If
    
End Function

Public Sub RESET_ALL_VARIABLES()

    OriginalFileNamePath = ""
    UserSizeOfPieces = 0
    UserNumOfPieces = 0

End Sub

Public Sub ExitWizard()

    On Error Resume Next
    
    Dim Frm As Form
    
    For Each Frm In Forms
        Unload Frm
    Next
    End
    
End Sub

Public Function FileExtension(ByVal FilePathName As String) As String

    Dim strTemp As String
    Dim i As Integer

    strTemp = StrReverse(FilePathName)
    i = InStr(1, strTemp, ".", vbTextCompare)
    FileExtension = Right(FilePathName, (i - 1))

End Function
