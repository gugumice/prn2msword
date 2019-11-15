Option Explicit

Function getDirectory(Optional ByVal path As String = "c:\")
    'Get dir to get files from
    On Error Resume Next
    Dim fDlg As FileDialog
    Dim fDir As String
    Set fDlg = Application.FileDialog(msoFileDialogFolderPicker)
    With fDlg
       .Title = "Izvēlies vietu"
       .AllowMultiSelect = False
       .InitialFileName = path
    End With
    fDlg.Show
    getDirectory = fDlg.SelectedItems(1)
End Function

Function toASCII(s As String) As String
'For debug. Returns the string to its respective ascii numbers
   Dim i As Integer
   For i = 1 To Len(s)
      'toASCII = toASCII & " " & CStr(Asc(Mid(s, i, 1)))
      toASCII = toASCII & " " & CStr(Asc(Mid(s, i, 1)))
   Next i
End Function
Function filterLines(f As String, s As String) As Boolean
'Filter lines beginning with comma delimited keywords (f)
    Dim i As Integer
    Dim strFilter() As String
    filterLines = False
    strFilter = Split(f, ",")
    For i = 0 To UBound(strFilter)
        'Debug.Print strFilter(i)
        If Left(s, Len(strFilter(i))) = strFilter(i) Then
            filterLines = True
            Exit For
        End If
    Next i
End Function
Sub Text2Doc(s As String, Optional ByVal pm As Boolean = False)
'Write to doc. pm - terminate line with soft line breake or paragraph mark
    If Len(s) > 0 Then
        With ActiveDocument.Content
            .InsertAfter Text:=s
            If pm Then
                .InsertParagraphAfter
            Else
                .InsertAfter vbVerticalTab
            End If
        End With
    End If
End Sub
Sub prnReport(fileName)
    Dim fso As Object
    Dim objFile As Object
    Dim strLine As String
    Dim intRowCnt As Integer
    Dim boolSkipBlank As Boolean
    Dim boolSkipLine As Boolean
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFile = fso.OpenTextFile(fileName, 1)
    boolSkipBlank = False
    boolSkipLine = False
    Const FILTER_KEYWORDS = "%%[,Answered by the,SUBJECT:,PROTOCOLE COMPLET,FAX,FROM,COMMENT,TO,Dr. LABORATORIJA,Zemitana,1006 RIGA,7545052,SUBCONTRACTING,Results and comments comming,Answered by,Hospital,Results and comments"
    Do Until objFile.AtEndOfStream
        boolSkipLine = False
        strLine = objFile.ReadLine
        'filter consequetive blank lines
        If Len(strLine) = 0 Then
            If boolSkipBlank Then
                boolSkipLine = True
            Else
                boolSkipBlank = True
            End If
        Else
            boolSkipBlank = False
        End If
        If filterLines(FILTER_KEYWORDS, strLine) Then boolSkipLine = True
        If Left(strLine, 2) = Chr(12) & Chr(26) Then
            Text2Doc String(40, "_"), True
            boolSkipLine = True
        End If
        If Not boolSkipLine Then
            intRowCnt = intRowCnt + 1
            'Debug.Print intRowCnt, strLine ', toASCII(strLine)
            'Debug.Print strLine
            Text2Doc strLine
        End If
    Loop
End Sub
