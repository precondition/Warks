Option Explicit
'
' Registry location via SaveSetting/GetSetting:
' HKCU\Software\VB and VBA Program Settings\${APPNAME}\Marks\${ActiveDocument.FullName}
Private Const APPNAME As String = "Warks"
Private Const SECTION As String = "Marks"
' The symbols <, /, \, and > are illegal in file names so this should not clash
' with local filepath-based keys.
Private Const GLOBAL_MARKS_KEY As String = "</GLOBAL\>"

' Preview length (for line context)
Private Const PREVIEW_LEN As Long = 60

Private localMarksDict As Object
Private globalMarksDict As Object
Private loadedDocKey As String

' Chr$() cannot be used in constant expressions
Private Const FIELD_SEP_CODE As Long = 31
Private Const RECORD_SEP_CODE   As Long = 30

Private Function FIELD_SEP$()
    FIELD_SEP = Chr$(FIELD_SEP_CODE)
End Function

Private Function RECORD_SEP$()
    RECORD_SEP = Chr$(RECORD_SEP_CODE)
End Function

Private Function CurrentDocKey() As String
    Dim s As String
    On Error Resume Next
    s = ActiveDocument.FullName
    If Len(s) = 0 Then s = ActiveDocument.Name ' unsaved doc fallback
    On Error GoTo 0
    CurrentDocKey = s
End Function



Private Function SerializeMarks(ByVal d As Object) As String
    Dim k As Variant, v As Variant
    Dim serializedBlob As String: serializedBlob = ""

    For Each k In d.Keys
        If CStr(k) <> "'" Then
            v = d(k)
            ' v(0) = absCharPos, v(1) = docPath (may be empty)
            serializedBlob = serializedBlob & CStr(k) & FIELD_SEP & CStr(CLng(v(0)))
            If Not IsEmpty(v) And UBound(v) >= 1 And Len(CStr(v(1))) > 0 Then
                serializedBlob = serializedBlob & FIELD_SEP & CStr(v(1))
            End If
            serializedBlob = serializedBlob & RECORD_SEP
        End If
    Next

    SerializeMarks = serializedBlob
End Function

Private Function DeserializeMarksToDict(ByVal serializedBlob As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    If Len(serializedBlob) = 0 Then
        Set DeserializeMarksToDict = d
        Exit Function
    End If

    Dim records() As String
    records = Split(serializedBlob, RECORD_SEP)

    Dim i As Long, fields() As String
    Dim fieldCount As Long
    Dim markName As String, posStr As String
    Dim absCharPos As Long, docPath As String

    For i = LBound(records) To UBound(records)
        If Len(records(i)) = 0 Then GoTo NextRecord

        ' Limit to at most 3 fields: markName | pos | path
        fields = Split(records(i), FIELD_SEP, 3)
        fieldCount = UBound(fields) - LBound(fields) + 1
        If fieldCount < 2 Then GoTo NextRecord

        markName = fields(0)
        If markName = "'" Then GoTo NextRecord

        posStr = fields(1)
        absCharPos = CLng(Val(posStr))    ' tolerant conversion

        If fieldCount >= 3 Then
            docPath = fields(2)
        Else
            docPath = ""
        End If

        d(markName) = Array(absCharPos, docPath)
NextRecord:
    Next i

    Set DeserializeMarksToDict = d
End Function

Private Sub EnsureMarksLoaded()
    Dim blob As String
    Dim docKey As String: docKey = CurrentDocKey()
    If localMarksDict Is Nothing Or loadedDocKey <> docKey Then
        blob = GetSetting(APPNAME, "Marks", docKey, "")
        Set localMarksDict = DeserializeMarksToDict(blob)
        loadedDocKey = docKey
    End If
    If globalMarksDict Is Nothing Or loadedDocKey <> docKey Then
        blob = GetSetting(APPNAME, "Marks", "</GLOBAL\>", "")
        Set globalMarksDict = DeserializeMarksToDict(blob)
        loadedDocKey = docKey
    End If
End Sub

Public Sub MarkSetTo(Optional ByVal markName As String = "")
    EnsureMarksLoaded
    If Len(markName) < 1 Then
        markName = InputBox("Set mark:", APPNAME)

        If markName = "'" Then
            MsgBox "Reserved mark name: Cannot manually set the ' (apostrophe) mark!", vbExclamation, APPNAME
            Exit Sub
        End If

    End If

    If Len(markName) < 1 Then Exit Sub

    If IsGlobalMark(markName) Then
        globalMarksDict(markName) = Array(Selection.Range.Start, CurrentDocKey())
        SaveSetting APPNAME, "Marks", "</GLOBAL\>", SerializeMarks(globalMarksDict)
    Else
        localMarksDict(markName) = Array(Selection.Range.Start, "")
        SaveSetting APPNAME, "Marks", CurrentDocKey(), SerializeMarks(localMarksDict)
    End If

End Sub

Private Function GetPreview(ByVal charPos As Long) As String
    Dim doc As Document: Set doc = ActiveDocument
    Dim docStart As Long: docStart = doc.Content.Start
    Dim docEnd As Long:   docEnd   = doc.Content.End

    ' Clamp target
    If charPos < docStart Then charPos = docStart
    If charPos >= docEnd Then charPos = docEnd

    ' Use integer half-width and clamp start/end
    Dim half As Long: half = PREVIEW_LEN \ 2
    Dim startPos As Long: startPos = charPos - half
    If startPos < docStart Then startPos = docStart
    Dim endPos As Long: endPos = charPos + half
    If endPos >= docEnd Then endPos = docEnd

    Dim previewRange As Range
    Set previewRange = doc.Range(Start:=startPos, End:=endPos)

    Dim s As String: s = previewRange.Text
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")

    ' Strip remaining control chars, collapse spaces
    Dim out As String, i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If AscW(ch) >= 32 Then out = out & ch Else out = out & " "
    Next
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    out = Trim$(out)

    If Len(out) > PREVIEW_LEN Then out = Left$(out, PREVIEW_LEN)
    GetPreview = out
End Function

Public Sub MarkList()
    EnsureMarksLoaded
    Dim markName As Variant
    Dim s As String
    Dim absCharPos As Long
    Dim selectedRange As Word.Range
    Dim pageNumber
    Dim verticalPosInPage
    Dim horizontalPosInPage

    Dim doc As Document: Set doc = ActiveDocument
    Dim docStart As Long: docStart = doc.Content.Start
    Dim docEnd As Long:   docEnd   = doc.Content.End

    s = "Local Marks (" & localMarksDict.Count & "):" & vbCrLf & vbCrLf

    For Each markName In localMarksDict.Keys
        absCharPos = CLng(localMarksDict(markName)(0))
        ' Skip invalid marks. Ideally, we would delete them but we are
        ' iterating on keys.
        If docStart <= absCharPos And absCharPos < docEnd Then
            Set selectedRange = ActiveDocument.Range(Start:=absCharPos, End:=absCharPos)
            pageNumber = selectedRange.Information(wdActiveEndAdjustedPageNumber)
            verticalPosInPage = selectedRange.Information(wdVerticalPositionRelativeToPage)
            horizontalPosInPage = selectedRange.Information(wdHorizontalPositionRelativeToPage)

            s = s & "• " & markName & " @ p" & pageNumber _
                & " (" & verticalPosInPage & " ; " & horizontalPosInPage & "): " _
                & GetPreview(absCharPos) & vbCrLf
        End If
    Next markName

    s = s & vbCrLf & vbCrLf & "Global Marks (" & globalMarksDict.Count & "):" & vbCrLf & vbCrLf

    For Each markName In globalMarksDict.Keys
        absCharPos = CLng(globalMarksDict(markName)(0))
        ' Skip invalid marks. Ideally, we would delete them but we are
        ' iterating on keys.
        If docStart <= absCharPos And absCharPos < docEnd Then

            Set selectedRange = ActiveDocument.Range(Start:=absCharPos, End:=absCharPos)
            Dim docPath as String: docPath = globalMarksDict(markName)(1)
            pageNumber = selectedRange.Information(wdActiveEndAdjustedPageNumber)
            verticalPosInPage = selectedRange.Information(wdVerticalPositionRelativeToPage)
            horizontalPosInPage = selectedRange.Information(wdHorizontalPositionRelativeToPage)

            s = s & "• " & markName & " @ " & docPath & " p" & pageNumber _
                & " (" & verticalPosInPage & " ; " & horizontalPosInPage & "): " _
                &  vbCrLf
        End If
    Next markName

    MsgBox s, vbOKOnly, APPNAME
End Sub


Public Sub MarkJumpTo(Optional ByVal markName As String = "")
    Dim docStart As Long
    Dim docEnd As Long
    EnsureMarksLoaded
    If Len(markName) < 1 Then Exit Sub


    Dim pos As Long

    If IsGlobalMark(markName) Then
        If Not globalMarksDict.Exists(markName) Then
            MsgBox "Mark not set: " & markName, vbExclamation, APPNAME
            Exit Sub
        End If

        Dim docPath As String: docPath = globalMarksDict(markName)(1)
        Dim doc As Document: Set doc = OpenDocIfNeeded(docPath)

        If doc Is Nothing Then
            MsgBox "Cannot open document for mark '" & markName & "':" & vbCrLf & docPath, vbCritical, APPNAME
            Exit Sub
        End If

        docStart = doc.Content.Start
        docEnd   = doc.Content.End

        pos = globalMarksDict(markName)(0)

        If pos < docStart Or docEnd <= pos Then
            MsgBox "Mark out of bounds: Cannot jump to mark " & markName & " in file " & docPath, vbExclamation, APPNAME
            globalMarksDict.Remove markName
            Exit Sub
        End If

        doc.Activate
    Else
        If Not localMarksDict.Exists(markName) Then
            MsgBox "Mark not set: " & markName, vbExclamation, APPNAME
            Exit Sub
        End If
        ' This must be called before we update the ' mark in case `MarkJumpTo` is
        ' called with apostrophe.
        pos = localMarksDict(markName)(0)


        docStart = ActiveDocument.Content.Start
        docEnd   = ActiveDocument.Content.End

        If pos < docStart Or docEnd <= pos Then
            MsgBox "Mark out of bounds: Cannot jump to mark " & markName & " in file " & docPath, vbExclamation, APPNAME
            localMarksDict.Remove markName
            Exit Sub
        End If
    End If

    ' Push current position to the reserved ' mark for back jumps.
    MarkSetTo("'")

    Selection.SetRange pos, pos
    ' Just to be sure that no chunk is selected.
    Selection.Collapse wdCollapseStart

    ' ScrollIntoView does not work in reading layout so we must temporarily
    ' switch back to the default print layout, scroll the selection into view
    ' and then switch back into the reading layout. Unfortunately a full freeze
    ' of the screen is not possible but we can reduce the jarring flashing a
    ' little bit with Application.ScreenUpdating = False.
    Dim isReadMode : isReadMode = ActiveDocument.ActiveWindow.View.ReadingLayout
    Application.ScreenUpdating = Not isReadMode
    ActiveDocument.ActiveWindow.View.ReadingLayout = False

    ActiveWindow.ScrollIntoView Selection.Range, True

    ' Restore layout.
    ActiveDocument.ActiveWindow.View.ReadingLayout = isReadMode
    Application.ScreenUpdating = True

End Sub


Public Sub MarkJumpToLine(Optional ByVal markName As String = "")
    Call MarkJumpTo(markName)
    Selection.HomeKey Unit:=wdLine
End Sub

Private Function IsGlobalMark(ByVal name As String) As Boolean
    If Len(name) = 0 Then IsGlobalMark = False: Exit Function
    IsGlobalMark = (Left$(name, 1) Like "[A-Z]")
End Function


Private Function FindOpenDocByPath(ByVal fullPath As String) As Document
    Dim dd As Document
    For Each dd In Application.Documents
        If StrComp(dd.FullName, fullPath, vbTextCompare) = 0 Then
            Set FindOpenDocByPath = dd: Exit Function
        End If
    Next dd
    Set FindOpenDocByPath = Nothing
End Function

Private Function OpenDocIfNeeded(ByVal fullPath As String) As Document
    Dim d As Document: Set d = FindOpenDocByPath(fullPath)
    If d Is Nothing Then
        On Error Resume Next
        Set d = Documents.Open(FileName:=fullPath, ReadOnly:=True, AddToRecentFiles:=True)
        On Error GoTo 0
    End If
    Set OpenDocIfNeeded = d
End Function



' Multi-arity functions/subroutines do not surface in Word's list of usable
' macros so we must create 0-arity wrappers.

Public Sub MarkJump()
    Call MarkJumpTo(InputBox("Jump to mark:", APPNAME))
End Sub

Public Sub MarkSet()
    Call MarkSetTo(InputBox("Set mark:", APPNAME))
End Sub

Public Sub MarkJumpLine()
    Call MarkJumpToLine(InputBox("Jump To mark (start of line):", APPNAME))
End Sub

Public Sub MarkJumpBack():      MarkJumpTo      "'": End Sub
Public Sub MarkJumpLineBack():  MarkJumpToLine  "'": End Sub

' Wrappers for quick toggling between two marks without an interactive prompt.
Public Sub MarkJumpGlobalA(): MarkJumpTo "A": End Sub
Public Sub MarkJumpLocalA():  MarkJumpTo "a": End Sub
Public Sub MarkSetGlobalA():  MarkSetTo "A":  End Sub
Public Sub MarkSetLocalA():   MarkSetTo "a":  End Sub

Public Sub MarkJumpGlobalB(): MarkJumpTo "B": End Sub
Public Sub MarkJumpLocalB():  MarkJumpTo "b": End Sub
Public Sub MarkSetGlobalB():  MarkSetTo "B":  End Sub
Public Sub MarkSetLocalB():   MarkSetTo "b":  End Sub
