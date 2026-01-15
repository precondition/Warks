Option Explicit
'
' Registry location via SaveSetting/GetSetting:
' HKCU\Software\VB and VBA Program Settings\${APP_NAME}\LocalMarks\${ActiveDocument.FullName}
Private Const APP_NAME As String = "Warks"

' Preview length (for line context)
Private Const PREVIEW_LEN As Long = 60

Private marksDict As Object        ' Scripting.Dictionary

Public Sub MarkSet(Optional ByVal markName As String = "")

    If Len(markName) < 1 Then
        markName = InputBox("Set mark:", "Set Mark")

        If markName = "'" Then
            MsgBox "Reserved mark name: Cannot manually set the ' (apostrophe) mark!", vbExclamation, APP_NAME
            Exit Sub
        End If

    End If

    If Len(markName) < 1 Then Exit Sub

    If marksDict Is Nothing Then
        Set marksDict = CreateObject("Scripting.Dictionary")
    End If

    marksDict(markName) = Selection.Range.Start

End Sub


Public Sub MarkJump(Optional ByVal markName As String = "")
    if Len(markName) < 1 Then
        markName = InputBox("Jump to mark:", "Jump to Mark")
    End If
    If Len(markName) < 1 Then Exit Sub

    If marksDict Is Nothing Then
        Set marksDict = CreateObject("Scripting.Dictionary")
    End If

    If Not marksDict.Exists(markName) Then
        MsgBox "Mark not set: " & markName, vbExclamation, APP_NAME
        Exit Sub
    End If

    ' This must be called before we update the ' mark incase `MarkJump` is
    ' called with apostrophe.
    Dim pos As Long: pos = marksDict(markName)

    ' Push current position to the reserved ' mark for back jumps.
    MarkSet("'")

    Selection.SetRange pos, pos
    ActiveWindow.ScrollIntoView Selection.Range, True
End Sub


Public Sub MarkJumpLine(Optional ByVal markName As String = "")
    If Len(markName) < 1 Then
        markName = InputBox("Jump To mark (start of line):", "Jump To Mark (Line)")
    End If
    If Len(markName) < 1 Then Exit Sub

    Call MarkJump(markName)
    Selection.HomeKey Unit:=wdLine
End Sub


Private Function GetPreview(ByVal charPos As Long) As String
    Dim docEnd As Long
    docEnd = ActiveDocument.Content.End

    ' Clamp start
    If charPos < 0 Then charPos = 0
    If charPos > docEnd Then charPos = docEnd

    ' Compute and clamp end
    Dim endPos As Long
    endPos = charPos + PREVIEW_LEN/2
    If endPos > docEnd Then endPos = docEnd

    ' Build range and get text
    Dim previewRange As Word.Range
    Set previewRange = ActiveDocument.Range(Start:=charPos - PREVIEW_LEN / 2, End:=endPos)
    Dim s As String
    s = previewRange.Text

    ' Normalize preview: replace CR/LF/TAB and strip other control chars
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    ' Strip any remaining control chars (incl. table cell end Chr(7), FS/RS/US, etc.)
    Dim i As Long, ch As String
    Dim out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If AscW(ch) >= 32 Then
            out = out & ch
        ElseIf ch = " " Then
            out = out & " "
        End If
    Next
    ' Collapse spaces and trim
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    out = Trim$(out)

    ' Enforce max length
    If Len(out) > PREVIEW_LEN Then
        out = Left$(out, PREVIEW_LEN)
    End If

    GetPreview = out
End Function

Public Sub MarkList()
    Dim k As Variant
    Dim s As String
    Dim absCharPos As Long
    Dim selectedRange As Word.Range
    Dim pageNumber
    Dim verticalPosInPage
    Dim horizontalPosInPage

    If marksDict Is Nothing Then
        MsgBox "No marks set.", vbInformation, APP_NAME
        Exit Sub
    ElseIf marksDict.Count = 0 Then
        MsgBox "No marks set.", vbInformation, APP_NAME
        Exit Sub
    End If

    s = "Marks (" & marksDict.Count & "):" & vbCrLf & vbCrLf

    For Each k In marksDict.Keys
        absCharPos = CLng(marksDict(k))

        Set selectedRange = ActiveDocument.Range(Start:=absCharPos, End:=absCharPos)
        pageNumber = selectedRange.Information(wdActiveEndAdjustedPageNumber)
        verticalPosInPage = selectedRange.Information(wdVerticalPositionRelativeToPage)
        horizontalPosInPage = selectedRange.Information(wdHorizontalPositionRelativeToPage)

        s = s & "• " & k & " @ p" & pageNumber _
            & " (" & verticalPosInPage & ";" & horizontalPosInPage & "): " _
            & GetPreview(absCharPos) & vbCrLf
    Next k

    MsgBox s, vbOKOnly, APP_NAME
End Sub

' ========== Convenience: A/B quick set/jump for non-interactive keyboard shortcuts ==========

Public Sub MarkJumpBack(): MarkJump "'": End Sub
Public Sub MarkJumpLineBack(): MarkJumpLine "'": End Sub
Public Sub MarkSetUpperCaseA(): MarkSet "A": End Sub
Public Sub MarkSetLowerCaseA(): MarkSet "a": End Sub
Public Sub MarkSetUpperCaseB(): MarkSet "B": End Sub
Public Sub MarkSetLowerCaseB(): MarkSet "b": End Sub
Public Sub MarkJumpUpperCaseA(): MarkJump "A": End Sub
Public Sub MarkJumpLowerCaseA(): MarkJump "a": End Sub
Public Sub MarkJumpUpperCaseB(): MarkJump "B": End Sub
Public Sub MarkJumpLowerCaseB(): MarkJump "b": End Sub
