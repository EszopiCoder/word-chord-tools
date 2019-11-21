Attribute VB_Name = "modFont"
Option Explicit

Public Sub TestFormatFx()
    Call ChordMarkerDoc
    Call FormatChords(RGB(0, 0, 0), False, False)
End Sub
Public Function BoldChords(ByVal isBold As Boolean) As Long
    
    Dim ChordCount As Long
    
    ' Go to the first line in the document
    Selection.GoTo wdGoToLine, 1
    ' Format and search for |Chord| and remove |
    Application.ScreenUpdating = False
    With Selection.Find
        .ClearFormatting
        .Text = "\|?*\|"
        .Wrap = wdFindContinue
        .Forward = True
        .MatchWildcards = True
        Do While .Execute
            Selection.Text = Replace(Selection.Text, "|", "")
            Selection.Font.Bold = isBold
            Selection.Collapse Direction:=wdCollapseEnd
            ChordCount = ChordCount + 1
        Loop
    End With
    Application.ScreenUpdating = True
    BoldChords = ChordCount - 1
End Function
Public Function FormatChords(ByVal lngColor As Long, _
    ByVal isBold As Boolean, ByVal isItalic As Boolean) As Long
    
    Dim ChordCount As Long
    
    ' Go to the first line in the document
    Selection.GoTo wdGoToLine, 1
    ' Format and search for |Chord| and remove |
    Application.ScreenUpdating = False
    With Selection.Find
        .ClearFormatting
        .Text = "\|?*\|"
        .Wrap = wdFindContinue
        .Forward = True
        .MatchWildcards = True
        Do While .Execute
            Selection.Text = Replace(Selection.Text, "|", "")
            Selection.Font.TextColor.RGB = lngColor
            Selection.Font.Bold = isBold
            Selection.Font.Italic = isItalic
            Selection.Collapse Direction:=wdCollapseEnd
            ChordCount = ChordCount + 1
        Loop
    End With
    Application.ScreenUpdating = True
    FormatChords = ChordCount - 1
End Function
Public Sub ChordMarkerDoc()
    Dim objRegEx As Object
    Dim strText As String
    
    ' Set variables
    strText = ActiveDocument.Range.Text
    
    ' Save formatting
    ActiveDocument.Range.Select
    Selection.CopyFormat
    
    ' Set up RegEx
    Set objRegEx = CreateObject("VBScript.RegExp")
    With objRegEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "([ABCDEFG][b#\u266F\u266D]?[m]?[\(]?(2|5|6|7|9|11|13|6\/9|7\-5|7\-9|7\#5|7\#9|7\+5|7\+9|7b5|7b9|7sus2|7sus4|add2|add4|add9|aug|dim|dim7|m\|maj7|m6|m7|m7b5|m9|m11|m13|maj7|maj9|maj11|maj13|mb5|m|sus|sus2|sus4)?(\))?)(?=\s|\.|\)|-|\/)"
    End With
    
    ' Format chords in the original string to: |Chord|
    ' Note this is the entire chord
    strText = objRegEx.Replace(strText, "|$1|")
    
    ' Replace text and paste formatting
    Selection.Text = Left(strText, Len(strText) - 1)
    Selection.PasteFormat
    Selection.Collapse wdCollapseEnd
End Sub
Public Sub ChordMarkerSelection()
    Dim objRegEx As Object
    Dim strText As String
    
    ' Set variables
    strText = Selection.Text
    
    ' Save formatting
    Selection.CopyFormat
    
    ' Set up RegEx
    Set objRegEx = CreateObject("VBScript.RegExp")
    With objRegEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "([ABCDEFG][b#\u266F\u266D]?[m]?[\(]?(2|5|6|7|9|11|13|6\/9|7\-5|7\-9|7\#5|7\#9|7\+5|7\+9|7b5|7b9|7sus2|7sus4|add2|add4|add9|aug|dim|dim7|m\|maj7|m6|m7|m7b5|m9|m11|m13|maj7|maj9|maj11|maj13|mb5|m|sus|sus2|sus4)?(\))?)(?=\s|\.|\)|-|\/)"
    End With
    
    ' Format chords in the original string to: |Chord|
    ' Note this is the entire chord
    strText = objRegEx.Replace(strText, "|$1|")
    
    ' Replace text and paste formatting
    ' Remove excess carriage return at the end of string
    Selection.Text = strText
    Selection.PasteFormat
    Selection.Collapse wdCollapseEnd
End Sub
