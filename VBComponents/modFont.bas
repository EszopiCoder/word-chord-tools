Attribute VB_Name = "modFont"
Option Explicit

Public Sub TestFormatFx()
    Application.ScreenUpdating = False
    Call ChordMarkerDoc
    Call FormatChords(RGB(0, 0, 0), False, False)
    Application.ScreenUpdating = True
End Sub
Public Function BoldChords(ByVal isBold As Boolean)
    
    ' Go to the first line in the document
    Selection.GoTo wdGoToLine, 1
    ' Format and search for |Chord| and remove |
    With Selection.Find
        .ClearFormatting
        .Text = "\|([A-Za-z0-9]*)\|"
        .Wrap = wdFindContinue
        .Forward = True
        .MatchWildcards = True
        .Replacement.Font.Bold = isBold
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
    End With
End Function
Public Function FormatChords(ByVal lngColor As Long, _
    ByVal isBold As Boolean, ByVal isItalic As Boolean)
    
    ' Go to the first line in the document
    Selection.GoTo wdGoToLine, 1
    ' Format and search for |Chord| and remove |
    With Selection.Find
        .ClearFormatting
        .Text = "\|([A-Za-z0-9]*)\|"
        .Wrap = wdFindContinue
        .Forward = True
        .MatchWildcards = True
        .Replacement.Font.Bold = isBold
        .Replacement.Font.Italic = isItalic
        .Replacement.Font.TextColor.RGB = lngColor
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
    End With
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
        .Pattern = "([A-G][b#\u266F\u266D]?[m]?[\(]?(2|5|6|7|9|11|13|6\/9|7\-5|7\-9|7\#5|7\#9|7\+5|7\+9|7b5|7b9|7sus2|7sus4|add2|add4|add9|aug|dim|dim7|m\|maj7|m6|m7|m7b5|m9|m11|m13|maj7|maj9|maj11|maj13|mb5|m|sus|sus2|sus4)?(\))?)(?=\s\s|\.|\)|-|\/|\r|\n|\s[A-G]|\s\W)"
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
        .Pattern = "([A-G][b#\u266F\u266D]?[m]?[\(]?(2|5|6|7|9|11|13|6\/9|7\-5|7\-9|7\#5|7\#9|7\+5|7\+9|7b5|7b9|7sus2|7sus4|add2|add4|add9|aug|dim|dim7|m\|maj7|m6|m7|m7b5|m9|m11|m13|maj7|maj9|maj11|maj13|mb5|m|sus|sus2|sus4)?(\))?)(?=\s\s|\.|\)|-|\/|\r|\n|\s[A-G]|\s\W)"
    End With
    
    ' Format chords in the original string to: |Chord|
    ' Note this is the entire chord
    strText = objRegEx.Replace(strText, "|$1|")
    
    ' Replace text and paste formatting
    Selection.Text = strText
    Selection.PasteFormat
    Selection.Collapse wdCollapseEnd
End Sub
