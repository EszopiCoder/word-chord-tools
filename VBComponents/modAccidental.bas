Attribute VB_Name = "modAccidental"
Option Explicit

Public Sub TestUnicodeFx()
    Call ChordMarkerDoc
    MsgBox UnicodeChords(True)
End Sub
Public Function UnicodeChords(ByVal blnUnicode As Boolean) As Long
    
    Dim ChordCount As Long
    
    ' Go to the first line in the document
    Selection.GoTo wdGoToLine, 1
    ' Format and search for |Chord| and remove |
    Application.ScreenUpdating = False
    With Selection.Find
        .ClearFormatting
        .Text = "\|[A-Za-z0-9]*\|"
        .Wrap = wdFindContinue
        .Forward = True
        .MatchWildcards = True
        Do While .Execute
            Selection.Text = Replace(Selection.Text, "|", "")
            If blnUnicode = True Then ' Output unicode
                Selection.Text = Replace(Selection.Text, "#", ChrW(9839))
                Selection.Text = Replace(Selection.Text, "b", ChrW(9837))
            Else ' Output ASCII
                Selection.Text = Replace(Selection.Text, ChrW(9839), "#")
                Selection.Text = Replace(Selection.Text, ChrW(9837), "b")
            End If
            Selection.Collapse Direction:=wdCollapseEnd
            ChordCount = ChordCount + 1
        Loop
    End With
    Application.ScreenUpdating = True
    UnicodeChords = ChordCount - 1
End Function

Sub TestAccidentalRatio()
Select Case Abs(RegexAccidentalRatio)
    Case Is > 1
        MsgBox "Sharp"
    Case Is < 1
        MsgBox "Flat"
    Case 1
        MsgBox "Neither"
End Select
End Sub

Public Function RegexAccidentalRatio() As Single
    Dim objRegEx As Object
    Dim strText As String
    Dim i As Long
    Dim RegExChords() As String
    
    Dim SharpCount As Integer
    Dim FlatCount As Integer
    Dim blnUnicode As Boolean
    
    SharpCount = 0
    FlatCount = 0
    blnUnicode = False
    
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
        .Pattern = "([A-G][b#\u266F\u266D]?(?=(\s(?![a-zH-Z])|\r|\n)|(?=(2|5|6|7|9|11|13|6\/9|7\-5|7\-9|7\#5|7\#9|7??\+5|7\+9|7b5|7b9|7sus2|7sus4|add2|add4|add9|aug|dim|dim7|m\|maj7|m6|m7|m7b5|m9|m1??1|m13|maj7|maj9|maj11|maj13|mb5|m|sus|sus2|sus4|\))(?=(\s|\/)))|(?=(\/|\.|-|\(|\)))))"
    End With
    
    ' Execute RegEx and put all chords in array
    With objRegEx.Execute(strText)
        If .Count = 0 Then Exit Function
        'ReDim RegExChords(.Count - 1)
        For i = 0 To .Count - 1
            Select Case Mid(.Item(i).Value, 2, 1)
                Case "#"
                    SharpCount = SharpCount + 1
                Case "b"
                    FlatCount = FlatCount + 1
                Case ChrW(9839) 'Unicode sharp
                    SharpCount = SharpCount + 1
                    blnUnicode = True
                Case ChrW(9837) 'Unicode flat
                    FlatCount = FlatCount + 1
                    blnUnicode = True
            End Select
        Next i
    End With
    
    'Prevent undefined and indeterminant values
    If SharpCount = 0 Then SharpCount = 1
    If FlatCount = 0 Then FlatCount = 1
    
    If blnUnicode = True Then
        RegexAccidentalRatio = -SharpCount / FlatCount
    Else
        RegexAccidentalRatio = SharpCount / FlatCount
    End If
End Function

Function AccidentalRatio() As Single
Dim SharpCount As Integer
Dim FlatCount As Integer
Dim blnUnicode As Boolean

SharpCount = 0
FlatCount = 0
blnUnicode = False

Application.ScreenUpdating = False

'Go to the first line in the document
Selection.GoTo What:=wdGoToLine, Count:=1

With Selection.Find
    .ClearFormatting 'Reset formatting
    
    'Only search for text with a specific font
    With .Font
        .Subscript = False
        .Superscript = False
    End With
    
    'Find arguments/parameters
    .Text = "[A-G]" 'Searches any capitalized letter from A to G; uses wildcards _
                    https://wordmvp.com/FAQs/General/UsingWildcards.htm
    .Wrap = wdFindContinue 'Wrap: Controls what happens if the search begins _
                            in the middle and the end of the document is reached
    .Forward = True 'Forward: Search toward the end of the document
    .Format = True
    .MatchWildcards = True 'Uses wildcards to search see above link
    .Replacement.Text = ""
    
    'Perform find function and loop through each find
    Do While .Execute
        Select Case Selection.Next(wdCharacter, 1)
            Case "#" 'Sharp
                SharpCount = SharpCount + 1
            Case "b" 'Flat
                FlatCount = FlatCount + 1
            Case ChrW(9839) 'Unicode sharp
                SharpCount = SharpCount + 1
                blnUnicode = True
            Case ChrW(9837) 'Unicode flat
                FlatCount = FlatCount + 1
                blnUnicode = True
            Case Else ' Natural
                'Do nothing
        End Select
        Selection.Collapse Direction:=wdCollapseEnd
    Loop
End With

Application.ScreenUpdating = True

'Prevent undefined and indeterminant values
If SharpCount = 0 Then SharpCount = 1
If FlatCount = 0 Then FlatCount = 1

If blnUnicode = True Then
    AccidentalRatio = -SharpCount / FlatCount
Else
    AccidentalRatio = SharpCount / FlatCount
End If
End Function

