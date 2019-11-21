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
        .Text = "\|?*\|"
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
Select Case Abs(AccidentalRatio)
    Case Is > 1
        MsgBox "Sharp"
    Case Is < 1
        MsgBox "Flat"
    Case 1
        MsgBox "Neither"
End Select
End Sub

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

