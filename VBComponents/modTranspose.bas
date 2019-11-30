Attribute VB_Name = "modTranspose"
' Based on RegEx from https://pastebin.com/994jy3ab
' Note: RegExp and Dictionary objects do not work on Mac OS

Option Explicit
Private Const SharpScale As String = "A A#B C C#D D#E F F#G G#"
Private Const FlatScale As String = "A BbB C DbD EbE F GbG Ab"

Public Sub TestFx()
    MsgBox TransposeDoc(True, -1, False)
End Sub

Public Function TransposeSelection(ByVal Sharp As Boolean, _
    ByVal NumSemitones As Integer, _
    Optional ByVal Unicode As Boolean = False) As Long
    Dim objRegEx As Object
    Dim strText As String
    Dim i As Long
    Dim RegExChords() As String
    Dim SongChords() As Variant
    Dim TransposedChords() As Variant
    
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
        .Pattern = "([A-G][b#\u266F\u266D]?(?=(\s(?![a-zH-Z])|\r|\n)|(?=(2|5|6|7|9|11|13|6\/9|7\-5|7\-9|7\#5|7\#9|7??\+5|7\+9|7b5|7b9|7sus2|7sus4|add2|add4|add9|aug|dim|dim7|m\|maj7|m6|m7|m7b5|m9|m1??1|m13|maj7|maj9|maj11|maj13|mb5|m|sus|sus2|sus4|\))(?=(\s|\/)))|(?=(\/|\.|-|\(|\)))))"
    End With
    
    ' Execute RegEx and put all chords in array
    With objRegEx.Execute(strText)
        If .Count = 0 Then Exit Function
        ReDim RegExChords(.Count - 1)
        For i = 0 To .Count - 1
            RegExChords(i) = .Item(i).Value
        Next i
    End With
    ' Remove duplicate chords from array
    SongChords = RemoveDupesDict(RegExChords)
    ' Format chords in the original string to: |Chord|
    strText = objRegEx.Replace(strText, "|$1|")
    
    ' Transpose chords and save to another array
    ReDim TransposedChords(UBound(SongChords))
    For i = LBound(SongChords) To UBound(SongChords)
        TransposedChords(i) = Transpose(CStr(SongChords(i)), Sharp, NumSemitones, Unicode)
    Next i
    
    ' Replace chords in original string
    For i = LBound(SongChords) To UBound(SongChords)
        With objRegEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = "\|" & SongChords(i) & "\|"
            strText = objRegEx.Replace(strText, TransposedChords(i))
        End With
    Next i
    
    ' Replace text and paste formatting
    Selection.Text = strText
    Selection.PasteFormat
    Selection.Collapse wdCollapseEnd
    
    ' Output transposed chord
    TransposeSelection = UBound(RegExChords) - LBound(RegExChords) + 1
End Function
Public Function TransposeDoc(ByVal Sharp As Boolean, _
    ByVal NumSemitones As Integer, _
    Optional ByVal Unicode As Boolean = False) As Long
    Dim objRegEx As Object
    Dim strText As String
    Dim i As Long
    Dim RegExChords() As String
    Dim SongChords() As Variant
    Dim TransposedChords() As Variant
    
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
        ReDim RegExChords(.Count - 1)
        For i = 0 To .Count - 1
            RegExChords(i) = .Item(i).Value
        Next i
    End With
    ' Remove duplicate chords from array
    SongChords = RemoveDupesDict(RegExChords)
    ' Format chords in the original string to: |Chord|
    ' Note: Chord is only root and sharp/flat (Ex: Bb)
    strText = objRegEx.Replace(strText, "|$1|")
    
    ' Transpose chords and save to another array
    ReDim TransposedChords(UBound(SongChords))
    For i = LBound(SongChords) To UBound(SongChords)
        TransposedChords(i) = Transpose(CStr(SongChords(i)), Sharp, NumSemitones, Unicode)
    Next i
    
    ' Replace chords in original string
    For i = LBound(SongChords) To UBound(SongChords)
        With objRegEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = "\|" & SongChords(i) & "\|"
            strText = objRegEx.Replace(strText, TransposedChords(i))
        End With
    Next i
    
    ' Replace text and paste formatting
    ' Remove excess carriage return at the end of string
    Selection.Text = Left(strText, Len(strText) - 1)
    Selection.PasteFormat
    Selection.Collapse wdCollapseEnd
    
    TransposeDoc = UBound(RegExChords) - LBound(RegExChords) + 1
End Function
Public Function Transpose(Chord As String, ByVal Sharp As Boolean, _
    Optional ByVal NumSemitones As Integer = 1, _
    Optional ByVal Unicode As Boolean = False) As String
    Dim ChordPosition As Integer
    Dim ChromaticScale As String
    
    'Ensure correct chord length to prevent errors
    Select Case Len(Chord)
        Case 0
            MsgBox "ERROR: No chord inputted.", vbExclamation
            Exit Function
        Case 1 'Add a space to avoid searching for wrong chord in constant string
            Chord = Chord & " "
        Case Is > 2
            MsgBox "ERROR: Chord length too long.", vbExclamation
            Exit Function
    End Select
    
    ' Determine if input is sharps/flats
    Select Case Right(Chord, 1)
        Case " "
            ' Will be the same position with either scale
            ChromaticScale = SharpScale
        Case "#"
            ChromaticScale = SharpScale
        Case "b"
            ChromaticScale = FlatScale
        Case ChrW(9839) ' Unicode sharp
            Chord = Replace(Chord, ChrW(9839), "#")
            ChromaticScale = SharpScale
        Case ChrW(9837) ' Unicode flat
            Chord = Replace(Chord, ChrW(9837), "b")
            ChromaticScale = FlatScale
    End Select
    
    'Find position of chord in constant string
    ChordPosition = InStr(1, ChromaticScale, Chord)
    If ChordPosition = 0 Then
    MsgBox "ERROR: Chord not found.", vbExclamation
        Transpose = Chord
        Exit Function
    End If
    
    'Add number of half steps
    ChordPosition = ChordPosition + (2 * NumSemitones)
    
    'Make sure chord is within constant string (pseudo text wrap)
    If ChordPosition > 24 Then ChordPosition = ChordPosition - 24
    If ChordPosition < 1 Then ChordPosition = ChordPosition + 24
    
    'Determine if output is sharps/flats
    If Sharp = True Then
        ChromaticScale = SharpScale
    Else
        ChromaticScale = FlatScale
    End If
    
    'Output chord
    Select Case Unicode
        Case True
            Transpose = Replace(Replace(Trim(Mid(ChromaticScale, ChordPosition, 2)), "#", ChrW(9839)), "b", ChrW(9837))
        Case False
            Transpose = Trim(Mid(ChromaticScale, ChordPosition, 2))
    End Select
End Function
Private Function RemoveDupesDict(MyArray As Variant) As Variant
'DESCRIPTION: Removes duplicates from your array using the dictionary method.
'NOTES: (1.a) You must add a reference to the Microsoft Scripting Runtime library via
' the Tools > References menu.
' (1.b) This is necessary because I use Early Binding in this function.
' Early Binding greatly enhances the speed of the function.
' (2) The scripting dictionary will not work on the Mac OS.
'SOURCE: https://wellsr.com
'-----------------------------------------------------------------------
    Dim i As Long
    Dim objDict As Object
    Set objDict = CreateObject("Scripting.Dictionary")
    With objDict
        For i = LBound(MyArray) To UBound(MyArray)
            If IsMissing(MyArray(i)) = False Then
                .Item(MyArray(i)) = 1
            End If
        Next
        RemoveDupesDict = .Keys
    End With
End Function
