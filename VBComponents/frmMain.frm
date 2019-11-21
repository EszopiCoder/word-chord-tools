VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Chord Tools"
   ClientHeight    =   4464
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isRunning As Boolean


Private Sub UserForm_Initialize()
'Turn on isRunning (prevent option button events from firing)
isRunning = True

'Set values for help text
lblColor.ControlTipText = "Click to change color"

'Set values for RGB
lblColor.Caption = ""
lblColor.BackColor = 0

'Set values for Bold, Italics, Underline
tglBold.ControlTipText = "Unbold chords"
tglItalic.ControlTipText = "Unitalicize chords"
tglUnderline.ControlTipText = "Remove underline"
tglUnderline.Enabled = False

'Set values for underline combobox
cboUnderline.AddItem "Underline"
cboUnderline.AddItem "Double underline"
cboUnderline.AddItem "Thick underline"
cboUnderline.AddItem "Dotted underline"
cboUnderline.AddItem "Dashed underline"
cboUnderline.AddItem "Dot-dash underline"
cboUnderline.AddItem "Dot-dot-dash underline"
cboUnderline.AddItem "Wave underline"
cboUnderline.Enabled = False

'Set values for encoding combobox
cboEncoding.AddItem "ASCII", 0
cboEncoding.AddItem "Unicode", 1

'Set values for chord options
Dim SFRatio As Single
SFRatio = AccidentalRatio
Select Case Abs(SFRatio)
    Case Is > 1
        optSharp.Value = True
    Case Is < 1
        optFlat.Value = True
    Case Else
        MsgBox "Cannot detect chord accidental." & vbNewLine & _
            "Choose chord accidental manually.", vbExclamation, Me.Caption
End Select
If SFRatio < 0 Then
    cboEncoding.ListIndex = 1
Else
    cboEncoding.ListIndex = 0
End If

'Turn off isRunning
isRunning = False
End Sub
Private Sub lblColor_Click()
    Dim col As Long
    col = lblColor.BackColor
    GetColor col
    lblColor.BackColor = Replace(col, "-", "")
End Sub
Private Sub tglBold_Click()
If tglBold.Value = True Then
    tglBold.ControlTipText = "Bold chords"
Else
    tglBold.ControlTipText = "Unbold chords"
End If
End Sub
Private Sub tglItalic_Click()
If tglItalic.Value = True Then
    tglItalic.ControlTipText = "Italicize chords"
Else
    tglItalic.ControlTipText = "Unitalicize chords"
End If
End Sub
Private Sub tglUnderline_Click()
If tglUnderline.Value = True Then
    tglItalic.ControlTipText = "Underline chords"
Else
    tglItalic.ControlTipText = "Remove underline"
End If
End Sub
Private Sub cmdApply_Click()
Dim UnderlineConstant As Integer

'Determine underline constant
If tglUnderline.Value = True Then
    Select Case cboUnderline.ListIndex
        Case 0, -1
            UnderlineConstant = wdUnderlineSingle
        Case 1
            UnderlineConstant = wdUnderlineDouble
        Case 2
            UnderlineConstant = wdUnderlineThick
        Case 3
            UnderlineConstant = wdUnderlineDotted
        Case 4
            UnderlineConstant = wdUnderlineDash
        Case 5
            UnderlineConstant = wdUnderlineDotDash
        Case 6
            UnderlineConstant = wdUnderlineDotDotDash
        Case 7
            UnderlineConstant = wdUnderlineWavy
    End Select
Else
    UnderlineConstant = 0 'No underline
End If

'Change the font of the chords based on user input
    Call modFont.ChordMarkerDoc
    MsgBox modFont.FormatChords(CLng(Replace(lblColor.BackColor, "-", "")), _
    tglBold.Value, tglItalic.Value) & " chords detected.", vbInformation, Me.Caption

End Sub
Private Sub optSharp_Click()
Dim intRtrn As Integer
Dim blnUnicode As Boolean

'Assume flat is the chosen option

'Check if running Userform_Initialize
If isRunning = True Then Exit Sub

'Determine if unicode is chosen
If cboEncoding.ListIndex = 1 Then
    blnUnicode = True
Else
    blnUnicode = False
End If

If MsgBox("Would you like to change the accidentals to sharps?", vbYesNo, Me.Caption) = vbYes Then _
    Call TransposeDoc(True, 0, blnUnicode)
End Sub
Private Sub optFlat_Click()
Dim intRtrn As Integer
Dim blnUnicode As Boolean

'Assume sharp is the chosen option

'Check if running Userform_Initialize
If isRunning = True Then Exit Sub

'Determine if unicode is chosen
If cboEncoding.ListIndex = 1 Then
    blnUnicode = True
Else
    blnUnicode = False
End If

If MsgBox("Would you like to change the accidentals to flats?", vbYesNo, Me.Caption) = vbYes Then _
    Call TransposeDoc(False, 0, blnUnicode)
End Sub
Private Sub cboEncoding_Change()
Dim intRtrn As Integer

'Check if running Userform_Initialize
If isRunning = True Then Exit Sub

Select Case cboEncoding.ListIndex
    Case 0 'ASCII
        If MsgBox("Would you like to change the encoding to ASCII?", vbYesNo, Me.Caption) = vbYes Then
            Call ChordMarkerDoc
            Call UnicodeChords(False)
        End If
    Case 1 'Unicode
        If MsgBox("Would you like to change the encoding to Unicode?", vbYesNo, Me.Caption) = vbYes Then
            Call ChordMarkerDoc
            Call UnicodeChords(True)
        End If
End Select
End Sub
Private Sub TransposeSpin_SpinUp()
Dim blnUnicode As Boolean

'Detect if chord accidental is chosen
If optSharp.Value = False And optFlat.Value = False Then
    MsgBox "Choose a chord accidental", vbExclamation, Me.Caption
    Exit Sub
End If
'Determine if unicode is chosen
If cboEncoding.ListIndex = 1 Then
    blnUnicode = True
Else
    blnUnicode = False
End If
'Transpose the chords in the document
Call modTranspose.TransposeDoc(optSharp.Value, 1, blnUnicode)
'Show semitone change to user
lblChange.Caption = Val(lblChange.Caption) + 1
'Since 12 semitones is equal to an octave
If Abs(Val(lblChange.Caption)) = 12 Then lblChange.Caption = 0
'Change the label to reflect the proper pluralization of numbers
If Abs(Val(lblChange.Caption)) = 1 Then
    lblSemitone.Caption = " Semitone"
Else
    lblSemitone.Caption = " Semitones"
End If
End Sub
Private Sub TransposeSpin_SpinDown()
Dim blnUnicode As Boolean

'Detect if chord accidental is chosen
If optSharp.Value = False And optFlat.Value = False Then
    MsgBox "Choose a chord accidental", vbExclamation, Me.Caption
    Exit Sub
End If
'Determine if unicode is chosen
If cboEncoding.ListIndex = 1 Then
    blnUnicode = True
Else
    blnUnicode = False
End If
'Transpose the chords in the document
Call modTranspose.TransposeDoc(optSharp.Value, -1, blnUnicode)
'Show semitone change to user
lblChange.Caption = Val(lblChange.Caption) - 1
'Since 12 semitones is equal to an octave
If Abs(Val(lblChange.Caption)) = 12 Then lblChange.Caption = 0
'Change the label to reflect the proper pluralization of numbers
If Abs(Val(lblChange.Caption)) = 1 Then
    lblSemitone.Caption = " Semitone"
Else
    lblSemitone.Caption = " Semitones"
End If
End Sub
Private Sub cmdReset_Click()
Dim blnUnicode As Boolean

'Detect if chord accidental is chosen
If optSharp.Value = False And optFlat.Value = False Then
    MsgBox "Choose a chord accidental", vbExclamation, Me.Caption
    Exit Sub
End If
'Determine if unicode is chosen
If cboEncoding.ListIndex = 1 Then
    blnUnicode = True
Else
    blnUnicode = False
End If
'Exit sub if semitone change is 0
If Val(lblChange.Caption) = 0 Then Exit Sub
'Reset the transposition to the original
Call modTranspose.TransposeDoc(optSharp.Value, -Val(lblChange.Caption), blnUnicode)
'Show semitone change to user
lblChange.Caption = 0
'Change the label to reflect the proper pluralization of numbers
lblSemitone.Caption = " Semitones"
End Sub
Private Sub btnHelp_Click()
    Dim URL As String
    URL = "https://github.com/EszopiCoder/word-chord-tools"
    ActiveDocument.FollowHyperlink URL
End Sub
Private Sub btnAbout_Click()
    MsgBox "'Chord Tools' was created by EszopiCoder." & vbNewLine & _
        "Open Source (https://github.com/EszopiCoder/word-chord-tools)" & vbNewLine & _
        "Please report bugs and send suggestions to pharm.coder@gmail.com", vbInformation
End Sub
Private Sub cmdClose_Click()
Me.Hide
End Sub
