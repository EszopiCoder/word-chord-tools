Attribute VB_Name = "modMenu"
'<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
'   <ribbon>
'      <tabs>
'         <tab id="TabChord" label="Chord Tools">
'            <group id="Transpose" label="Transpose Chords">
'               <button id="TransposeUp"
'                   imageMso = "AnimationMoveEarlier"
'                   label = "Transpose Up"
'                   screentip="Transpose Up"
'                   supertip = "Tranpose 1 semitone up."
'                   onAction = "TransposeUp_Click"
'                   size="large"/>
'               <button id="TransposeDown"
'                   imageMso = "AnimationMoveLater"
'                   label = "Transpose Down"
'                   screentip="Transpose Down"
'                   supertip = "Tranpose 1 semitone down."
'                   onAction = "TransposeDown_Click"
'                   size="large"/>
'            </group>
'            <group id="Accidentals" label="Format Accidentals">
'               <button id="SwitchAccidental"
'                   imageMso = "GridShowHide"
'                   label = "Switch Accidental"
'                   screentip="Switch Accidental"
'                   supertip = "Toggle between sharps and flats."
'                   onAction = "SwitchAccidental_Click"
'                   size="large"/>
'               <button id="SwitchUnicode"
'                   imageMso = "SymbolInsert"
'                   label = "Switch Unicode"
'                   screentip="Switch Unicode"
'                   supertip = "Toggle between ASCII sharps and flats."
'                   onAction = "SwitchUnicode_Click"
'                   size="large"/>
'            </group>
'            <group id="Other" label="Other">
'               <button id="Bold"
'                   imageMso = "Bold"
'                   label = "Bold Chords"
'                   screentip="Bold Chords"
'                   supertip = "Bold all chords."
'                   onAction = "BoldAllChords_Click"
'                   size="large"/>
'               <button id="openUF"
'                   imageMso = "DataFormExcel"
'                   label = "Open Userform"
'                   screentip="Open Userform"
'                   supertip = "Open main userform."
'                   onAction = "OpenUserform_Click"
'                   size="large"/>
'               <button id="getInfo"
'                   imageMso = "ARMPreviewButton"
'                   label = "Info"
'                   screentip="Information"
'                   supertip = "Return contact information."
'                   onAction = "getInfo_Click"
'                   size="large"/>
'            </group>
'         </tab>
'      </tabs>
'   </ribbon>
'</customUI>
'*********************************XML CODE*********************************

Option Explicit
Private Sharp As Boolean
Private Unicode As Boolean

Sub TransposeUp_Click(control As IRibbonControl)
    
    Call GetDocFormat
    Call TransposeDoc(Sharp, 1, Unicode)
    
End Sub

Sub TransposeDown_Click(control As IRibbonControl)
    
    Call GetDocFormat
    Call TransposeDoc(Sharp, -1, Unicode)
    
End Sub

Sub SwitchAccidental_Click(control As IRibbonControl)
    
    Call GetDocFormat
    Call TransposeDoc(Not Sharp, 0, Unicode)
    
End Sub

Sub SwitchUnicode_Click(control As IRibbonControl)

    Call GetDocFormat
    Call ChordMarkerDoc
    Call UnicodeChords(Not Unicode)

End Sub

Sub BoldAllChords_Click(control As IRibbonControl)

    Call ChordMarkerDoc
    Call BoldChords(True)

End Sub

Sub OpenUserForm_Click(control As IRibbonControl)
    
    frmMain.Show
    
End Sub

Sub getInfo_Click(control As IRibbonControl)

    MsgBox "'Chord Tools' was created by EszopiCoder." & vbNewLine & _
        "Open Source (https://github.com/EszopiCoder/word-chord-tools)" & vbNewLine & _
        "Please report bugs and send suggestions to pharm.coder@gmail.com", vbInformation
    
End Sub

Private Sub GetDocFormat()
    'Set values for chord options
    Dim SFRatio As Single
    SFRatio = AccidentalRatio
    Select Case Abs(SFRatio)
        Case Is > 1
            Sharp = True
        Case Is < 1
            Sharp = False
        Case Else
            MsgBox "Cannot detect chord accidental.", vbExclamation
            Exit Sub
    End Select
    If SFRatio < 0 Then
        Unicode = True
    Else
        Unicode = False
    End If
End Sub
