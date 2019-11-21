Attribute VB_Name = "modColor"
Option Explicit

Private Type CHOOSECOLOR
  lStructSize As Long
  hwndOwner As LongPtr
  hInstance As LongPtr
  rgbResult As Long
  lpCustColors As LongPtr
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Declare PtrSafe Function MyChooseColor _
    Lib "comdlg32.dll" Alias "ChooseColorW" _
    (ByRef pChoosecolor As CHOOSECOLOR) As Boolean

Private Declare PtrSafe Function VarPtrArray _
  Lib "VBE7" Alias _
  "VarPtr" (ByRef Var() As Any) As LongPtr

Sub FontColorTest()
  Dim col As Long
  col = RGB(200, 100, 50)
  GetColor col
  Debug.Print col
End Sub

Public Function GetColor(ByRef col As Long) As _
    Boolean
    
    Static CS As CHOOSECOLOR
    Static CustColor(15) As Long
    
    With CS
        .lStructSize = Len(CS)
        .hwndOwner = 0
        .flags = &H1 Or &H2
        .lpCustColors = VarPtr(CustColor(0))
        .rgbResult = col
        .hInstance = 0
    End With

    GetColor = MyChooseColor(CS)
    If GetColor = False Then Exit Function

    GetColor = True
    col = CS.rgbResult
End Function

