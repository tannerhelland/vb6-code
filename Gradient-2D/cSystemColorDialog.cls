VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSystemColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Display a system color selection dialog
' Adopted from a reference implementation at http://www.devx.com/vb2themax/Tip/19257, originally by Marco Bellinaso
' See also the MSDN documentation for ChooseColor at https://msdn.microsoft.com/en-us/library/windows/desktop/ms646912%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396

Option Explicit

Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpChooseColor As ChooseColorStruct) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
    
Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_PREVENTFULLOPEN = &H4&
Private Const CC_SHOWHELP = &H8&
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_ENABLETEMPLATEHANDLE = &H40&
Private Const CC_SOLIDCOLOR = &H80&
Private Const CC_ANYCOLOR = &H100&
Private Const CLR_INVALID = &HFFFF

' Show the common dialog for choosing a color.
' Returns TRUE if OK was selected; FALSE otherwise
'
' dstColor holds the return value; this value is nonsense if the function returns FALSE
' hParent is the handle (hWnd) of the parent form
' showFullDialog specifies whether the dialog will show both selection panels by default
' initColor is the color initially selected when the dialog is open

' Example:
'    Dim oleNewColor As OLE_COLOR
'    If ShowColorsDialog(oleNewColor, Me.hwnd, True, vbRed) Then Me.BackColor = oleNewColor
Function ShowColorDialog(ByRef dstColor As Long, Optional ByVal hParent As Long = 0&, Optional ByVal showFullDialog As Boolean, Optional ByVal initColor As OLE_COLOR) As Boolean
    
    Dim cc As ChooseColorStruct
    Dim aColorRef(15) As Long
    Dim lInitColor As Long
  
    ' Translate the initial OLE color to a long value
    If (initColor <> 0) Then
        If OleTranslateColor(initColor, 0, lInitColor) Then lInitColor = CLR_INVALID
    End If
    
    ' Fill the ChooseColorStruct struct
    With cc
        .lStructSize = Len(cc)
        .hwndOwner = hParent
        .lpCustColors = VarPtr(aColorRef(0))
        .rgbResult = initColor
        .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(showFullDialog, CC_FULLOPEN, 0)
    End With
    
    ' Show the dialog and return the result (if any)
    ShowColorDialog = CBool(ChooseColor(cc))
    If ShowColorDialog Then dstColor = cc.rgbResult
    
End Function

