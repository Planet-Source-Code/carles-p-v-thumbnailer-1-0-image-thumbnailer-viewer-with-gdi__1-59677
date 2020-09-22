Attribute VB_Name = "mDialogColor"
'================================================
' Module:        mDialogColor.bas
' Author:
' Dependencies:  None
' Last revision: 2004.06.14
'================================================

Option Explicit

Private Type CHOOSECOLOR
    lStructSize    As Long
    hwndOwner      As Long
    hInstance      As Long
    rgbResult      As Long
    lpCustColors   As Long
    Flags          As Long
    lCustData      As Long
    lpfnHook       As Long
    lpTemplateName As String
End Type

Private Const CC_RGBINIT   As Long = &H1
Private Const CC_FULLOPEN  As Long = &H2
Private Const CC_ANYCOLOR  As Long = &H100

Private Const CC_NORMAL    As Long = CC_ANYCOLOR Or CC_RGBINIT
Private Const CC_EXTENDED  As Long = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN

Private Declare Function CHOOSECOLOR Lib "comdlg32" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Private m_lCustomColors(15) As Long
Private m_bInitialized      As Boolean



Public Function SelectColor(ByVal hWndParent As Long, ByVal DefaultColor As Long, Optional ByVal Extended As Boolean = 0) As Long
 
  Dim uCC  As CHOOSECOLOR
  Dim lRet As Long
  Dim lIdx As Long
 
    With uCC
        
        '-- Initiliaze custom colors
        If (m_bInitialized = False) Then
            m_bInitialized = True
            
            For lIdx = 0 To 15
                m_lCustomColors(lIdx) = QBColor(lIdx)
            Next lIdx
        End If
        
        '-- Prepare struct.
        .lStructSize = Len(uCC)
        .hwndOwner = hWndParent
        .rgbResult = DefaultColor
        .lpCustColors = VarPtr(m_lCustomColors(0))
        .Flags = IIf(Extended, CC_EXTENDED, CC_NORMAL)
        
        '-- Show Color dialog
        lRet = CHOOSECOLOR(uCC)
         
        '-- Get color / Cancel
        If (lRet) Then
            SelectColor = .rgbResult
          Else
            SelectColor = -1
        End If
    End With
End Function
