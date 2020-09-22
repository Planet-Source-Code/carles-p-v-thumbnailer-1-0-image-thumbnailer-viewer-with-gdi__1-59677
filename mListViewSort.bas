Attribute VB_Name = "mListViewSort"
'================================================
' Module:        mListViewEx.bas
' Last revision: 2004.11.19
'================================================

Option Explicit
Option Compare Text

Private Type LVITEM_lp
    mask       As Long
    iItem      As Long
    iSubItem   As Long
    State      As Long
    stateMask  As Long
    pszText    As Long
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
End Type
 
Private Const MAX_PATH        As Long = 260
Private Const LVIF_TEXT       As Long = &H1
Private Const LVM_FIRST       As Long = &H1000
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const LVM_SORTITEMSEX As Long = (LVM_FIRST + 81)
     
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'//

Private m_uLVI            As LVITEM_lp
Private m_a(MAX_PATH - 1) As Byte
Private m_lLen            As Long
Private m_PRECEDE         As Long
Private m_FOLLOW          As Long

Private Function pvCompareIndex( _
                 ByVal lParam1 As Long, _
                 ByVal lParam2 As Long, _
                 ByVal hWnd As Long _
                 ) As Long

    If (lParam1 > lParam2) Then
        pvCompareIndex = m_PRECEDE
    ElseIf (lParam1 < lParam2) Then
        pvCompareIndex = m_FOLLOW
    End If
End Function

Private Function pvCompareText( _
                 ByVal lParam1 As Long, _
                 ByVal lParam2 As Long, _
                 ByVal hWnd As Long _
                 ) As Long

  Dim val1 As String
  Dim val2 As String
     
    val1 = pvGetItemText(hWnd, lParam1)
    val2 = pvGetItemText(hWnd, lParam2)
     
    If (val1 > val2) Then
        pvCompareText = m_PRECEDE
    ElseIf (val1 < val2) Then
        pvCompareText = m_FOLLOW
    End If
End Function

Private Function pvCompareValue( _
                 ByVal lParam1 As Long, _
                 ByVal lParam2 As Long, _
                 ByVal hWnd As Long _
                 ) As Long

  Dim val1 As Double
  Dim val2 As Double
     
    val1 = CDbl(pvGetItemText(hWnd, lParam1))
    val2 = CDbl(pvGetItemText(hWnd, lParam2))
     
    If (val1 > val2) Then
        pvCompareValue = m_PRECEDE
    ElseIf (val1 < val2) Then
        pvCompareValue = m_FOLLOW
    End If
End Function

Private Function pvCompareDate( _
                 ByVal lParam1 As Long, _
                 ByVal lParam2 As Long, _
                 ByVal hWnd As Long _
                 ) As Long

  Dim val1 As Date
  Dim val2 As Date
     
    val1 = CDate(pvGetItemText(hWnd, lParam1))
    val2 = CDate(pvGetItemText(hWnd, lParam2))
     
    If (val1 > val2) Then
        pvCompareDate = m_PRECEDE
    ElseIf (val1 < val2) Then
        pvCompareDate = m_FOLLOW
    End If
End Function

'//

Private Function pvGetItemText( _
                 ByVal hWnd As Long, _
                 ByVal lParam As Long _
                 ) As String

    m_lLen = SendMessage(hWnd, LVM_GETITEMTEXT, lParam, m_uLVI)
    pvGetItemText = Left$(StrConv(m_a(), vbUnicode), m_lLen)
End Function

Private Function AddressOfFunction(lpfn As Long) As Long
    AddressOfFunction = lpfn
End Function

'//

Public Function Sort( _
                ByVal hListView As Long, _
                ByVal Column As Long, _
                ByVal SortOrder As tvSortOrderConstants, _
                ByVal SortType As tvSortTypeConstants _
                ) As Boolean

  Dim lRet As Long
  
    With m_uLVI
        .mask = LVIF_TEXT
        .pszText = VarPtr(m_a(0))
        .cchTextMax = MAX_PATH
        .iSubItem = Column
    End With
        
    Select Case SortOrder
        
        Case [tvDefault]
            
            m_PRECEDE = 1
            m_FOLLOW = -1
            lRet = SendMessageLong(hListView, LVM_SORTITEMSEX, hListView, AddressOfFunction(AddressOf pvCompareIndex))
            
        Case [tvAscending], [tvDescending]
        
            m_PRECEDE = SortOrder
            m_FOLLOW = -SortOrder
            
            Select Case SortType
                Case [tvString]
                    lRet = SendMessageLong(hListView, LVM_SORTITEMSEX, hListView, AddressOfFunction(AddressOf pvCompareText))
                Case [tvNumeric]
                    lRet = SendMessageLong(hListView, LVM_SORTITEMSEX, hListView, AddressOfFunction(AddressOf pvCompareValue))
                Case [tvDate]
                    lRet = SendMessageLong(hListView, LVM_SORTITEMSEX, hListView, AddressOfFunction(AddressOf pvCompareDate))
            End Select
    End Select
    
    Sort = CBool(lRet)
End Function
