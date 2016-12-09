Attribute VB_Name = "modListView"
Option Explicit

Private Const LVM_FIRST             As Long = &H1000
Private Const LVM_SCROLL            As Long = (LVM_FIRST + 20)
Private Const LVM_GETCOUNTPERPAGE   As Long = (LVM_FIRST + 40)
Private Const LVM_HITTEST           As Long = (LVM_FIRST + 18)
Private Const LVM_SUBITEMHITTEST    As Long = (LVM_FIRST + 57)
Private Const LVHT_ONITEMICON       As Long = &H2
Private Const LVHT_ONITEMLABEL      As Long = &H4
Private Const LVHT_ONITEMSTATEICON  As Long = &H8
Private Const LVHT_ONITEM           As Long = (LVHT_ONITEMICON Or _
                                               LVHT_ONITEMLABEL Or _
                                               LVHT_ONITEMSTATEICON)
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type LVHITTESTINFO
   pt As POINTAPI
   Flags As Long
   iItem As Long
   iSubItem  As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub GetLVCellData(pobjLV As ListView, _
                         psngX As Single, _
                         psngY As Single, _
                         ByRef pstrCellText As String, _
                         ByRef plngItemIndex As Long, _
                         ByRef plngSubItemIndex As Long)
    
    Dim HTI As LVHITTESTINFO
    Dim lst As ListItem

    With HTI
        .pt.X = (psngX \ Screen.TwipsPerPixelX)
        .pt.Y = (psngY \ Screen.TwipsPerPixelY)
        .Flags = LVHT_ONITEM
    End With
      
    SendMessage pobjLV.hwnd, LVM_SUBITEMHITTEST, 0, HTI

    If (HTI.iItem > -1) Then

        Set lst = pobjLV.ListItems(HTI.iItem + 1)

        plngItemIndex = HTI.iItem + 1
        plngSubItemIndex = HTI.iSubItem
        
        If HTI.iSubItem = 0 Then
            pstrCellText = pobjLV.ListItems(HTI.iItem + 1).Text
        Else
            pstrCellText = pobjLV.ListItems(HTI.iItem + 1).SubItems(HTI.iSubItem)
        End If
    Else
        pstrCellText = ""
        plngItemIndex = 0
        plngSubItemIndex = 0
    End If

End Sub
