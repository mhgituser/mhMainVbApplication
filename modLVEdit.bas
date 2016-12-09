Attribute VB_Name = "modLVEdit"
Option Explicit

'*********************************************************************************
'*                        EDITABLE LISTVIEW DEMO
'*                    by: Andre Aylestock  (Nov 2002)
'*            (not perfect still :) but improved on it May 2003)
'*********************************************************************************
'*
'*                     You can use this code freely but
'*                      if you improve on it however,
'*            I would really appreciate if you could let me know :)
'*
'*                   e-mail bakerywizard@globetrotter.net
'*
'*         MAKE SURE to insert an icon to your ImageList1 first
'*          before running this demo or you will get an error.
'*         Any icons such as a 16 x 16 or 32 x 32 icon will do.
'*
'*          You can add an icon to imageList1 by going to the
'*           ImageList1's properties and clicking on custom.
'*
'*********************************************************************************

Public gblnLVHasScrollBar   As Boolean
Public glngColIndex         As Long
Public gsngColX             As Single

Public Sub PopulateAndPositionEditControl(ByVal pobjEditControl As Control, _
                                          ByVal pobjLV As ListView, _
                                          ByVal plngItemIndex As Long, _
                                          ByVal plngSubItemIndex As Long, _
                                          ByVal pstrCellText As String, _
                                          Optional plngItemDataValue As Long = 0)
 
    Dim objCH   As ColumnHeader
    
    Set objCH = pobjLV.ColumnHeaders(plngSubItemIndex + 1)
 
    'NOTE: ± 30 values below are only needed as fine adjustements to the pobjEditControl position
        
    SetEditControlText pobjEditControl, _
                       pstrCellText, _
                       plngItemDataValue
      
    'If (column's left edge cannot be seen) Then trim off left portion of pobjEditControl
    If -objCH.Left > pobjLV.ListItems(plngItemIndex).Left Then
        pobjEditControl.Left = pobjLV.Left + 30
        pobjEditControl.Width = objCH.Left + pobjLV.ListItems(plngItemIndex).Left + objCH.Width - 30
    Else
       'If (column's right edge cannot be seen) then trim off right portion of pobjEditControl
        If -objCH.Left - objCH.Width < pobjLV.ListItems(plngItemIndex).Left - pobjLV.Width Then
            pobjEditControl.Left = objCH.Left + 30 + pobjLV.ListItems(plngItemIndex).Left + pobjLV.Left
            'determines if there is a vertical scrollbar or not
            If IsVertScrollbarVisible(pobjLV) Then
                pobjEditControl.Width = objCH.Width _
                                      - ((pobjLV.ListItems(plngItemIndex).Left - pobjLV.Width) + objCH.Left + objCH.Width) _
                                      - 310 '310 is aproximate width of vertical scrollbar
            Else
                pobjEditControl.Width = objCH.Width _
                                      - ((pobjLV.ListItems(plngItemIndex).Left - pobjLV.Width) + objCH.Left + objCH.Width) _
                                      - 50
            End If
        Else
            'if both column's edges can be seen
            pobjEditControl.Left = objCH.Left + 30 + pobjLV.ListItems(plngItemIndex).Left + pobjLV.Left
            pobjEditControl.Width = objCH.Width - 30
        End If
    End If

    'adjust pobjEditControl height.
      
    'play around with the height & top values to see what comes out best. IE: "- 10" & "+ 53"
    ' Note: The height of a combobox cannot be adjusted, hence the textbox test ...
    If TypeOf pobjEditControl Is TextBox Then
        pobjEditControl.Height = pobjLV.ListItems(plngItemIndex).Height - 10
    End If

    pobjEditControl.Top = pobjLV.ListItems(plngItemIndex).Top + pobjLV.Top + 53
      
    pobjEditControl.Visible = True
    pobjEditControl.SetFocus
      
End Sub


Private Sub SetEditControlText(ByVal pobjEditControl As Control, _
                               pstrLVText As String, _
                               Optional plngItemDataValue As Long = 0)

    Dim lngX    As Long

    If TypeOf pobjEditControl Is TextBox Then
        pobjEditControl.Text = pstrLVText
    ElseIf TypeOf pobjEditControl Is DTPicker Then
        If IsDate(pstrLVText) Then
            pobjEditControl.Value = CDate(pstrLVText)
        Else
            pobjEditControl.Value = Date
        End If
    Else
        ' it's a combo box
        If pobjEditControl.ListIndex <> -1 Then
            ' if the combo box already has an item selected, see if it matches
            ' the current listview cell, to avoid resetting ...
            If plngItemDataValue = 0 Then
                If Trim$(pobjEditControl.List(pobjEditControl.ListIndex)) = Trim$(pstrLVText) Then
                    Exit Sub
                End If
            Else
                If pobjEditControl.ItemData(pobjEditControl.ListIndex) = plngItemDataValue Then
                    Exit Sub
                End If
            End If
        End If
        If plngItemDataValue = 0 Then
            For lngX = 0 To pobjEditControl.ListCount - 1
                If Trim$(pobjEditControl.List(lngX)) = Trim$(pstrLVText) Then
                    pobjEditControl.ListIndex = lngX
                    Exit For
                End If
            Next
        Else
            For lngX = 0 To pobjEditControl.ListCount - 1
                If pobjEditControl.ItemData(lngX) = plngItemDataValue Then
                    pobjEditControl.ListIndex = lngX
                    Exit For
                End If
            Next
        End If
    End If

End Sub

Public Sub TransferText(ByVal pobjEditControl As Control, _
                        ByVal pobjLV As ListView, _
                        ByVal plngItemIndex As Long, _
                        ByVal plngSubItemIndex As Long, _
                        Optional plngItemDataCol As Long = 0)

    'sub to transfer/copy the text from the pobjEditControl to the selected ListView cell.
    If pobjEditControl.Visible = False Then Exit Sub
    
    With pobjLV.ListItems(plngItemIndex)
    
        Select Case plngSubItemIndex
            Case 0
                If TypeOf pobjEditControl Is DTPicker Then
                    .Text = Format$(pobjEditControl.Value, "m/d/yyyy")
                Else
                    .Text = Trim$(pobjEditControl.Text)
                End If
            Case Else
                If TypeOf pobjEditControl Is DTPicker Then
                    .SubItems(plngSubItemIndex) = Format$(pobjEditControl.Value, "m/d/yyyy")
                Else
                    .SubItems(plngSubItemIndex) = Trim$(pobjEditControl.Text)
                End If
        End Select
        
        If plngItemDataCol <> 0 Then
            If TypeOf pobjEditControl Is ComboBox Then
                .SubItems(plngItemDataCol) = pobjEditControl.ItemData(pobjEditControl.ListIndex)
            End If
        End If
        
    End With
    
    pobjEditControl.Visible = False

End Sub

Public Function IsVertScrollbarVisible(pobjLV As ListView) As Boolean

    'function to determine if the vertical scrollbar is visible or not.
    'call this function whenever you need to. IE: after adding items to
    'your ListView or before calling the postEditBox Sub
  
    Dim sngListHeight As Single
    Dim i As Long
    
    For i = 1 To pobjLV.ListItems.Count
        sngListHeight = sngListHeight + pobjLV.ListItems(i).Height
    Next i
    
    sngListHeight = sngListHeight + (615)
   
    If pobjLV.Height < sngListHeight Then
        IsVertScrollbarVisible = True
    Else
        IsVertScrollbarVisible = False
    End If
    
End Function

Public Function GetColHdr(psngColX As Single, pobjLV As ListView) As Long

  Dim objCH As ColumnHeader
   
    For Each objCH In pobjLV.ColumnHeaders
        'The next line is needed to determine on which column the user has clicked on
         If psngColX > (objCH.Left + (pobjLV.SelectedItem.Left)) _
         And psngColX < (objCH.Left + (pobjLV.SelectedItem.Left) + objCH.Width) Then
            'Found the column
             GetColHdr = objCH.Index
            Exit For
         End If
    Next

End Function
