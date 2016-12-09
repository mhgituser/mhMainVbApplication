VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form frmduplicatecheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duplicate Check"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16095
   Icon            =   "frmduplicatecheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   16095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Distribution List"
      Height          =   3255
      Left            =   12120
      TabIndex        =   13
      Top             =   3600
      Width           =   3855
      Begin VSFlex7Ctl.VSFlexGrid distgrid 
         Height          =   2895
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3615
         _cx             =   6376
         _cy             =   5106
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmduplicatecheck.frx":0E42
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   600
      TabIndex        =   9
      Top             =   0
      Width           =   5175
      Begin VB.OptionButton optdropouts 
         Caption         =   "Expected Drop outs"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optactive 
         Caption         =   "Active"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optduplicate 
         Caption         =   "Duplicate"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      Caption         =   "Planted List"
      Height          =   3255
      Left            =   8280
      TabIndex        =   3
      Top             =   3600
      Width           =   3855
      Begin VSFlex7Ctl.VSFlexGrid pgrid 
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3615
         _cx             =   6376
         _cy             =   5106
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmduplicatecheck.frx":0EA6
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Field"
      Height          =   3255
      Left            =   4440
      TabIndex        =   2
      Top             =   3600
      Width           =   3855
      Begin VSFlex7Ctl.VSFlexGrid fgrid 
         Height          =   2895
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3615
         _cx             =   6376
         _cy             =   5106
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmduplicatecheck.frx":0F0A
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Storage"
      Height          =   3255
      Left            =   600
      TabIndex        =   1
      Top             =   3600
      Width           =   3855
      Begin VSFlex7Ctl.VSFlexGrid sgrid 
         Height          =   2895
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3615
         _cx             =   6376
         _cy             =   5106
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmduplicatecheck.frx":0F6E
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Duplicate Farmers"
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   9735
      Begin VSFlex7Ctl.VSFlexGrid dgrid 
         Height          =   2535
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9495
         _cx             =   16748
         _cy             =   4471
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmduplicatecheck.frx":0FD2
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
End
Attribute VB_Name = "frmduplicatecheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MTYPE As String
Private Sub chka_Click()
If chka.Value = 1 Then
fillastorage
fillafield
fillaplanted
Else
filldstorage
filldfield
filldplanted
End If
End Sub

Private Sub Command1_Click()
Exit Sub
Dim i As Integer
If optactive.Value = True Then
MsgBox "You cannot update the Active Farmers."
Exit Sub
End If

Select Case MTYPE
    Case "D"
    'registration
   
     For i = 1 To dgrid.Rows - 1
    If Len(dgrid.TextMatrix(i, 1)) = 0 Then Exit For
    
    If Len(Trim(dgrid.TextMatrix(i, 2))) > 14 Then
    MHVDB.Execute "update tblfarmer set status='D',monitor='',remarks='X' where idfarmer='" & Mid(Trim(dgrid.TextMatrix(i, 1)), 1, 14) & "'"
    MHVDB.Execute "update tbllandreg set status='D' where farmerid='" & Mid(Trim(dgrid.TextMatrix(i, 1)), 1, 14) & "'"

    Else
    MHVDB.Execute "update tblfarmer set idfarmer='" & Mid(Trim(dgrid.TextMatrix(i, 2)), 1, 14) & "' where idfarmer='" & Mid(Trim(dgrid.TextMatrix(i, 1)), 1, 14) & "'"
    MHVDB.Execute "update tbllandreg set farmerid='" & Mid(Trim(dgrid.TextMatrix(i, 2)), 1, 14) & "' where farmerid='" & Mid(Trim(dgrid.TextMatrix(i, 1)), 1, 14) & "'"
    End If
       
    
    
    Next
    
    
    
    
    'storage
    For i = 1 To dgrid.Rows - 1
    If Len(dgrid.TextMatrix(i, 1)) = 0 Then Exit For
    ODKDB.Execute "update storagehub6_core set farmerbarcode='" & Mid(Trim(dgrid.TextMatrix(i, 2)), 1, 14) & "' where farmerbarcode='" & Mid(Trim(dgrid.TextMatrix(i, 1)), 1, 14) & "'"
    Next
    
    'field
    For i = 1 To dgrid.Rows - 1
    If Len(dgrid.TextMatrix(i, 1)) = 0 Then Exit For
    ODKDB.Execute "update phealthhub15_core set farmerbarcode='" & Mid(Trim(dgrid.TextMatrix(i, 2)), 1, 14) & "' where farmerbarcode='" & Mid(Trim(dgrid.TextMatrix(i, 1)), 1, 14) & "'"
    Next
    'distribution
    For i = 1 To dgrid.Rows - 1
    If Len(dgrid.TextMatrix(i, 1)) = 0 Then Exit For
    MHVDB.Execute "update tblplantdistributiondetail set farmercode='" & Mid(Trim(dgrid.TextMatrix(i, 2)), 1, 14) & "' where farmercode='" & Mid(Trim(dgrid.TextMatrix(i, 1)), 1, 14) & "'"
    Next
    'planted list
    For i = 1 To dgrid.Rows - 1
    If Len(dgrid.TextMatrix(i, 1)) = 0 Then Exit For
    MHVDB.Execute "update tblplanted set farmercode='" & Mid(Trim(dgrid.TextMatrix(i, 2)), 1, 14) & "' where farmercode='" & Mid(Trim(dgrid.TextMatrix(i, 1)), 1, 14) & "'"
    Next
    
    
    
    
    Case "X"
    
    
    
    
    
    
    End Select



End Sub

Private Sub Form_Load()
fillduplicate
filldstorage
filldfield
filldplanted
fillddist
End Sub
Private Sub fillduplicate()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mchk = True
dgrid.Clear
dgrid.Rows = 1
dgrid.FormatString = "^Sl.No.|^Duplicate Farmers|^Active Farmers|^Expected Drop outs|^"
dgrid.ColWidth(0) = 525
dgrid.ColWidth(1) = 2865
dgrid.ColWidth(2) = 2865
dgrid.ColWidth(3) = 2865
dgrid.ColWidth(4) = 180
If optduplicate.Value = True Then
rs.Open "select * from tblduplicate WHERE remarks<>'1' order by dfarmercode", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf optdropouts.Value = True Then
rs.Open "select * from tblduplicate WHERE remarks<>'1' order by dropfarmercode desc", MHVDB, adOpenForwardOnly, adLockOptimistic
End If
i = 1
Do While rs.EOF <> True



dgrid.Rows = dgrid.Rows + 1
dgrid.TextMatrix(i, 0) = i

If Len(rs!dfarmercode) > 0 Then
FindFA rs!dfarmercode, "F"
dgrid.TextMatrix(i, 1) = rs!dfarmercode & "  " & FAName
End If



If Len(rs!afarmercode) > 0 Then
FindFA rs!afarmercode, "F"
dgrid.TextMatrix(i, 2) = rs!afarmercode & "  " & FAName

End If

If Len(rs!dropfarmercode) > 0 Then
FindFA rs!dropfarmercode, "F"
dgrid.TextMatrix(i, 3) = rs!dropfarmercode & "  " & FAName

End If


rs.MoveNext
i = i + 1
Loop

rs.Close


dgrid.ColAlignment(1) = flexAlignLeftTop
dgrid.ColAlignment(2) = flexAlignLeftTop
dgrid.ColAlignment(3) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description



End Sub

Private Sub filldstorage()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mchk = True
sgrid.Clear
sgrid.Rows = 1
sgrid.FormatString = "^Sl.No.|^Farmer Name|^"
sgrid.ColWidth(0) = 525
sgrid.ColWidth(1) = 2865
sgrid.ColWidth(2) = 165


rs.Open "select distinct farmerbarcode from storagehub6_core where farmerbarcode in (select dfarmercode from tblduplicate) order by farmerbarcode", ODKDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
sgrid.Rows = sgrid.Rows + 1
sgrid.TextMatrix(i, 0) = i

FindFA rs!farmerbarcode, "F"
sgrid.TextMatrix(i, 1) = rs!farmerbarcode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close


sgrid.ColAlignment(1) = flexAlignLeftTop

Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub filldropstorage()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
sgrid.Clear
sgrid.Rows = 1
sgrid.FormatString = "^Sl.No.|^Farmer Name|^"
sgrid.ColWidth(0) = 525
sgrid.ColWidth(1) = 2865
sgrid.ColWidth(2) = 165

mchk = True
rs.Open "select distinct farmerbarcode  from storagehub6_core where farmerbarcode in (select dropfarmercode from tblduplicate) order by farmerbarcode", ODKDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
sgrid.Rows = sgrid.Rows + 1
sgrid.TextMatrix(i, 0) = i

FindFA rs!farmerbarcode, "F"
sgrid.TextMatrix(i, 1) = rs!farmerbarcode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
sgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description

End Sub


Private Sub fillastorage()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
sgrid.Clear
sgrid.Rows = 1
sgrid.FormatString = "^Sl.No.|^Farmer Name|^"
sgrid.ColWidth(0) = 525
sgrid.ColWidth(1) = 2865
sgrid.ColWidth(2) = 165

mchk = True
rs.Open "select distinct farmerbarcode  from storagehub6_core where farmerbarcode in (select  afarmercode from tblduplicate) order by farmerbarcode", ODKDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
sgrid.Rows = sgrid.Rows + 1
sgrid.TextMatrix(i, 0) = i
FindFA rs!farmerbarcode, "F"
sgrid.TextMatrix(i, 1) = rs!farmerbarcode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
sgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description

End Sub
Private Sub filldfield()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
fgrid.Clear
fgrid.Rows = 1
fgrid.FormatString = "^Sl.No.|^Farmer Name|^"
fgrid.ColWidth(0) = 525
fgrid.ColWidth(1) = 2865
fgrid.ColWidth(2) = 165


rs.Open "select distinct farmerbarcode  from phealthhub15_core where farmerbarcode in (select dfarmercode from tblduplicate) order by farmerbarcode", ODKDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
fgrid.Rows = fgrid.Rows + 1
fgrid.TextMatrix(i, 0) = i
FindFA rs!farmerbarcode, "F"
fgrid.TextMatrix(i, 1) = rs!farmerbarcode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
fgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description
End Sub


Private Sub filldropfield()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mchk = True
fgrid.Clear
fgrid.Rows = 1
fgrid.FormatString = "^Sl.No.|^Farmer Name|^"
fgrid.ColWidth(0) = 525
fgrid.ColWidth(1) = 2865
fgrid.ColWidth(2) = 165


rs.Open "select distinct farmerbarcode  from phealthhub15_core where farmerbarcode in (select dropfarmercode from tblduplicate) order by farmerbarcode", ODKDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
fgrid.Rows = fgrid.Rows + 1
fgrid.TextMatrix(i, 0) = i
FindFA rs!farmerbarcode, "F"
fgrid.TextMatrix(i, 1) = rs!farmerbarcode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
fgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description
End Sub


Private Sub fillafield()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
fgrid.Clear
fgrid.Rows = 1
fgrid.FormatString = "^Sl.No.|^Farmer Name|^"
fgrid.ColWidth(0) = 525
fgrid.ColWidth(1) = 2865
fgrid.ColWidth(2) = 165
mchk = True

rs.Open "select distinct farmerbarcode  from phealthhub15_core where farmerbarcode in (select afarmercode from tblduplicate) order by farmerbarcode", ODKDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
fgrid.Rows = fgrid.Rows + 1
fgrid.TextMatrix(i, 0) = i
FindFA rs!farmerbarcode, "F"
fgrid.TextMatrix(i, 1) = rs!farmerbarcode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
fgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub filldplanted()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
pgrid.Clear
pgrid.Rows = 1
pgrid.FormatString = "^Sl.No.|^Farmer Name|^"
pgrid.ColWidth(0) = 525
pgrid.ColWidth(1) = 2865
pgrid.ColWidth(2) = 165


rs.Open "select * from tblplanted where farmercode in (select dfarmercode from tblduplicate) order by farmercode", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
pgrid.Rows = pgrid.Rows + 1
pgrid.TextMatrix(i, 0) = i
FindFA rs!farmercode, "F"
pgrid.TextMatrix(i, 1) = rs!farmercode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
pgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description
End Sub


Private Sub filldropplanted()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
pgrid.Clear
pgrid.Rows = 1
pgrid.FormatString = "^Sl.No.|^Farmer Name|^"
pgrid.ColWidth(0) = 525
pgrid.ColWidth(1) = 2865
pgrid.ColWidth(2) = 165
mchk = True

rs.Open "select * from tblplanted where farmercode in (select dropfarmercode from tblduplicate) order by farmercode", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
pgrid.Rows = pgrid.Rows + 1
pgrid.TextMatrix(i, 0) = i
FindFA rs!farmercode, "F"
pgrid.TextMatrix(i, 1) = rs!farmercode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
pgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub fillaplanted()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
pgrid.Clear
pgrid.Rows = 1
pgrid.FormatString = "^Sl.No.|^Farmer Name|^"
pgrid.ColWidth(0) = 525
pgrid.ColWidth(1) = 2865
pgrid.ColWidth(2) = 165
mchk = True

rs.Open "select * from tblplanted where farmercode in (select afarmercode from tblduplicate) order by farmercode", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
pgrid.Rows = pgrid.Rows + 1
pgrid.TextMatrix(i, 0) = i
FindFA rs!farmercode, "F"
pgrid.TextMatrix(i, 1) = rs!farmercode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
pgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub optactive_Click()
MTYPE = "A"
fillastorage
fillafield
fillaplanted
filladist
End Sub

Private Sub optdropouts_Click()
fillduplicate
MTYPE = "X"
filldropstorage
filldropfield
filldropplanted
filldropdist
End Sub

Private Sub optduplicate_Click()
fillduplicate
MTYPE = "D"
filldstorage
filldfield
filldplanted
fillddist
End Sub




Private Sub fillddist()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
distgrid.Clear
distgrid.Rows = 1
distgrid.FormatString = "^Sl.No.|^Farmer Name|^"
distgrid.ColWidth(0) = 525
distgrid.ColWidth(1) = 2865
distgrid.ColWidth(2) = 165
mchk = True

rs.Open "select distinct farmercode  from tblplantdistributiondetail where farmercode in (select dfarmercode from tblduplicate) order by farmercode", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
distgrid.Rows = distgrid.Rows + 1
distgrid.TextMatrix(i, 0) = i
FindFA rs!farmercode, "F"
distgrid.TextMatrix(i, 1) = rs!farmercode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
distgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description
End Sub


Private Sub filldropdist()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
distgrid.Clear
distgrid.Rows = 1
distgrid.FormatString = "^Sl.No.|^Farmer Name|^"
distgrid.ColWidth(0) = 525
distgrid.ColWidth(1) = 2865
distgrid.ColWidth(2) = 165

mchk = True
rs.Open "select distinct farmercode  from tblplantdistributiondetail where farmercode in (select dropfarmercode from tblduplicate) and length(farmercode)>0  order by farmercode ", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
distgrid.Rows = distgrid.Rows + 1
distgrid.TextMatrix(i, 0) = i
FindFA rs!farmercode, "F"
distgrid.TextMatrix(i, 1) = rs!farmercode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
distgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description
End Sub


Private Sub filladist()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
distgrid.Clear
distgrid.Rows = 1
distgrid.FormatString = "^Sl.No.|^Farmer Name|^"
distgrid.ColWidth(0) = 525
distgrid.ColWidth(1) = 2865
distgrid.ColWidth(2) = 165

mchk = True
rs.Open "select distinct farmercode  from tblplantdistributiondetail where farmercode in (select afarmercode from tblduplicate) order by farmercode", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
distgrid.Rows = distgrid.Rows + 1
distgrid.TextMatrix(i, 0) = i
FindFA rs!farmercode, "F"
distgrid.TextMatrix(i, 1) = rs!farmercode & "  " & FAName
rs.MoveNext
i = i + 1
Loop

rs.Close
distgrid.ColAlignment(1) = flexAlignLeftTop
Exit Sub
err:
MsgBox err.Description
End Sub
