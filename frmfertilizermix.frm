VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmfertilizermix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F E R T I L I Z E R  M I X . . . "
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10275
   Icon            =   "frmfertilizermix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Chemical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   7335
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   2775
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7215
         _cx             =   12726
         _cy             =   4895
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   8438015
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
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmfertilizermix.frx":0E42
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7335
      Begin VB.TextBox txtvolume 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5640
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker txtdate 
         Height          =   375
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   111017985
         CurrentDate     =   41479
      End
      Begin MSDataListLib.DataCombo cbofertilizer 
         Bindings        =   "frmfertilizermix.frx":0EC8
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cbostaffid 
         Bindings        =   "frmfertilizermix.frx":0EDD
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Volume"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   9
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   4
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fertilizer Mix No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1440
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   5280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfertilizermix.frx":0EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfertilizermix.frx":128C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfertilizermix.frx":1626
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfertilizermix.frx":2300
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfertilizermix.frx":2752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfertilizermix.frx":2F0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   1164
      ButtonWidth     =   1217
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ADD"
            Key             =   "ADD"
            Object.ToolTipText     =   "ADDS NEW RECORD"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "New"
                  Text            =   "New"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Open"
                  Text            =   "Open"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OPEN"
            Key             =   "OPEN"
            Object.ToolTipText     =   "OPEN/EDIT EXISTING RECORD"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "SAVE"
            Key             =   "SAVE"
            Object.ToolTipText     =   "SAVES RECORD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "DELETE"
            Key             =   "DELETE"
            Object.ToolTipText     =   "DELETE THE RECORD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EXIT"
            Key             =   "EXIT"
            Object.ToolTipText     =   "EXIT FROM THE FORM"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   3
   End
End
Attribute VB_Name = "frmfertilizermix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbofertilizer_LostFocus()
On Error GoTo err
   
   cbofertilizer.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblqmsfertilizermixhdr where fertilizermixno='" & cbofertilizer.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
   fillcontroll cbofertilizer.BoundText
   
   Else
   MsgBox "Record Not Found."
   End If
   rs.Close
   Exit Sub
err:
   MsgBox err.Description
   'rs.Close
End Sub
Private Sub fillcontroll(id As String)
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmsfertilizermixhdr where fertilizermixno = '" & id & "'", MHVDB
If rs.EOF <> True Then
   txtdate.Value = Format(rs!mixeddate, "dd/MM/yyyy")
   
   txtvolume.Text = IIf(rs!totalqty = 0, "", rs!totalqty)
   FindsTAFF rs!staffid
   cbostaffid.Text = rs!staffid & " " & sTAFF
 
End If
fillgrd id
End Sub
Private Sub fillgrd(id As String)
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
rs.Open "select * from tblqmsfertilizermixdetail where fertilizermixno = '" & id & "'", MHVDB
If rs.EOF <> True Then
i = 1
    Do While rs.EOF <> True
   mygrid.TextMatrix(i, 0) = i
   FindqmsChemicalTradeName rs!chemicalid
   mygrid.TextMatrix(i, 1) = qmsChemicalTradeName
   mygrid.TextMatrix(i, 2) = rs!qty
      
        i = i + 1
        rs.MoveNext
    Loop
End If
End Sub

Private Sub Form_Load()
On Error GoTo err
Operation = ""
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString



Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select fertilizermixno as description  from tblqmsfertilizermixhdr order by fertilizermixno", db
Set cbofertilizer.RowSource = rsF
cbofertilizer.ListField = "description"
cbofertilizer.BoundColumn = "description"

Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff where nursery='1' order by STAFFCODE", db
Set cbostaffid.RowSource = rsF
cbostaffid.ListField = "STAFFNAME"
cbostaffid.BoundColumn = "STAFFCODE"



Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub mygrid_Click()
If mygrid.col = 1 And Len(mygrid.TextMatrix(mygrid.row - 1, 2)) > 0 Then
mygrid.Editable = flexEDKbdMouse
FillGridCombo
Else
mygrid.ComboList = ""
mygrid.Editable = flexEDNone
End If

If Len(mygrid.TextMatrix(mygrid.row, 1)) > 0 And mygrid.col = 2 Then
mygrid.Editable = flexEDKbdMouse
End If

End Sub
Private Sub FillGridCombo()
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        StrComboList = "          |"
        
        
            Set RstTemp = Nothing
            RstTemp.Open ("select chemicalid,tradename from tblqmschemicalhdr where status='ON' ORDER BY chemicalid"), MHVDB

            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                    NmcItemData = ""
                    StrComboList = IIf(StrComboList = "", Right("0000" & RstTemp("chemicalid").Value, 3) & " " & RstTemp("tradename").Value, StrComboList & "|" & Right("0000" & RstTemp("chemicalid").Value, 3)) & " " & RstTemp("tradename").Value
                    RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       mygrid.ComboList = StrComboList



    End Sub

Private Sub Mygrid_LeaveCell()
Dim i As Integer
If Trim(mygrid.TextMatrix(mygrid.row, 1)) = "" Then
mygrid.RemoveItem (mygrid.row)
mygrid.rows = mygrid.rows + 1
For i = 1 To mygrid.rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
mygrid.TextMatrix(i, 0) = i
Next

End If
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
            
        cbofertilizer.Enabled = False
        TB.buttons(3).Enabled = True
        Operation = "ADD"
        CLEARCONTROLL
        Dim rs As New ADODB.Recordset
        Set rs = Nothing
        rs.Open "SELECT MAX(fertilizermixno )+1 AS MaxID from tblqmsfertilizermixhdr", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        cbofertilizer.Text = IIf(IsNull(rs!MaxId), 1, rs!MaxId)
        Else
        cbofertilizer.Text = rs!MaxId
        End If
       Case "OPEN"
        Operation = "OPEN"
        CLEARCONTROLL
        cbofertilizer.Enabled = True
        TB.buttons(3).Enabled = True
             
       Case "SAVE"
       MNU_SAVE

       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub CLEARCONTROLL()
   
   cbostaffid.Text = ""
   txtvolume.Text = ""
mygrid.Clear
mygrid.FormatString = "^Sl.No.|^Chemical Name|^Kg./Litres|^"
mygrid.ColWidth(0) = 960
mygrid.ColWidth(1) = 3795
mygrid.ColWidth(2) = 1320
mygrid.ColWidth(3) = 1005
End Sub
Private Sub MNU_SAVE()
Dim rs As New ADODB.Recordset
On Error GoTo err
Dim i As Integer
If Len(cbofertilizer.Text) = 0 Then
MsgBox "Input the Fertilizer Mix No."
Exit Sub
End If

MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tblqmsfertilizermixhdr (fertilizermixno,mixeddate,totalqty,staffid," _
            & "status,location) " _
            & "VALUEs('" & cbofertilizer.Text & "','" & Format(txtdate.Value, "yyyy-MM-dd") & "','" & Val(txtvolume.Text) & "', " _
            & " '" & cbostaffid.BoundText & "','ON','" & Mlocation & "')"
 
 
LogRemarks = "Inserted new record" & cbofertilizer.BoundText & "," & Format(txtdate.Value, "yyyy-MM-dd") & "," & Mlocation & "," & txtremarks
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmsfertilizermixhdr set " _
            & "mixeddate='" & Format(txtdate.Value, "yyyy-MM-dd") & "',totalqty='" & Val(txtvolume.Text) & "' " _
            & ",staffid='" & cbostaffid.BoundText & "' " _
            & " where fertilizermixno='" & cbofertilizer.BoundText & "' and location='" & Mlocation & "'"

LogRemarks = "Updated  record" & cbofertilizer.BoundText & "," & Format(txtdate.Value, "yyyy-MM-dd") & "," & Mlocation
updatemhvlog Now, MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
MHVDB.RollbackTrans
Exit Sub
End If

MHVDB.Execute "delete from tblqmsfertilizermixdetail where fertilizermixno='" & cbofertilizer.BoundText & "'"

For i = 1 To mygrid.rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
MHVDB.Execute "insert into tblqmsfertilizermixdetail (fertilizermixno,mixeddate,chemicalid,qty,location) values" _
            & "('" & cbofertilizer.Text & "','" & Format(txtdate.Value, "yyyy-MM-dd") & "','" & Mid(mygrid.TextMatrix(i, 1), 1, 3) & "', " _
            & " '" & Val(mygrid.TextMatrix(i, 2)) & "','" & Mlocation & "')"

Next



MHVDB.CommitTrans
TB.buttons(3).Enabled = False
Exit Sub

err:
MsgBox err.Description
MHVDB.RollbackTrans


End Sub

