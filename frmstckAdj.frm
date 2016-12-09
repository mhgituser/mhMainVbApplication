VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmsTOCKAdj 
   Caption         =   "ADJUSTENT"
   ClientHeight    =   5055
   ClientLeft      =   6465
   ClientTop       =   1020
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmstckAdj.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7725
   Visible         =   0   'False
   Begin MSComCtl2.DTPicker lblnow 
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   675
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   81985537
      CurrentDate     =   37084
   End
   Begin VB.TextBox txtRemark 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   7455
      Begin MSDataListLib.DataCombo CboItemDesc 
         Bindings        =   "frmstckAdj.frx":076A
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1440
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
      End
      Begin VB.TextBox cboItemCode 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDBCtls.DBCombo CboItemDesc1 
         Bindings        =   "frmstckAdj.frx":077F
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         ListField       =   "itemname"
         BoundColumn     =   "itemcode"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid ItemGrd 
         Height          =   3060
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5398
         _Version        =   393216
         Rows            =   300
         Cols            =   6
         RowHeightMin    =   315
         ForeColorFixed  =   -2147483635
         ScrollTrack     =   -1  'True
         HighLight       =   0
         ScrollBars      =   2
         FormatString    =   $"frmstckAdj.frx":0798
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Remarks :"
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
         Index           =   8
         Left            =   240
         TabIndex        =   7
         Top             =   4920
         Width           =   870
      End
      Begin VB.Line Line2 
         X1              =   5280
         X2              =   8880
         Y1              =   4290
         Y2              =   4290
      End
   End
   Begin MSDataListLib.DataCombo CboBillNo 
      Bindings        =   "frmstckAdj.frx":0822
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1680
      TabIndex        =   13
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   10560
      Top             =   720
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
            Picture         =   "frmstckAdj.frx":0837
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstckAdj.frx":0BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstckAdj.frx":0F6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstckAdj.frx":1C45
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstckAdj.frx":2097
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstckAdj.frx":2851
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7725
      _ExtentX        =   13626
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
   Begin VB.Label Label4 
      Caption         =   "Date"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblYR 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " "
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   2280
      TabIndex        =   9
      Top             =   720
      Width           =   60
   End
   Begin VB.Label Label3 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Entry No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmsTOCKAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BillDetailRec As New ADODB.Recordset
Dim Bill As New ADODB.Recordset
Dim rsdatInvItem As New ADODB.Recordset
Dim rsDatBrBill As New ADODB.Recordset
Dim CurrRow, Jkey As Long
Dim ValidRow As Boolean
Dim Operation As String
Dim ltot As Double
Const fmString = "       |^ Code       |^                                    Item Name                                    |^ Unit     |^   Qty          |"
Private Sub PrintBill()
Dim i As Integer
Dim pfile As String
pfile = "SAD" + Trim(CboBillNo) + ".txt"
On Error GoTo jerr:
Do While True
   Select Case MsgBox("Printer Ready ? ", vbYesNoCancel)
          Case vbYes
               Open "lpt1:" For Output As #1
               Exit Do
          Case vbNo
               If MsgBox("Connect/Swich on the printer and retry.", vbRetryCancel) = vbCancel Then
                  MsgBox "Bill stored to " + pfile
                  Open pfile For Output As #1
                  Exit Do
               End If
          Case vbCancel
               MsgBox "Bill stored to " + pfile
               Open pfile For Output As #1
               Exit Do
   End Select
Loop
'Open "lpt1:" For Output As #1
Print #1, Chr(18);
Print #1, Chr(14) + "MHV" + Chr(20)
Print #1, "Mongar : Bhutan"
Print #1, "Beverage Bill" + PadWithChar("Bill No. : ", 17, " ", 1) + PadWithChar(CboBillNo, 12, " ", 1)
Print #1, "Date /Time : " + lblnow
If Operation = "Open" Then
   Print #1, Chr(15) + PadWithChar("MODIFIED Bill ", 72, " ", 1) + Chr(18)
End If
Print #1, Chr(15); String(72, "-")
Print #1, "NAME : " + UCase(txtBillTo)
Select Case CboPMode.ListIndex
       Case 0
       Case 1
       Print #1, "Room No : " + txtRoomNo + " ; Registration No. : " + txtRoomNo.Tag
       Case 2
       Print #1, "Bill to : " + IIf(DBcboParty.Visible, DBcboParty, "")
End Select
Print #1, String(72, "-")
Print #1, "Srl.|Code  |Description             |    Qty.      |  Rate  |     Amount  "
Print #1, String(72, "-")
For i = 1 To ItemGrd.Rows - 1
    If Len(Trim(ItemGrd.TextMatrix(i, 1))) = 0 Then Exit For
    Print #1, PadWithChar(str(i), 4, " ", 1) + " " + PadWithChar(ItemGrd.TextMatrix(i, 1), 7, " ", 0);
    Print #1, PadWithChar(ItemGrd.TextMatrix(i, 2), 24, " ", 0) + PadWithChar(ItemGrd.TextMatrix(i, 4), 7, " ", 1) + " " + PadWithChar(ItemGrd.TextMatrix(i, 3), 7, " ", 0);
    Print #1, PadWithChar(ItemGrd.TextMatrix(i, 5), 9, " ", 1) + " " + PadWithChar(ItemGrd.TextMatrix(i, 6), 11, " ", 1)
    If Len(ItemGrd.TextMatrix(i, 2)) > 24 Then
       Print #1, "     " + Mid(ItemGrd.TextMatrix(i, 2), 25)
    End If
Next
Print #1, String(72, "-")
Print #1, PadWithChar("Total", 50, " ", 1) + PadWithChar(lblTot, 22, " ", 1)
Print #1, String(72, "-")
Print #1, PadWithChar("Discount (%)", 50, " ", 1) + PadWithChar(txtDisc, 7, " ", 1) + PadWithChar(lbldisc, 15, " ", 1)
Print #1, String(72, "-")
Print #1, PadWithChar("Amount payable", 50, " ", 1) + PadWithChar(lblAmt, 22, " ", 1)
Print #1, Chr(18)
If CboPMode.ListIndex = 0 Then
   Print #1, Chr(14) + "CASH PAYMENT" + Chr(20)
Else
   Print #1, "SIGNATURE"
   Print #1, "(Please do not sign, if you have paid)"
End If
Print #1, Chr(18)
Print #1,
Print #1,
Print #1,
Print #1,
Print #1,
Print #1,
Print #1,
Print #1,
Print #1,
Print #1,
Close #1 '*/
Exit Sub
jerr:
MsgBox err.Description
err.Clear
End Sub



Private Sub CboBillNo_GotFocus()
'Frame2.Enabled = False
End Sub

Private Sub CboBillNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub CboBillNo_LostFocus()
Dim i As Integer
Dim Issue, Recv As Double
Dim imast As ADODB.Recordset
If Operation = "Add" Then Exit Sub
Set Bill = MHVDB.Execute("select * from tranhdr where procyear='" & SysYear & "' and ((billno))=('" & CboBillNo & "') AND billtype = 'AD' and status<>'C'")
If Bill.EOF Then
   MsgBox CboBillNo + " Does not exists "
   'CboBillNo.SetFocus
   Exit Sub
Else
   With Bill
   lblnow = Format(!billdate, "dd/mm/yyyy") ' DatKotHead.Recordset!Time
   txtRemark = IIf(IsNull(!remarks), "", !remarks)
   End With
   Set BillDetailRec = MHVDB.Execute("select d.itemcode,d.qty,D.RATE,B.unit,itemname,avgstockrate from tranfile as d,invitems as b where d.itemcode=b.itemcode and d.procyear='" & SysYear & "' and d.billtype='AD' and billno=('" & CboBillNo & "')")
   ItemGrd.Clear
   ItemGrd.FormatString = fmString
   With BillDetailRec
   i = 1
   ltot = 0
   Do While Not .EOF
      ItemGrd.TextMatrix(i, 1) = !itemcode
      ItemGrd.TextMatrix(i, 2) = !ITEMNAME
      ItemGrd.TextMatrix(i, 3) = !unit
      ItemGrd.TextMatrix(i, 4) = !qty
      ItemGrd.TextMatrix(i, 5) = !Rate
      .MoveNext
      i = i + 1
   Loop
   End With
'   If MsgBox("Cant be modified.You can Cancel or Print It ! Do you want Print ?", vbYesNo) = vbYes Then
'      PrintBill
'   End If
'   Operation = ""
  ' txtRemark.SetFocus
End If
Frame2.Enabled = True
TB.Buttons(3).Enabled = True

TB.Buttons(4).Enabled = True
End Sub
Private Sub cboItemCode_LostFocus()
Dim prevamt, CurrAmt As Double
cboItemCode.Text = UCase(cboItemCode.Text)
If ItemGrd.TextMatrix(CurrRow, 1) = cboItemCode Then
   cboItemCode.Visible = False
   Exit Sub
End If
'prevamt = Val(ItemGrd.TextMatrix(CurrRow, 7))
'rsdatInvItem.Recordset.FindFirst "trim(itemcode) = trim('" & cboItemCode & "')"
 rsdatInvItem.Find " itemcode='" & cboItemCode & "'", , adSearchForward, 1
With rsdatInvItem
If .EOF Then
   ItemGrd.TextMatrix(CurrRow, 1) = ""
   ItemGrd.TextMatrix(CurrRow, 2) = ""
   ItemGrd.TextMatrix(CurrRow, 3) = ""
   ItemGrd.TextMatrix(CurrRow, 4) = ""
   ItemGrd.TextMatrix(CurrRow, 5) = ""
   txtQty = ""
   ValidRow = True
Else
   ItemGrd.TextMatrix(CurrRow, 1) = !itemcode
   ItemGrd.TextMatrix(CurrRow, 2) = !ITEMNAME
   ItemGrd.TextMatrix(CurrRow, 3) = !unit
   ItemGrd.TextMatrix(CurrRow, 5) = !avgstockRate
   txtQty = ItemGrd.TextMatrix(CurrRow, 4)
   If Not Val(txtQty) > 0 Then
      ValidRow = False
      ItemGrd.row = CurrRow
      txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
      txtQty.Visible = True
      txtQty.SetFocus
   End If

End If
End With

cboItemCode.Visible = False
End Sub

Private Sub cboItemDesc_LostFocus()
Dim Issue, Recv As Double
If ItemGrd.TextMatrix(CurrRow, 1) = CboItemDesc.BoundText Then
   CboItemDesc.Visible = False
   Exit Sub
End If
'datInvItem.Recordset.FindFirst "itemcode = '" & CboItemDesc.BoundText & "'"
  rsdatInvItem.Find " itemcode='" & CboItemDesc.BoundText & "'", , adSearchForward, 1
With rsdatInvItem
If .EOF Then
   ItemGrd.TextMatrix(CurrRow, 1) = ""
   ItemGrd.TextMatrix(CurrRow, 2) = ""
   ItemGrd.TextMatrix(CurrRow, 3) = ""
   ItemGrd.TextMatrix(CurrRow, 4) = ""
   ItemGrd.TextMatrix(CurrRow, 5) = ""
   txtQty = ""
   ValidRow = True
Else
   ItemGrd.TextMatrix(CurrRow, 1) = !itemcode
   ItemGrd.TextMatrix(CurrRow, 2) = !ITEMNAME
   ItemGrd.TextMatrix(CurrRow, 3) = !unit
   ItemGrd.TextMatrix(CurrRow, 5) = !avgstockRate
   txtQty = ItemGrd.TextMatrix(CurrRow, 4)
   If Not Val(txtQty) > 0 Then
      ValidRow = False
      ItemGrd.row = CurrRow
      txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
      txtQty.Visible = True
      txtQty.SetFocus
   End If
End If
End With
CboItemDesc.Visible = False
End Sub




Private Sub DBcboParty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub DBcboParty_Validate(Cancel As Boolean)
Dim Party As ADODB.Recordset
If Len(Trim(DBcboParty.BoundText)) = 0 Then
   MsgBox "Supplier should not be blank "
   Exit Sub
End If
Set Party = MHVDB.Execute("SELECT suplCODE FROM supplier WHERE suplCODE='" & DBcboParty.BoundText & "'", dbOpenDynaset, DBReadOnly)
If Party.EOF Then
   MsgBox "Not a valid Supplier !!!"
   Cancel = True
End If
End Sub

Private Sub Form_Load()
'Set datInvItem.Recordset = MHVDB.Execute("select * from invitems order by itemname", dbOpenDynaset, dbReadOnly)
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
' ITEM
If rsdatInvItem.State = adStateOpen Then rsdatInvItem.Close
rsdatInvItem.Open "select *  from invitems order by itemname", db
Set CboItemDesc.RowSource = rsdatInvItem
CboItemDesc.ListField = "ITEMNAME"
CboItemDesc.BoundColumn = "ItemCode"


'Set rsDatBrBill.Recordset = MHVDB.Execute("select * from tranhdr where status <> 'C' and billtype='AD' and procyear='" & SysYear & "' order by billno desc", dbOpenDynaset)
' BILLNO
If rsDatBrBill.State = adStateOpen Then rsDatBrBill.Close
 rsDatBrBill.Open "select * from tranhdr where status <> 'C' and billtype='AD' and procyear='" & SysYear & "' order by billno desc", db
Set CboBillNo.RowSource = rsDatBrBill
CboBillNo.ListField = "billno"
CboBillNo.BoundColumn = "billno"


lblYR = "AD\" & SysYear & "\"
ValidRow = True
CurrRow = 1
ItemGrd.row = 1
ItemGrd.col = 1
cboItemCode.Left = ItemGrd.Left + ItemGrd.CellLeft
cboItemCode.Width = ItemGrd.CellWidth
cboItemCode.Height = ItemGrd.CellHeight
ItemGrd.col = 2
CboItemDesc.Left = ItemGrd.Left + ItemGrd.CellLeft
CboItemDesc.Width = ItemGrd.CellWidth
CboItemDesc.Height = ItemGrd.CellHeight
ItemGrd.col = 4
txtQty.Left = ItemGrd.Left + ItemGrd.CellLeft
txtQty.Width = ItemGrd.CellWidth
txtQty.Height = ItemGrd.CellHeight

End Sub
Private Sub ItemGrd_Click()
Dim jrow, jCol As Integer
If Not ValidRow And CurrRow <> ItemGrd.row Then
   ItemGrd.row = CurrRow
   Exit Sub
End If
jrow = ItemGrd.row
jCol = ItemGrd.col
If jrow = 0 Then Exit Sub
If jrow > 1 And Len(ItemGrd.TextMatrix(jrow - 1, 1)) = 0 Then
   Beep
   Exit Sub
End If
ItemGrd.TextMatrix(CurrRow, 0) = CurrRow
CurrRow = jrow
ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)
Select Case jCol
       
       Case 1
            cboItemCode.Top = ItemGrd.Top + ItemGrd.CellTop
            cboItemCode = ItemGrd.Text
            cboItemCode.Visible = True
            cboItemCode.SetFocus
       Case 2
            CboItemDesc.Top = ItemGrd.Top + ItemGrd.CellTop
            CboItemDesc = ItemGrd.Text
            CboItemDesc.BoundText = ItemGrd.TextMatrix(CurrRow, 1)
            CboItemDesc.Visible = True
            CboItemDesc.SetFocus
       Case 4
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
               txtQty = ItemGrd.Text
               txtQty.Visible = True
               txtQty.SetFocus
            End If
       Case 5
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtRate.Top = ItemGrd.Top + ItemGrd.CellTop
               txtRate = ItemGrd.Text
               txtRate.Visible = True
               txtRate.SetFocus
            End If
    End Select
End Sub

Private Sub ItemGrd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And Shift = 2 Then
   If MsgBox("Are u sure to Delete this row ?", vbYesNo) = vbNo Then Exit Sub
   If CurrRow > 0 And Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
'      ltot = ltot - val(ItemGrd.TextMatrix(CurrRow, 6))
 '     lblTot.Caption = Format(ltot, "######0.00")
      ItemGrd.RemoveItem CurrRow
      ItemGrd.AddItem ""
   Else
      Beep
      Beep
   End If
End If
End Sub

Private Sub ItemGrd_Scroll()
SendKeys "{TAB}", True
End Sub

Private Sub mnuadd_Click()
Dim lastbill As ADODB.Recordset
lblnow = Format(Now, "dd/mm/yyyy")
ValidRow = True
Operation = "ADD"
CurrRow = 1
txtRemark = ""
cboItemCode.Visible = False
CboItemDesc.Visible = False
txtQty.Visible = False
Set lastbill = MHVDB.Execute("select max(billno) as lno from tranhdr where procyear='" & SysYear & "'  and billtype='AD'", dbOpenForwardOnly)
CboBillNo = IIf(IsNull(lastbill!lno), 1, lastbill!lno + 1)
Set lastbill = Nothing
CboBillNo.Enabled = False
Frame2.Enabled = True
ItemGrd.Enabled = True
ItemGrd.Clear
ItemGrd.FormatString = fmString

TB.Buttons(4).Enabled = False

TB.Buttons(3).Enabled = True
End Sub
Private Sub mnuCancel_Click()
Dim UpdtStr
Dim jrec As ADODB.Recordset
If MsgBox("Cancel it !!!Are u Sure ?", vbYesNo) = vbNo Then Exit Sub
On Error GoTo err
MHVDB.BeginTrans
UpdtStr = "UPDATE  tranhdr SET STATUS = 'C',REMARKs = '" & txtRemark & "' WHERE  procyear='" & SysYear & "' and billno = VAL('" & CboBillNo & "') and billtype='AD'"
Set jrec = MHVDB.Execute("select * from tranfile where billno =val('" & CboBillNo & "') and billtype='AD'", dbOpenDynaset)
With jrec
Do While Not .EOF
   MHVDB.Execute "UPDATE ITEMSTOCK SET totadj=totadj-val('" & !qty & "') where  procyear='" & SysYear & "' and itemcode='" & !itemcode & "'", dbFailOnError
   .MoveNext
Loop
End With
MHVDB.Execute UpdtStr, dbSeeChanges + dbFailOnError
Frame2.Enabled = False
MHVDB.CommitTrans
DatBrBill.Refresh
Operation = ""
CboBillNo.Enabled = False

TB.Buttons(4).Enabled = False
Exit Sub
err:
MsgBox "error :" + IIf(IsNull(err.Description), " ", err.Description)
err.Clear
MHVDB.Rollback
End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuOpen_Click()
Operation = "OPEN"
'ItemGrd.Enabled = False
Frame2.Enabled = True
CboBillNo.Enabled = True

TB.Buttons(3).Enabled = False

TB.Buttons(4).Enabled = 0
CboBillNo.Refresh
End Sub
Private Sub mnuSave_Click()
Dim i, j, K As Integer
Dim printNow As Boolean
Dim InsStr, JStat, pcODE As String
'If txtQty.Visible Then txtQty_validate
If Not (Operation = "OPEN" Or Operation = "ADD") Then
   Beep
   Exit Sub
End If
If Not ValidRow Then Exit Sub
printNow = True
On Error GoTo err

If Operation = "ADD" Then
   
   InsStr = "insert into tranHdR (procyear,billtype, billno,BILLDATE,Status,remarks) values ( '" & SysYear & "','AD','" & CboBillNo & "'," _
                  & " '" & Format(lblnow, "yyyyMMdd") & "','OK','" & txtRemark & "')"
   MHVDB.Execute InsStr, dbSeeChanges + dbFailOnError
   For i = 1 To 394
       If Len(Trim(ItemGrd.TextMatrix(i, 1))) > 0 Then
          InsStr = "insert into tranfile ( procyear,billtype,billno,itemcode,qty,rate,unit) values ( '" & SysYear & "','AD','" & CboBillNo & "'," _
                  & " '" & ItemGrd.TextMatrix(i, 1) & "','" & ItemGrd.TextMatrix(i, 4) & "','" & ItemGrd.TextMatrix(i, 5) & "','" & ItemGrd.TextMatrix(i, 3) & "')"
          MHVDB.Execute InsStr, dbSeeChanges + dbFailOnError
          MHVDB.Execute "update ITEMSTOCK set totadj=totadj+('" & ItemGrd.TextMatrix(i, 4) & "') where procyear='" & SysYear & "' and ITEMCODE = '" & ItemGrd.TextMatrix(i, 1) & "'", dbFailOnError
       Else
          Exit For
       End If
   Next
Else
   InsStr = "update tranHdR set remarks='" & txtRemark & "',billdate='" & Format(lblnow, "yyyyMMdd") & "' where  procyear='" & SysYear & "' and billno =( '" & CboBillNo & "') and billtype='AD'"
   MHVDB.Execute InsStr, dbSeeChanges + dbFailOnError
   Set jrec = MHVDB.Execute("select * from tranfile where procyear='" & SysYear & "' and billno =('" & CboBillNo & "') and billtype='AD'", dbOpenDynaset)
   With jrec
   Do While Not .EOF
      MHVDB.Execute "update ITEMSTOCK set totadj=totadj-('" & !qty & "') where procyear='" & SysYear & "' and ITEMCODE = '" & !itemcode & "'", dbFailOnError
     .MoveNext
   Loop
   End With
   MHVDB.Execute "delete  from tranfile where procyear='" & SysYear & "' and billno =('" & CboBillNo & "') and billtype='AD'", dbFailOnError
   For i = 1 To 394
       If Len(Trim(ItemGrd.TextMatrix(i, 1))) > 0 Then
          InsStr = "insert into tranfile ( procyear,billtype,billno,itemcode,qty,rate,unit) values ( '" & SysYear & "','AD','" & CboBillNo & "'," _
                  & " '" & ItemGrd.TextMatrix(i, 1) & "','" & ItemGrd.TextMatrix(i, 4) & "','" & ItemGrd.TextMatrix(i, 5) & "','" & ItemGrd.TextMatrix(i, 3) & "')"
          MHVDB.Execute InsStr, dbSeeChanges + dbFailOnError
          MHVDB.Execute "update ITEMSTOCK set totadj=totadj+('" & ItemGrd.TextMatrix(i, 4) & "') where procyear='" & SysYear & "' and ITEMCODE = '" & ItemGrd.TextMatrix(i, 1) & "'", dbFailOnError
       Else
          Exit For
       End If
   Next
End If
printNow = IIf(MsgBox("Print Now ?", vbYesNo) = vbYes, True, False)
'If printNow Then PrintBill

'DatBrBill.Refresh
Operation = ""
CboBillNo.Enabled = False
Frame2.Enabled = False

TB.Buttons(3).Enabled = False

TB.Buttons(4).Enabled = 0
Exit Sub
err:
MsgBox err.Description
If DBEngine.Errors.Count > 0 Then
For Each errLoop In DBEngine.Errors
    MsgBox "Error number: " & errLoop.Number & vbCr & _
    errLoop.Description
Next errLoop
'Exit Sub
End If
err.Clear

End Sub
Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
       Case "ADD"
           mnuadd_Click
       Case "OPEN"
           mnuOpen_Click
       Case "SAVE"
           mnuSave_Click
       Case "DELETE"
          ' mnuCancel_Click
       Case "EXIT"
           Unload Me
End Select
End Sub


Private Sub txtChDate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub txtQty_validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not IsNumeric(txtQty) Then
   Beep
   MsgBox "Enter a valid Quantity"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   ItemGrd.TextMatrix(CurrRow, 4) = Val(txtQty)
   ValidRow = True
End If
End If
txtQty.Visible = False
End Sub


