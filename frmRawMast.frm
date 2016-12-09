VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmitemMAST 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "STORE ITEM MASTER MAINTAINANCE"
   ClientHeight    =   6075
   ClientLeft      =   6615
   ClientTop       =   1275
   ClientWidth     =   8310
   Icon            =   "frmRawMast.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Frmraw"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtSpec 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      MaxLength       =   60
      TabIndex        =   28
      Text            =   " "
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   7320
      TabIndex        =   2
      Text            =   "1"
      ToolTipText     =   "eg. pur. in kg sale in 200gm pack factor=.5"
      Top             =   3120
      Width           =   615
   End
   Begin VB.ComboBox CboUnit 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      ItemData        =   "frmRawMast.frx":0E42
      Left            =   4560
      List            =   "frmRawMast.frx":0E61
      TabIndex        =   1
      Text            =   "Combo1"
      ToolTipText     =   "Choose unit"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   12
      Top             =   3960
      Width           =   7935
      Begin VB.TextBox TxtFuel 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5880
         TabIndex        =   33
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtOpRate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtOp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox LBLRATE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "Fuel Unit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   34
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label OpVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   7680
         TabIndex        =   32
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label1 
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Year Opening Stock"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   2
         Left            =   2400
         TabIndex        =   24
         Top             =   600
         Width           =   105
      End
      Begin VB.Label Label3 
         Caption         =   "Current Stock"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   1
         Left            =   4560
         TabIndex        =   22
         Top             =   600
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   16
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label txtQty 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label txtValue 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Avg. Rate"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   6480
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Enter the quantity below which stock you want to make a pur. order"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ComboBox CboUnit 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmRawMast.frx":0E9F
      Left            =   1680
      List            =   "frmRawMast.frx":0EB5
      TabIndex        =   0
      ToolTipText     =   "Choose unit"
      Top             =   3120
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo cboItemCode 
      Bindings        =   "frmRawMast.frx":0ED6
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1920
      TabIndex        =   35
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboCatGrp 
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   36
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483643
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboCatGrp 
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   37
      Top             =   2640
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483643
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo txtDesc 
      Bindings        =   "frmRawMast.frx":0EEB
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1920
      TabIndex        =   38
      Top             =   1200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   6360
      Top             =   600
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
            Picture         =   "frmRawMast.frx":0F00
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRawMast.frx":129A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRawMast.frx":1634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRawMast.frx":230E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRawMast.frx":2760
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRawMast.frx":2F1A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   8310
      _ExtentX        =   14658
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
      Caption         =   "Specification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Conv. Factor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5880
      TabIndex        =   19
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Purchase Unit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   18
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Minimum Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Reorder Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Item Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Description :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Issue Unit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Group:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmitemMAST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db
Dim ItemMast As New ADODB.Recordset
Dim datInvItem As New ADODB.Recordset
Dim Datcategory As New ADODB.Recordset
Dim datGroup As New ADODB.Recordset
Dim DatLoc As New ADODB.Recordset
Dim DatComp As New ADODB.Recordset
Dim ErrCTR As Long
Private Sub IHIST()
Dim pfile As String
Dim jrec As ADODB.Recordset
Dim JQty, TotAmt As Double
pfile = App.Path + "\IHIST.TXT"
Open pfile For Output As #1
Set jrec = MHVDB.Execute("SELECT A.BILLDATE,A.BILLNO,B.QTY,c.deptname from tranhdr AS A,tranfile AS B ,Departments as c WHERE A.STATUS<>'C' and A.PROCYEAR='" & SysYear & "' and a.billtype='II' AND A.BILLtype=B.BILLtype AND A.BILLNO=B.BILLNO and a.procyear=b.procyear AND B.ITEMCODE='" & cboItemCode & "' and a.suplcode=c.deptcode ORDER BY BILLDATE ", dbOpenSnapshot)
JQty = 0
Print #1, Chr(14) + JsysName + Chr(20)
   Print #1, "Issue history of " + txtDesc
   Print #1, String(70, "-")
   Print #1, "Date       BILLNO              Qty       Department    "
   Print #1, String(70, "-")
With jrec
Do While Not .EOF
   Print #1, Format(!billdate, "dd/mm/yyyy") + " " + PadWithChar(Str(!billno), 12, " ", 0) + " " + PadWithChar(Str(!qty), 10, " ", 1) + "       " + !deptName
   'Print #1, "     " + PadWithChar(!qty * !salerate * (1 - !DISCRATE / 100), 10, " ", 1)
   'TotAmt = TotAmt + !qty * !salerate * (1 - !DISCRATE / 100)
   JQty = JQty + !qty
   .MoveNext
Loop
End With
Print #1, String(70, "-")
Print #1, "Total Qty=" + Str(JQty)
Set jrec = Nothing
Close #1
frmRpt.rtb.FileName = pfile
frmRpt.Show 1
End Sub

Public Sub blanflds()
   txtDesc = ""
   'txtRate(0) = ""
   txtRate(1) = ""
   txtRate(2) = ""
   txtQty = ""
   txtValue = ""
   LBLRATE = ""
   Label2(0) = ""
   Label2(1) = ""
   txtOp = 0
   'CboUnit = ""
End Sub
Private Sub CboFG_KeyPress(KeyAscii As Integer)
If Len(CboFG) >= 30 Then
   Beep
   Beep
   KeyAscii = 0
End If
End Sub




Private Sub cboCatGrp_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub cboCatGrp_Validate(Index As Integer, Cancel As Boolean)

'Datcategory.Requery
'datGroup.Requery

If Index = 0 Then
   'Data1(1).Recordset.FindFirst "category='" & cboCatGrp(0).BoundText & "'"
  ' RS.Open "SELECT * FROM CATEGORYFILE WHERE CATEGORY='" & Trim(cboCatGrp(0).BoundText) & "'", mhvdb
   Datcategory.Find "category='" & cboCatGrp(0).BoundText & "'", , adSearchForward, 1
   
   'datInvItem.Find ()
   
   If Datcategory.EOF = True Then
      MsgBox "Wrong Category"
      cboCatGrp(0) = ""
      Cancel = True
      Exit Sub
   End If
   If cboCatGrp(Index).BoundText = "02" Then
      TxtFuel.Enabled = True
   Else
      TxtFuel.Enabled = False
      TxtFuel = 0
   End If
Else
   'Data1(2).Recordset.FindFirst "grp='" & cboCatGrp(1).BoundText & "'"
   datGroup.Find "grp='" & cboCatGrp(1).BoundText & "'", , adSearchForward, 1
    'RS.Open "SELECT * FROM GRP WHERE GRP='" & cboCatGrp(1).BoundText & "'", mhvdb, adOpenForwardOnly, adLockReadOnly
   If datGroup.EOF = True Then
      MsgBox "Wrong Group"
      cboCatGrp(1) = ""
      Cancel = True
      Exit Sub
   End If
End If
End Sub

Private Sub cboItemCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub cboItemCode_LostFocus()
Dim imast As New ADODB.Recordset
Dim Issue, Recv As Double
If Len(Trim(cboItemCode)) = 0 Then
   'CboItemCode.SetFocus
   Exit Sub
End If
If Len(Trim(cboItemCode)) > 5 Then
   Beep
   Beep
   MsgBox "Max length of itemcode = 5 "
   cboItemCode.SetFocus
   Exit Sub
End If
'On Error GoTo JR
'MHVDB.BeginTrans
cboItemCode = UCase(cboItemCode)
Set ItemMast = MHVDB.Execute("select * from invitems where itemcode = '" & cboItemCode & "'")

If ItemMast.EOF Then
   If Operation = "Open" Then
      MsgBox cboItemCode + " Does not exists !!!"
      cboItemCode.SetFocus
      Exit Sub
   End If
Else
   If Operation = "ADD" Then
      If MsgBox(cboItemCode + " Already  exists ! Do you want to edit ?", vbYesNo) = vbYes Then
         Operation = "Open"
      Else
         cboItemCode.SetFocus
         Exit Sub
      End If
   End If
   With ItemMast
   txtDesc = !ITEMNAME
  'cboCatGrp(0) = !category
   'cboCatGrp(1) = IIf(IsNull(!gRP), "", !gRP)
    Datcategory.Find "category='" & !Category & "'", , adSearchForward, 1
   If Not Datcategory.EOF Then cboCatGrp(0).Text = Datcategory!Description
   
  
   datGroup.Find "grp='" & !grp & "'", , adSearchForward, 1
   If Not datGroup.EOF Then cboCatGrp(1).Text = datGroup!Description
   
   
   CboUnit(0) = IIf(IsNull(!unit), "", !unit)
   CboUnit(1) = IIf(IsNull(!purunit), !unit, !purunit)
   txtRate(1) = IIf(IsNull(!reorderlevel), 0, !reorderlevel)
   txtRate(2) = IIf(IsNull(!minstock), 0, !minstock)
   txtRate(3) = IIf(IsNull(!convfactor), 1, !convfactor)
   TxtSpec = IIf(IsNull(!spec), "", !spec)
   TxtFuel = IIf(IsNull(!fuelunit), 0, !fuelunit)
   If !Category = "02" Then
      TxtFuel.Enabled = True
   Else
      TxtFuel.Enabled = False
   End If
   
   LBLRATE = Round(IIf(IsNull(!avgstockRate), 0, !avgstockRate), 4)
   End With
   
   Set imast = MHVDB.Execute("select oprate,opbal,totpur,totsale,totadj from itemstock where procyear='" & SysYear & "' and itemcode= '" & cboItemCode.BoundText & "'")
   With imast
   If Not .EOF Then
      txtOp = Round(IIf(IsNull(!opbal), 0, !opbal), 2)
      txtOpRate = Round(IIf(IsNull(!oPrATE), 0, !oPrATE), 4)
      OpVal = Round(Val(txtOp) * Val(txtOpRate), 2)
      Label2(0).Caption = "Pur. Qty = " + Trim(Str(IIf(IsNull(!totpur), 0, !totpur)))
      Label2(1).Caption = "Issue Qty = " + Trim(Str(IIf(IsNull(!TOTSALE), 0, !TOTSALE)))
      Label2(2).Caption = "Adj. Qty = " + Trim(Str(IIf(IsNull(!totadj), 0, !totadj)))
      txtQty.Caption = Round(!opbal + !totpur + !totadj - !TOTSALE, 2)
      txtValue.Caption = Round((!opbal + !totpur + !totadj - !TOTSALE) * Val(LBLRATE), 2)
   Else
      txtOp = 0
      txtOpRate = 0
      OpVal = 0
      
      SQLSTR = "insert into itemstock (Procyear,Itemcode,opbal,totpur,totsale,totadj,oprate) values " _
           & "('" & SysYear & "','" & cboItemCode & "',0,0,0,0,0)"
      MHVDB.Execute SQLSTR, dbFailOnError
   End If
   End With
End If
cboItemCode.Enabled = False
'mnuSave.Enabled = True
TB.Buttons(4).Enabled = True
If Operation = "OPEN" Then
  ' mnuDelete.Enabled = True
   'mnuPH.Enabled = True
   'mnuIH.Enabled = True
   TB.Buttons(3).Enabled = True
   'TB.Buttons(5).Enabled = True
   'TB.Buttons(6).Enabled = True
End If
'MHVDB.CommitTrans
Exit Sub
'JR:
'MsgBox ERR.Description
''MHVDB.Rollback
'ERR.Clear
End Sub

Private Sub CboUnit_Change(Index As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
   
End Sub

Private Sub CboUnit_LostFocus(Index As Integer)
If Index = 0 Then CboUnit(1) = CboUnit(0)
End Sub

Private Sub DataCombo2_Click(Area As Integer)

End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
If datInvItem.State = adStateOpen Then datInvItem.Close
'datInvItem.Open "select a.itemcode,itemname+' '+specs as nm from invitems  as a,itemstock as c where a.itemcode=c.itemcode and c.procyear='" & SysYear & "' and c.unitcode='" & JunitCode & "' order by itemname,specs", db
If OrderBySpec Then
   datInvItem.Open "select itemcode,itemname  from invitems order by Specs,itemname", db, adOpenForwardOnly, adLockReadOnly
Else
   datInvItem.Open "select itemcode,itemname  from invitems order by itemname", db, adOpenForwardOnly, adLockReadOnly
End If
Set cboItemCode.RowSource = datInvItem
cboItemCode.ListField = "ItemCode"
cboItemCode.BoundColumn = "ItemCode"

Set txtDesc.RowSource = datInvItem

txtDesc.ListField = "itemname"
txtDesc.BoundColumn = "ItemCode"
'category
If Datcategory.State = adStateOpen Then Datcategory.Close
Datcategory.Open "select * from categoryfile", db, adOpenForwardOnly, adLockReadOnly
Set cboCatGrp(0).RowSource = Datcategory
cboCatGrp(0).ListField = "description"
cboCatGrp(0).BoundColumn = "category"
'group
If datGroup.State = adStateOpen Then datGroup.Close
datGroup.Open "select * from Grp", db
Set cboCatGrp(1).RowSource = datGroup
cboCatGrp(1).ListField = "description"
cboCatGrp(1).BoundColumn = "Grp"

If Not SuperUser Then LBLRATE.Enabled = False
End Sub
Private Sub PHist()
Dim pfile, Nm As String
Dim jrec, JSUP As ADODB.Recordset
Dim JRate, AvRate, JQty, TotAmt, TotFr, TotPh, totbst, totulc, totmi As Double
pfile = App.Path + "\pHIST.TXT"
Open pfile For Output As #1
SQLSTR = "SELECT a.billdate,a.billtype,a.challanno,A.challanDATE,A.BILLNO,a.suplcode,b.amt,B.QTY,b.rate,freight,bst,ulc,misc,phcharge,lrate FROM tranhdr as A,TRANFILE AS B " _
& " where A.PROCYEAR='" & SysYear & "' AND (A.BILLTYPE='EN' or A.BILLTYPE='SP' )AND A.PROCYEAR=B.PROCYEAR AND A.BILLTYPE=B.BILLTYPE AND A.BILLNO=B.BILLNO AND a.status<>'C' and b.itemcode='" & cboItemCode & "' ORDER BY a.BILLDATE"
Set jrec = MHVDB.Execute(SQLSTR, dbOpenSnapshot)
JQty = 0
 TotAmt = 0
   TotFr = 0 'TotFr + !freight
   TotPh = 0 'TotPh + !phcharge
   totbst = 0 ' Totbst + !bst
   totulc = 0 'Totulc + !ulc
   totmi = 0 'Totmi + !misc
Print #1, Chr(14) + JsysName + Chr(20)
Print #1, "Purchase history of " + txtDesc + "( " + cboItemCode + " )"
Print #1, Chr(15) + String(182, "-")
Print #1, "Challan No &     Date             EntryNo        Qty   Rate      Amount|  F/P Charge|  Adjust|  Freight|    BST  |  Un Load|  L.Rate SuplCode & Name    "
Print #1, String(182, "-")

With jrec
Do While Not .EOF
   Print #1, PadWithChar(IIf(IsNull(!challanno), "", Trim(!challanno)), 17, " ", 0) + PadWithChar(Format(IIf(IsNull(!challandate), !billdate, !challandate), "dd/mm/yyyy"), 10, " ", 0) + " " + PadWithChar(!billtype + "\" + SysYear + "\" + Str(!billno), 13, " ", 1) + " " + PadWithChar(Str(!qty), 10, " ", 1) + PadWithChar(Round(!Rate, 2), 7, " ", 1) + " " + PadWithChar(Format(Round(!amt, 2), "########0.00"), 12, " ", 1) + " ";
   Print #1, PadWithChar(Format(Round(!phcharge, 2), "#######0.00"), 11, " ", 1) + PadWithChar(Format(Round(!misc, 2), "#######0.00"), 10, " ", 1);
   Print #1, PadWithChar(Format(Round(!freight, 2), "######0.00"), 10, " ", 1) + PadWithChar(Format(Round(!bst, 2), "######0.00"), 10, " ", 1);
   Print #1, PadWithChar(Format(Round(!ulc, 2), "#######0.00"), 10, " ", 1) + " " + PadWithChar(Format(Round(!lRate, 2), "####0.00"), 8, " ", 1);
   Nm = ""
   If Not IsNull(!suplcode) Then
      Set srec = MHVDB.Execute("SELECT NAME FROM SUPPLIER WHERE SUPLCODE='" & !suplcode & "'", dbOpenSnapshot)
      If Not srec.EOF Then Nm = srec!Name
      Print #1, "  " + Nm
   Else
      Print #1, IIf(!billtype = "SP", "  Production", "")
   End If
   TotAmt = TotAmt + !amt
   TotFr = TotFr + !freight
   TotPh = TotPh + !phcharge
   totbst = totbst + !bst
   totulc = totulc + !ulc
   totmi = totmi + !misc
   Totlcost = Totlcost + Round(!amt + !freight + !phcharge + !bst + !ulc + !misc, 2)
   JQty = JQty + !qty
   .MoveNext
Loop
End With
Print #1, String(182, "-")
If JQty <> 0 Then
   JRate = Round(Totlcost / JQty, 2)
   AvRate = Round(TotAmt / JQty, 2)
Else
   JRate = 0
   AvRate = 0
End If

Print #1, PadWithChar("Total", 41, " ", 0) + " " + PadWithChar(Str(JQty), 10, " ", 1) + PadWithChar(Round(AvRate, 2), 7, " ", 1) + " " + PadWithChar(Format(Round(TotAmt, 2), "########0.00"), 12, " ", 1) + " ";
   Print #1, PadWithChar(Format(Round(TotPh, 2), "#######0.00"), 11, " ", 1) + PadWithChar(Format(Round(totmi, 2), "#######0.00"), 10, " ", 1);
   Print #1, PadWithChar(Format(Round(TotFr, 2), "######0.00"), 10, " ", 1) + PadWithChar(Format(Round(totbst, 2), "######0.00"), 10, " ", 1);
   Print #1, PadWithChar(Format(Round(totulc, 2), "#######0.00"), 10, " ", 1) + " " + PadWithChar(Format(Round(JRate, 2), "####0.00"), 8, " ", 1) + "  Total Landing Cost " + Str(Totlcost)


'Print #1, "Total Qty=" + Str(JQty) + "   Amount = " + Str(TotAmt) + " Total Landing Cost " + Str(totlcost)
Set jrec = Nothing
Print #1, Chr(18)
Close #1
frmRpt.rtb.FileName = pfile
frmRpt.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub mnuadd_Click()
 ErrCTR = 0
 Operation = "ADD"
 blanflds
 cboItemCode.Enabled = True
 cboItemCode.SetFocus
 'mnuSave.Enabled = False
 TB.Buttons(3).Enabled = True
 'mnuDelete.Enabled = False
 'TB.Buttons(4).Enabled = False
' TB.Buttons(5).Enabled = False
'TB.Buttons(6).Enabled = False
End Sub
Private Sub mnuexit_Click()
 Screen.MousePointer = vbDefault
 Unload Me
End Sub
Private Sub mnuDelete_Click()
Dim Billrec As ADODB.Recordset
ErrCTR = 0
If MsgBox("Delete it !!! Are u Sure ?", vbYesNo) = vbNo Then Exit Sub
On Error GoTo ERR
Set Billrec = MHVDB.Execute("Select a.billno from tranfile as a , tranhdr as b where a.billno=b.billno and b.status<>'C' and a.itemcode='" & cboItemCode & "'", dbOpenDynaset)
If Not Billrec.EOF Then
   Beep
   MsgBox "You Cant Delete the ItemCode. Stock Entry is there against it Billno " + Str(Billrec!billno)
   Exit Sub
End If
MHVDB.Execute "delete from INVITEMS where itemcode='" & cboItemCode & "'", dbSeeChanges + dbFailOnError
MHVDB.Execute "delete from ITEMStock where itemcode='" & cboItemCode & "'", dbSeeChanges + dbFailOnError
blanflds
cboItemCode.Refresh
cboItemCode.Enabled = False


TB.Buttons(4).Enabled = False
TB.Buttons(3).Enabled = False
Exit Sub
ERR:
MsgBox "error: Operation Cancelled"
ERR.Clear
End Sub

Private Sub mnuIH_Click()
IHIST
End Sub

Private Sub mnuOpen_Click()
ErrCTR = 0
Operation = "OPEN"
blanflds
cboItemCode.Enabled = True
cboItemCode.SetFocus
TB.Buttons(3).Enabled = True

End Sub

Private Sub mnuPH_Click()
PHist
If False Then
Set lastbill = MHVDB.Execute("select itemcode from invitems ", dbOpenDynaset)
With lastbill
Do While Not .EOF
   'For i = 1 To 12
   MHVDB.Execute "insert into itemstock (itemcode,procyear) values ('" & !itemcode & "','" & SysYear & "')"   'Next
   .MoveNext
Loop
End With
End If
End Sub

Private Sub mnuSave_Click()
'Datcategory.Requery
'datGroup.Requery
'datInvItem.Requery
Dim SQLSTR, Cat As String
0:
On Error GoTo ERR
 If Len(cboCatGrp(0).BoundText) = 0 Then
    MsgBox "Wrong Category"
    cboCatGrp(0).SetFocus
    Exit Sub
 End If
 If Len(cboCatGrp(1).BoundText) = 0 Then
    MsgBox "Wrong Group"
    cboCatGrp(1).SetFocus
    Exit Sub
     End If
 Cat = cboCatGrp(0).BoundText
 
MHVDB.BeginTrans
 
 If Operation = "ADD" Then
    SQLSTR = "INSERT INTO INVITEMS (ITEMCODE,ITEMNAME,UNIT,PURUNIT,REORDERLEVEL,MINSTOCK,CONVFACTOR,CATEGORY,GRP,AVGSTOCKRATE,spec,fuelunit) VALUES" _
           & "('" & cboItemCode & "','" & txtDesc & "','" & CboUnit(0) & "','" & CboUnit(1) & "',('" & txtRate(1) & "'),('" & txtRate(2) & "'),('" & txtRate(3) & "')" _
           & " ,'" & cboCatGrp(0).BoundText & "','" & cboCatGrp(1).BoundText & "',('" & LBLRATE & "'),'" & TxtSpec & "',('" & TxtFuel & "'))"
    MHVDB.Execute SQLSTR
    SQLSTR = "insert into itemstock (Procyear,Itemcode,opbal,totpur,totsale,totadj,oprate) values " _
           & "('" & SysYear & "','" & cboItemCode & "',('" & txtOp & "'),0,0,0,('" & txtOpRate & "'))"
    MHVDB.Execute SQLSTR
 ElseIf Operation = "OPEN" Then
     SQLSTR = "UPDATE INVITEMS SET itemNAME = '" & txtDesc & "',spec = '" & TxtSpec & "',unit = '" & CboUnit(0) & "',purunit = '" & CboUnit(1) & "', " _
            & " reORDERlevel = ('" & txtRate(1) & "'),minstock = ('" & txtRate(2) & "'),fuelunit = ('" & TxtFuel & "'), " _
            & " convfactor = ('" & txtRate(3) & "'),category = '" & Cat & "',Grp ='" & cboCatGrp(1).BoundText & "', " _
            & " avgstockrate = ('" & LBLRATE & "') WHERE ITEMCODE = '" & cboItemCode & "'"
     MHVDB.Execute SQLSTR
     MHVDB.Execute "Update itemstock set opbal=('" & txtOp & "'),oprate=('" & txtOpRate & "') where procyear='" & SysYear & "' and ITEMCODE = '" & cboItemCode & "'", dbFailOnError
 Else
    Beep
    Beep
    Exit Sub
 End If
 MHVDB.CommitTrans
 Operation = ""
 Datcategory.Requery
 cboItemCode.Refresh
 TB.Buttons(4).Enabled = False
 TB.Buttons(3).Enabled = False
 ErrCTR = 0
 Exit Sub
ERR:
MsgBox ERR.Description
ErrCTR = ErrCTR + 1
If ErrCTR > 5 Then
   If DBEngine.Errors.Count > 0 Then
      For Each errLoop In DBEngine.Errors
          MsgBox "Error number: " & errLoop.Number & vbCr & _
          errLoop.Description
      Next errLoop

   End If
End If

MHVDB.RollbackTrans
If ErrCTR < 6 Then
   For i = 1 To 1000
       For j = 1 To 9999
       Next
   Next
   GoTo 0
Else
  MsgBox "Error ! Cant Update"
End If
End Sub






Private Sub Tb_ButtonClick(ByVal Button As msComctlLib.Button)
Select Case Button.Key
       Case "ADD"
           mnuadd_Click
       Case "OPEN"
           mnuOpen_Click
       Case "SAVE"
           mnuSave_Click
       Case "DELETE"
          ' mnuDelete_Click
       Case "PH"
           PHist
       Case "IH"
           IHIST
       Case "EXIT"
           Unload Me
End Select
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub



Private Sub txtdesc_Validate(Cancel As Boolean)
If Len(txtDesc) = 0 Then Cancel = True
Set ItemMast = Nothing
'Set ItemMast = mhvdb.EXECUTE("select * from invitems where itemname = '" & txtDesc & "'", dbOpenDynaset, dbReadOnly)
ItemMast.Open "select * from invitems where itemname = '" & txtDesc & "'", MHVDB, adOpenForwardOnly, adLockReadOnly
If ItemMast.EOF Then
   If Operation = "Open" Then
      MsgBox txtDesc + " Does not exists !!!"
      Cancel = True
      Exit Sub
   End If
Else
   If Operation = "Add" Then
      If MsgBox(txtDesc + " Already  exists ! Do you want to edit ?", vbYesNo) = vbYes Then
         Operation = "Open"
      Else
         Cancel = True
         Exit Sub
      End If
   End If
   With ItemMast
   cboItemCode = !itemcode
   txtDesc = !ITEMNAME
   
   Datcategory.Find "category='" & !Category & "'", , adSearchForward, 1
   If Not Datcategory.EOF Then cboCatGrp(0).Text = Datcategory!Description
   
  
   datGroup.Find "grp='" & !grp & "'", , adSearchForward, 1
   If Not datGroup.EOF Then cboCatGrp(1).Text = datGroup!Description
  
  ' cboCatGrp(1).Text = IIf(IsNull(!gRP), "", !gRP)
   CboUnit(0).Text = IIf(IsNull(!unit), "", !unit)
   CboUnit(1) = IIf(IsNull(!purunit), !unit, !purunit)
   txtRate(1) = IIf(IsNull(!reorderlevel), 0, !reorderlevel)
   txtRate(2) = IIf(IsNull(!minstock), 0, !minstock)
   txtRate(3) = IIf(IsNull(!convfactor), 1, !convfactor)
   LBLRATE = Round(IIf(IsNull(!avgstockRate), 0, !avgstockRate), 2)
   End With
   Set imast = MHVDB.Execute("select oprate,opbal,totpur,totsale,totadj from itemstock where procyear='" & SysYear & "' and itemcode= '" & cboItemCode & "'")
   With imast
   If Not .EOF Then
      txtOp = Round(IIf(IsNull(!opbal), 0, !opbal), 2)
      txtOpRate = Round(IIf(IsNull(!oPrATE), 0, !oPrATE), 2)
      OpVal = Round(Val(txtOp) * Val(txtOpRate), 2)
      Label2(0).Caption = "Pur. Qty = " + Trim(Str(IIf(IsNull(!totpur), 0, !totpur)))
      Label2(1).Caption = "Issue Qty = " + Trim(Str(IIf(IsNull(!TOTSALE), 0, !TOTSALE)))
      Label2(2).Caption = "Adj. Qty = " + Trim(Str(IIf(IsNull(!totadj), 0, !totadj)))
      txtQty.Caption = Round(!opbal + !totpur + !totadj - !TOTSALE, 2)
      txtValue.Caption = Round((!opbal + !totpur + !totadj - !TOTSALE) * Val(LBLRATE), 2)
   Else
      txtOp = 0
      txtOpRate = 0
      OpVal = 0
      
      SQLSTR = "insert into itemstock (Procyear,Itemcode,opbal,totpur,totsale,totadj,oprate) values " _
           & "('" & SysYear & "','" & cboItemCode & "',0,0,0,0,0)"
      MHVDB.Execute SQLSTR, dbFailOnError
   End If
   End With
End If
cboItemCode.Enabled = False
'mnuSave.Enabled = True
TB.Buttons(4).Enabled = True
If Operation = "Open" Then
   'mnuDelete.Enabled = True
  ' mnuPH.Enabled = True
   'mnuIH.Enabled = True
   TB.Buttons(3).Enabled = True
   'TB.Buttons(5).Enabled = True
   'TB.Buttons(6).Enabled = True
End If
End Sub

Private Sub txtOp_Validate(Cancel As Boolean)
OpVal = Val(txtOp) * Val(txtOpRate)
End Sub

Private Sub txtOpRate_Validate(Cancel As Boolean)
OpVal = Val(txtOp) * Val(txtOpRate)
End Sub

Private Sub txtRate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

