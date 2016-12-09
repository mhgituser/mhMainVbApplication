VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmdashbordtrn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard File Maintainance......"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6030
   Icon            =   "frmdashbordtrn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   5775
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   2775
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5535
         _cx             =   9763
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
         BackColor       =   12632256
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12632256
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
         FormatString    =   $"frmdashbordtrn.frx":0E42
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
   Begin MSComDlg.CommonDialog CD 
      Left            =   7320
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame itemcode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   5415
      Begin VB.TextBox txtsheetname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cmdview 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         Picture         =   "frmdashbordtrn.frx":0EDF
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtfilesize 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtfilename 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   4680
         Picture         =   "frmdashbordtrn.frx":1689
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Remove File"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   3960
         Picture         =   "frmdashbordtrn.frx":1A13
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Browse File..."
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtfilepath 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "frmdashbordtrn.frx":1D9D
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cbodept 
         Bindings        =   "frmdashbordtrn.frx":1DB2
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Department"
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
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sheet Name"
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
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "File"
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
         Left            =   1080
         TabIndex        =   5
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transaction ID"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1275
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   2640
      Top             =   3240
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
            Picture         =   "frmdashbordtrn.frx":1DC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdashbordtrn.frx":2161
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdashbordtrn.frx":24FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdashbordtrn.frx":31D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdashbordtrn.frx":3627
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdashbordtrn.frx":3DE1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6030
      _ExtentX        =   10636
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
Attribute VB_Name = "frmdashbordtrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim picfile As String

Private Sub cbotrnid_LostFocus()
If Len(cbotrnid.Text) = 0 Then Exit Sub
Dim rs As New ADODB.Recordset
Dim mystream As ADODB.Stream
Set rs = Nothing




'rs.Open "select * from tblfarmer where idfarmer='" & cbofarmerid.BoundText & "'", MHVDB, adOpenStatic, adLockOptimistic
'Set mystream = New ADODB.Stream
'mystream.Type = adTypeBinary
'mystream.Open
'mystream.Write rs!picfile
'mystream.SaveToFile "c:\\" & cbofarmerid.BoundText & ".jpg", adSaveCreateOverWrite
'mystream.Close
'ImgPic.Picture = LoadPicture("c:\\" & cbofarmerid.BoundText & ".jpg")
'ImgPic.Width = 2000
'ImgPic.Height = 2000
'
' If Dir$("c:\\" & cbofarmerid.BoundText & ".jpg") <> vbNullString Then
'    Kill "c:\\" & cbofarmerid.BoundText & ".jpg"
'    End If




rs.Open "select * from tbldashbordtrn where trnid='" & cbotrnid.BoundText & "'", MHVDB, adOpenStatic, adLockOptimistic
If rs.EOF <> True Then
Finddept rs!Dept
cbodept.Text = deptName
txtfilename.Text = rs!FileName
txtfilepath.Text = ""
txtfilesize.Text = rs!fileSize
txtsheetname.Text = rs!sheetname
End If
cmdview.Visible = True
cmdview.Enabled = True
TB.Buttons(3).Enabled = True
End Sub

Private Sub cmdview_Click()
If Len(cbotrnid.Text) = 0 Then Exit Sub
Dim rs As New ADODB.Recordset
Dim mystream As ADODB.Stream
Set rs = Nothing
rs.Open "select * from tbldashbordtrn where trnid='" & cbotrnid.BoundText & "'", MHVDB, adOpenStatic, adLockOptimistic
If rs.EOF <> True Then
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
mystream.Open
mystream.Write rs!mfile
mystream.SaveToFile "c:\\" & rs!FileName, adSaveCreateOverWrite
mystream.Close
End If
showxlfile "c:\\" & rs!FileName
    




End Sub
Private Sub showxlfile(FileName As String)
Dim xl As Excel.Application
    Dim var As Variant
    Set xl = CreateObject("excel.Application")
           
     If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + txtfilename.Text) <> vbNullString Then
        Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + txtfilename.Text
     End If
           
    FileCopy FileName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + txtfilename.Text
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + txtfilename.Text
     If Dir$(FileName) <> vbNullString Then
        Kill FileName
     End If
            
            
            
    xl.Visible = True
   Set xl = Nothing
Screen.MousePointer = vbDefault
End Sub


Private Sub Command1_Click()
On Error GoTo ErrHandler
 Dim MySize As Long
picfile = ""
    CD.CancelError = True
    CD.InitDir = "C:\"
    'CD.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Batch Files (*.bat)|*.bat"
    CD.Filter = "All Files (*.*)|*.*"
    CD.ShowOpen
    
   
    picfile = CD.FileName
    txtfilepath.Text = CD.FileName
    
   
 fsize = ""
If Dir$(txtfilepath.Text) <> vbNullString Then
   MySize = FileLen(txtfilepath.Text)
Else
   MsgBox "file doesn't exist"
End If
txtfilesize.Text = Round((MySize / 1024) / 1025, 2)
txtfilename.Text = LCase(CD.FileTitle)
    
Exit Sub

ErrHandler:
'    User pressed Cancel button.
   Exit Sub
End Sub

Private Sub Command2_Click()
txtfilepath.Text = ""
txtfilename.Text = ""
txtfilesize.Text = ""
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Dim tt As String
Dim rs As New ADODB.Recordset
db.CursorLocation = adUseClient
db.Open CnnString
Set rs = Nothing

Set rs = Nothing
rs.Open "select trnid from tbldashbordtrn" _
& "  order by trnid", db

Set cbotrnid.RowSource = rs
cbotrnid.ListField = "trnid"
cbotrnid.BoundColumn = "trnid"

Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
If UCase(MUSER) = "ADMIN" Then
rs.Open "select deptid,deptname from tbldept order by deptid", db
Else

 rs.Open "select deptid,deptname from tbldept where remarks like  " & "'%" & UserId & "%'" & "  order by deptid", db

End If
Set cbodept.RowSource = rs
cbodept.ListField = "deptname"
cbodept.BoundColumn = "deptid"
FillGrid
End Sub
Private Sub FillGrid()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^SL.NO.|^TRN. ID|^DDEPARTMENT|^SHEET NAME|"
Mygrid.ColWidth(0) = 750
Mygrid.ColWidth(1) = 900
Mygrid.ColWidth(2) = 1515
Mygrid.ColWidth(3) = 2025

rs.Open "select * from tbldashbordtrn  order by trnid", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i
Finddept rs!Dept
Mygrid.TextMatrix(i, 1) = rs!trnid
Mygrid.TextMatrix(i, 2) = deptName
Mygrid.TextMatrix(i, 3) = rs!sheetname

rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       cbotrnid.Enabled = False
        TB.Buttons(3).Enabled = True
       operation = "ADD"
       CLEARCONTROLL
       Dim rs As New ADODB.Recordset
       Set rs = Nothing
       rs.Open "SELECT MAX(trnid)+1 AS MaxID from tbldashbordtrn", MHVDB, adOpenForwardOnly, adLockOptimistic
       If rs.EOF <> True Then
       
       
        cbotrnid.Text = IIf(IsNull(rs!MaxId), 1, rs!MaxId)
        
       Else
       cbotrnid.Text = 1
       End If
       
       
       Case "OPEN"
       operation = "OPEN"
       CLEARCONTROLL
       cbotrnid.Enabled = True
       
       
       Case "SAVE"
       MNU_SAVE
        FillGrid
     
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub MNU_SAVE()
Dim rs As New ADODB.Recordset
Dim mystream As ADODB.Stream
On Error GoTo err
If Len(cbotrnid.Text) = 0 Or Len(cbodept.Text) = 0 Or Len(txtsheetname.Text) = 0 Then
MsgBox "Invalid Selection."
Exit Sub
End If


MHVDB.BeginTrans
If operation = "ADD" Then
If Len(txtfilepath.Text) = 0 Or Len(txtfilename.Text) = 0 Then
MsgBox "Please Select The file to save."
Exit Sub
End If

Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
Set rs = Nothing
rs.Open "select * from tbldashbordtrn where 1", MHVDB, adOpenStatic, adLockOptimistic
rs.AddNew
rs!trnid = cbotrnid.Text
rs!Dept = cbodept.BoundText
rs!sheetname = txtsheetname.Text
rs!FileName = txtfilename.Text
rs!fileSize = txtfilesize.Text
rs!UserId = MUSER
mystream.Open
mystream.LoadFromFile txtfilepath
rs!mfile = mystream.Read
rs.Update


ElseIf operation = "OPEN" Then
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
Set rs = Nothing
rs.Open "select * from tbldashbordtrn where trnid = '" & cbotrnid.Text & "'", MHVDB, adOpenStatic, adLockOptimistic


rs!Dept = cbodept.BoundText
rs!sheetname = txtsheetname.Text
rs!FileName = txtfilename.Text
If Len(txtfilepath.Text) > 0 Then
rs!fileSize = txtfilesize.Text
rs!UserId = MUSER
mystream.Open
mystream.LoadFromFile txtfilepath
rs!mfile = mystream.Read
End If

rs.Update

Else
MsgBox "OPERATION NOT SELECTED."
End If
TB.Buttons(3).Enabled = False
MHVDB.CommitTrans

Exit Sub

err:
MsgBox err.Description
MHVDB.RollbackTrans


End Sub
Private Sub CLEARCONTROLL()
cbodept.Text = ""
txtsheetname.Text = ""
txtfilepath.Text = ""
txtfilename.Text = ""
txtfilesize.Text = ""
cmdview.Enabled = False



End Sub

Private Sub txtsheetname_KeyPress(KeyAscii As Integer)
If InStr(1, "abcdefghijklmnopqrstuvwxyz", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
