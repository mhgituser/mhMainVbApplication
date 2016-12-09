VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRptPreview 
   Caption         =   " Report Printing . . . "
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   Icon            =   "frmRptPreview.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   5220
   ScaleWidth      =   7290
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2460
      Top             =   2310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "WinTSoft : Export Report To File..."
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer21 
      Height          =   5085
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8969
      SectionData     =   "frmRptPreview.frx":076A
   End
End
Attribute VB_Name = "frmRptPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ARViewer21_ToolbarClick(ByVal tool As DDActiveReportsViewer2Ctl.DDTool)
    If tool.Caption = "Export To File" Then
       FileTextExport
    End If
End Sub

Private Sub Form_Load()
    ARViewer21.Toolbar.Tools.Add "Export To File"
    ARViewer21.Toolbar.Tools.Add "EMail"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ARViewer21.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Public Sub RunReport(rpt As Object)
    Set ARViewer21.ReportSource = rpt
    ARViewer21.Zoom = 100
End Sub

Private Sub FileTextExport()
'On Error Resume Next
 'Dim txt As New ActiveReportsExcelExport.ARExportExcel

'CommonDialog1.Filter = "PDF Format (*.pdf)|*.pdf"
'CommonDialog1.DefaultExt = ".pdf"
'Dim txt As New ActiveReportsRTFExport.ARExportRTF
'CommonDialog1.Filter = "DOC Format (*.doc)|*.doc"
'CommonDialog1.DefaultExt = ".doc"
'Dim txt As New ActiveReportsExcelExport.ARExportExcel
CommonDialog1.Filter = "Excel Format (*.xls)|*.xls"
CommonDialog1.DefaultExt = ".xls"


CommonDialog1.ShowSave
txt.FileName = CommonDialog1.FileName
If ARViewer21.Pages.Count > 0 Then

   'txt.Export ARViewer21.Pages
      txt.Export ARViewer21.ReportSource.Pages
      'txt.Export 1
ElseIf Not ARViewer21.ReportSource Is Nothing Then
   If ARViewer21.ReportSource.Pages.Count > 0 Then
      txt.Export rpt_StockOverView.Pages   'ARViewer21.ReportSource.Pages
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRptPreview = Nothing
End Sub

