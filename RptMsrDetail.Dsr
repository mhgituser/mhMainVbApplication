VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptMsrDetail 
   Caption         =   "REPORT"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   Icon            =   "RptMsrDetail.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "RptMsrDetail.dsx":0E42
End
Attribute VB_Name = "RptMsrDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
FindFA txtFARMERID.Text, "F"
TXTFARMERNAME.Text = FAName
End Sub

Private Sub GroupHeader1_Format()
FindsTAFF txtMONITOR.Text
TXTSTAFFNAME.Text = sTAFF
End Sub

Private Sub PageHeader_Format()
Select Case RptOption
Case "MSR"
Field20.Text = "DETAIL INFORMATION ON  MONITOR (SHARED REGISTRATION)"
Case 2

End Select



End Sub
