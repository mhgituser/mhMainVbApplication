VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMNEWLANDREG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NEW LAND REGISTRATION  FOR FARMER"
   ClientHeight    =   10260
   ClientLeft      =   4545
   ClientTop       =   765
   ClientWidth     =   13605
   Icon            =   "FRMNEWLANDREG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   13605
   Begin VB.CommandButton Command1 
      Caption         =   "Farmer Reg....."
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
      Left            =   11640
      Picture         =   "FRMNEWLANDREG.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "PERSON REGISTERING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4935
      Left            =   120
      TabIndex        =   23
      Top             =   4560
      Width           =   13455
      Begin VB.CommandButton Command2 
         Caption         =   "Divide Rest"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12000
         TabIndex        =   67
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox TXTSUPPORT5PERCENT 
         Alignment       =   1  'Right Justify
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
         Left            =   12120
         TabIndex        =   66
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TXTSUPPORT4PERCENT 
         Alignment       =   1  'Right Justify
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
         Left            =   12120
         TabIndex        =   63
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox TXTSUPPORT3PERCENT 
         Alignment       =   1  'Right Justify
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
         Left            =   12120
         TabIndex        =   62
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox TXTSUPPORT2PERCENT 
         Alignment       =   1  'Right Justify
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
         Left            =   12120
         TabIndex        =   61
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox TXTSUPPORT1PERCENT 
         Alignment       =   1  'Right Justify
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
         Left            =   12120
         TabIndex        =   60
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox TXTLEADSTAFFPERCENT 
         Alignment       =   1  'Right Justify
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
         Left            =   12120
         TabIndex        =   59
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Frame Frame11 
         Caption         =   "MEETING REGISTRATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3375
         Left            =   6720
         TabIndex        =   45
         Top             =   1200
         Width           =   5295
         Begin MSDataListLib.DataCombo CBOLEADSTAFF 
            Bindings        =   "FRMNEWLANDREG.frx":14F4
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   1800
            TabIndex        =   46
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSDataListLib.DataCombo CBOSUPPORT1 
            Bindings        =   "FRMNEWLANDREG.frx":1509
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   1800
            TabIndex        =   47
            Top             =   720
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSDataListLib.DataCombo CBOSUPPORT2 
            Bindings        =   "FRMNEWLANDREG.frx":151E
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   1800
            TabIndex        =   50
            Top             =   1200
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSDataListLib.DataCombo CBOSUPPORT3 
            Bindings        =   "FRMNEWLANDREG.frx":1533
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   1800
            TabIndex        =   51
            Top             =   1680
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSDataListLib.DataCombo CBOSUPPORT4 
            Bindings        =   "FRMNEWLANDREG.frx":1548
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   1800
            TabIndex        =   52
            Top             =   2160
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSDataListLib.DataCombo CBOSUPPORT5 
            Bindings        =   "FRMNEWLANDREG.frx":155D
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   1800
            TabIndex        =   64
            Top             =   2640
            Width           =   3135
            _ExtentX        =   5530
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
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "SUPPORT STAFF5"
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
            TabIndex        =   65
            Top             =   2760
            Width           =   1635
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "SUPPORT STAFF4"
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
            TabIndex        =   55
            Top             =   2280
            Width           =   1635
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "SUPPORT STAFF3"
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
            TabIndex        =   54
            Top             =   1800
            Width           =   1635
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "SUPPORT STAFF2"
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
            TabIndex        =   53
            Top             =   1320
            Width           =   1635
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "SUPPORT STAFF1"
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
            TabIndex        =   49
            Top             =   840
            Width           =   1635
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "LEAD STAFF"
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
            TabIndex        =   48
            Top             =   360
            Width           =   1125
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "OTHERS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   3480
         Width           =   1455
         Begin VB.TextBox TXTCIDOTHER 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   57
            Top             =   720
            Width           =   5175
         End
         Begin VB.TextBox TXTOTHER 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   56
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "CID"
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
            TabIndex        =   44
            Top             =   840
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "NAME"
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
            TabIndex        =   43
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "INDIVIDUAL REGISTRATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   6720
         TabIndex        =   40
         Top             =   360
         Width           =   5295
         Begin MSDataListLib.DataCombo cboindividual 
            Bindings        =   "FRMNEWLANDREG.frx":1572
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   4215
            _ExtentX        =   7435
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
      End
      Begin VB.Frame Frame8 
         Caption         =   "SHARED REGISTRATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3240
         TabIndex        =   35
         Top             =   3720
         Width           =   495
         Begin MSDataListLib.DataCombo CBOMONITOR 
            Bindings        =   "FRMNEWLANDREG.frx":1587
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   1800
            TabIndex        =   36
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
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
         Begin MSDataListLib.DataCombo CBOOUTREACH 
            Bindings        =   "FRMNEWLANDREG.frx":159C
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   1800
            TabIndex        =   37
            Top             =   720
            Width           =   4215
            _ExtentX        =   7435
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "MONITOR"
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
            TabIndex        =   39
            Top             =   360
            Width           =   885
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "OUTREACH STAFF"
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
            TabIndex        =   38
            Top             =   720
            Width           =   1665
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "CG REGISTRATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   480
         TabIndex        =   30
         Top             =   2760
         Width           =   615
         Begin MSDataListLib.DataCombo cbocgid 
            Bindings        =   "FRMNEWLANDREG.frx":15B1
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   1800
            TabIndex        =   31
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
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
         Begin MSDataListLib.DataCombo cbomhvstaff 
            Bindings        =   "FRMNEWLANDREG.frx":15C6
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   1800
            TabIndex        =   32
            Top             =   720
            Width           =   4215
            _ExtentX        =   7435
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "MONITOR"
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
            TabIndex        =   34
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "ID OF CG"
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
            TabIndex        =   33
            Top             =   360
            Width           =   825
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "GENERAL INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   13455
      Begin VB.Frame Frame6 
         Caption         =   "SORT BY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   6000
         TabIndex        =   27
         Top             =   600
         Width           =   2295
         Begin VB.OptionButton OPTBYID 
            Caption         =   " ID"
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
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton OPTNAME 
            Caption         =   "NAME"
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
            Left            =   960
            TabIndex        =   28
            Top             =   240
            Width           =   975
         End
      End
      Begin MSDataListLib.DataCombo cbofarmerid 
         Bindings        =   "FRMNEWLANDREG.frx":15DB
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSComCtl2.DTPicker txtregdate 
         Height          =   375
         Left            =   11760
         TabIndex        =   12
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   102367233
         CurrentDate     =   41208
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   ""
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
      Begin MSDataListLib.DataCombo cbosessionid 
         Height          =   360
         Left            =   9720
         TabIndex        =   72
         Top             =   720
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   ""
         BoundColumn     =   ""
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
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "REG. SESSION"
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
         Left            =   8400
         TabIndex        =   71
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TRANSACTION ID"
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
         TabIndex        =   21
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "REG. DATE"
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
         Left            =   8400
         TabIndex        =   11
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FARMER ID"
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
         TabIndex        =   9
         Top             =   840
         Width           =   1035
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "LAND INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   12135
      Begin VB.ComboBox cboplantedstatus 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FRMNEWLANDREG.frx":15F0
         Left            =   9840
         List            =   "FRMNEWLANDREG.frx":15FD
         TabIndex        =   70
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkconfirmed 
         Caption         =   "Is the land confirmed?"
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
         Left            =   3720
         TabIndex        =   68
         ToolTipText     =   "Check if the land is confirmed OR Uncheck if the land is not yet confrmed."
         Top             =   360
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.ComboBox cbolandtype 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FRMNEWLANDREG.frx":162E
         Left            =   7200
         List            =   "FRMNEWLANDREG.frx":1638
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtregland 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "PLANTED STATUS"
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
         Left            =   8160
         TabIndex        =   69
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "LAND TYPE"
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
         Left            =   6120
         TabIndex        =   6
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "REGISTERED LAND ACRE"
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
         TabIndex        =   5
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "REMARKS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   -120
      TabIndex        =   1
      Top             =   9600
      Width           =   11655
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   11415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "THRAM HOLDER INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   12135
      Begin VB.TextBox txtagreement 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   74
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox cbosex 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FRMNEWLANDREG.frx":1644
         Left            =   1680
         List            =   "FRMNEWLANDREG.frx":164E
         TabIndex        =   26
         Top             =   1200
         Width           =   3375
      End
      Begin VB.ComboBox cborelation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FRMNEWLANDREG.frx":1660
         Left            =   8640
         List            =   "FRMNEWLANDREG.frx":169A
         TabIndex        =   24
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtthramholdername 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   19
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtthramno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "AGREEMENTNO."
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
         TabIndex        =   73
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "RELATIONSHIP TO THRAM HOLDER"
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
         Left            =   5280
         TabIndex        =   17
         Top             =   1320
         Width           =   3240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "SEX"
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
         TabIndex        =   16
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "THRAM HOLDER NAME"
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
         Left            =   6480
         TabIndex        =   15
         Top             =   360
         Width           =   2085
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "THRAM NO."
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
         TabIndex        =   13
         Top             =   360
         Width           =   1065
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   2640
      Top             =   2760
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
            Picture         =   "FRMNEWLANDREG.frx":175A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDREG.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDREG.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDREG.frx":2B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDREG.frx":2FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDREG.frx":3774
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   13605
      _ExtentX        =   23998
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
   Begin VB.Label LBLDESC 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   360
      TabIndex        =   22
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "THRAM NO."
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
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "FRMNEWLANDREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsfr As New ADODB.Recordset
Dim RSTR As New ADODB.Recordset
Dim rsCg As New ADODB.Recordset
Dim rsMs As New ADODB.Recordset
Dim rssid As New ADODB.Recordset
Dim rsMoniter As New ADODB.Recordset
Dim rsOutreach As New ADODB.Recordset
Dim rsLeadstaff As New ADODB.Recordset
Dim rssupport1 As New ADODB.Recordset
Dim rssupport2 As New ADODB.Recordset
Dim rssupport3 As New ADODB.Recordset
Dim rssupport4 As New ADODB.Recordset
Dim rssupport5 As New ADODB.Recordset
Dim rsind As New ADODB.Recordset
Dim magno As String
Dim FrName, CGname, MHVName, MHVMONITOR, MHVINDIVIDUAL, MHVLEADSTAFF, SUPPORT1, SUPPORT2, SUPPORT3, SUPPORT4, SUPPORT5, MHVOUTREACH As String

Private Sub cbocgid_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbofarmerid_KeyPress(KeyAscii As Integer)
If InStr(1, "DGTFC0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbofarmerid_LostFocus()
cbofarmerid.Text = UCase(cbofarmerid.Text)
If Mid(UCase(cbofarmerid.Text), 10, 1) = "G" Or Mid(UCase(cbofarmerid.Text), 10, 1) = "C" Then
cboplantedstatus.Text = "Partial Plantation"
Else
cboplantedstatus.Text = "New"
End If

End Sub

Private Sub cboindividual_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbolandtype_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CBOLEADSTAFF_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbomhvstaff_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CBOMONITOR_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CBOOUTREACH_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cboplantedstatus_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cborelation_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbosex_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CBOSUPPORT1_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CBOSUPPORT2_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CBOSUPPORT3_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CBOSUPPORT4_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CBOSUPPORT5_KeyPress(KeyAscii As Integer)
If InStr(1, "~", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Command1_Click()
'LBLDESC.Caption = "THIS DROP DOWN CONTROLL POPULATES THE FARMER ID AND FARMER NAME ,THAT ARE SO FAR REGISTERED."
If TB.Buttons(3).Enabled = True Then
MsgBox "Please Save This Information First."
Exit Sub
Else
Unload Me
frmfarmerreg.Show 1
End If
End Sub

Private Sub Command2_Click()
If IsNumeric(TXTLEADSTAFFPERCENT.Text) Then
TXTSUPPORT1PERCENT.Text = Format((100 - Val(TXTLEADSTAFFPERCENT.Text)) / 5, "###0.00")
TXTSUPPORT2PERCENT.Text = Format((100 - Val(TXTLEADSTAFFPERCENT.Text)) / 5, "###0.00")
TXTSUPPORT3PERCENT.Text = Format((100 - Val(TXTLEADSTAFFPERCENT.Text)) / 5, "###0.00")
TXTSUPPORT4PERCENT.Text = Format((100 - Val(TXTLEADSTAFFPERCENT.Text)) / 5, "###0.00")
TXTSUPPORT5PERCENT.Text = Format((100 - Val(TXTLEADSTAFFPERCENT.Text)) / 5, "###0.00")


Else
TXTSUPPORT1PERCENT.Text = 0
TXTSUPPORT2PERCENT.Text = 0
TXTSUPPORT3PERCENT.Text = 0
TXTSUPPORT4PERCENT.Text = 0
TXTSUPPORT5PERCENT.Text = 0


End If

End Sub



Private Sub cbotrnid_LostFocus()
On Error GoTo err
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tbllandreg where trnid='" & cbotrnid.BoundText & "'", MHVDB
If rs.EOF <> True Then
FindFarmer rs!farmerid


cbofarmerid.Text = rs!farmerid & " " & FrName
txtthramno.Text = IIf(IsNull(rs!thramno), "", rs!thramno)
txtthramholdername.Text = IIf(IsNull(rs!thramname), "", rs!thramname)
txtagreement.Text = IIf(Val(magno) = 0, "", magno)
If rs!sex = 0 Then
cbosex.Text = "Male"
ElseIf rs!sex = 1 Then
cbosex.Text = "Female"
Else
cbosex.Text = ""
End If
'cborelation.Text = cborelation.ItemData(rs!Relation)
cborelation.Text = IIf(IsNull(rs!RELATION), "", rs!RELATION)
txtregland.Text = Format(IIf(IsNull(rs!regland), 0, rs!regland), "####0.00")
txtregdate.Value = Format(rs!regdate, "dd/MM/yyyy")
If rs!islandconfirmed = "Y" Then
chkconfirmed.Value = 1
Else
chkconfirmed.Value = 0
End If
cbolandtype.Text = rs!LANDTYPE
If rs!cgid <> "" Then
FindCG rs!cgid
cbocgid.Text = rs!cgid & " " & CGname
Else
cbocgid.Text = ""
End If
If rs!cgmonitor <> "" Then
FindMHV rs!cgmonitor
cbomhvstaff.Text = rs!cgmonitor & " " & MHVName
End If

If rs!monitor <> "" Then
FindMONITOR rs!monitor
CBOMONITOR.Text = rs!monitor & " " & MHVMONITOR
End If

If rs!outreach <> "" Then
FindOUTREACH rs!outreach
CBOOUTREACH.Text = rs!outreach & " " & MHVOUTREACH
End If

If rs!individual <> "" Then
FindINDIVIDUAL rs!individual
cboindividual.Text = rs!individual & " " & MHVINDIVIDUAL
End If

If rs!LEADSTAFF <> "" Then
FindLEADSTAFF rs!LEADSTAFF
CBOLEADSTAFF.Text = rs!LEADSTAFF & " " & MHVLEADSTAFF
End If

If rs!SUPPORT1 <> "" Then
FindSPPORT1 rs!SUPPORT1
CBOSUPPORT1.Text = rs!SUPPORT1 & " " & SUPPORT1
End If


If rs!SUPPORT2 <> "" Then
FindSPPORT2 rs!SUPPORT2
CBOSUPPORT2.Text = rs!SUPPORT2 & " " & SUPPORT2
End If

If rs!SUPPORT3 <> "" Then
FindSPPORT3 rs!SUPPORT3
CBOSUPPORT3.Text = rs!SUPPORT3 & " " & SUPPORT3
End If

If rs!SUPPORT4 <> "" Then
FindSPPORT4 rs!SUPPORT4
CBOSUPPORT4.Text = rs!SUPPORT4 & " " & SUPPORT4
End If

If rs!SUPPORT5 <> "" Then
FindSPPORT4 rs!SUPPORT5
CBOSUPPORT4.Text = rs!SUPPORT5 & " " & SUPPORT5
End If




txtregdate.Value = Format(rs!regdate)
TXTOTHER.Text = IIf(IsNull(rs!OTHER), "", rs!OTHER)
TXTCIDOTHER.Text = IIf(IsNull(rs!CIDOTHER), "", rs!CIDOTHER)
txtremarks.Text = IIf(IsNull(rs!remarks), "", rs!remarks)
Select Case rs!plantedstatus
Case "N"
cboplantedstatus.Text = "New"
Case "C"
cboplantedstatus.Text = "Completely Planted"
Case "P"
cboplantedstatus.Text = "Partially Planted"
End Select


Else
MsgBox "No Records Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub CHKFARMER_Click()

End Sub

Private Sub Form_Load()

On Error GoTo err
Operation = ""
txtregdate.Value = Format(Now, "dd/MM/yyyy")
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select concat(cast(trnid as char) ,' ', farmerid,' ',farmername) as farmername,trnid  from tbllandreg as a,tblfarmer as b where a.farmerid=b.idfarmer order by trnid", db
Set cbotrnid.RowSource = RSTR
cbotrnid.ListField = "farmername"
cbotrnid.BoundColumn = "trnid"

Set rssid = Nothing
If rssid.State = adStateOpen Then rssid.Close
rssid.Open "select id,sessionname from tblregsession order by status desc", db
Set cbosessionid.RowSource = rssid
cbosessionid.ListField = "sessionname"
cbosessionid.BoundColumn = "id"


Set rsfr = Nothing
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"


Set rsCg = Nothing
If rsCg.State = adStateOpen Then rsCg.Close
rsCg.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbocgid.RowSource = rsCg
cbocgid.ListField = "farmername"
cbocgid.BoundColumn = "idfarmer"


Set rsMs = Nothing
If rsMs.State = adStateOpen Then rsMs.Close
rsMs.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff where moniter='1' order by staffcode", db
Set cbomhvstaff.RowSource = rsMs
cbomhvstaff.ListField = "staffname"
cbomhvstaff.BoundColumn = "staffcode"


' rsMoniter As New ADODB.Recordset
Set rsMoniter = Nothing
If rsMoniter.State = adStateOpen Then rsMoniter.Close
rsMoniter.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff where moniter='1' order by staffcode", db
Set CBOMONITOR.RowSource = rsMoniter
CBOMONITOR.ListField = "staffname"
CBOMONITOR.BoundColumn = "staffcode"


'Dim rsOutreach As New ADODB.Recordset
Set rsOutreach = Nothing
If rsOutreach.State = adStateOpen Then rsOutreach.Close
rsOutreach.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff where outreach='1' order by staffcode", db
Set CBOOUTREACH.RowSource = rsOutreach
CBOOUTREACH.ListField = "staffname"
CBOOUTREACH.BoundColumn = "staffcode"

'Dim rsLeadstaff As New ADODB.Recordset
Set rsLeadstaff = Nothing
If rsLeadstaff.State = adStateOpen Then rsLeadstaff.Close
rsLeadstaff.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff  order by staffcode", db
Set CBOLEADSTAFF.RowSource = rsLeadstaff
CBOLEADSTAFF.ListField = "staffname"
CBOLEADSTAFF.BoundColumn = "staffcode"


'Dim rssupport1 As New ADODB.Recordset
Set rssupport1 = Nothing
If rssupport1.State = adStateOpen Then rssupport1.Close
rssupport1.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff  order by staffcode", db
Set CBOSUPPORT1.RowSource = rssupport1
CBOSUPPORT1.ListField = "staffname"
CBOSUPPORT1.BoundColumn = "staffcode"

'Dim rssupport2 As New ADODB.Recordset
Set rssupport2 = Nothing
If rssupport2.State = adStateOpen Then rssupport2.Close
rssupport2.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff  order by staffcode", db
Set CBOSUPPORT2.RowSource = rssupport2
CBOSUPPORT2.ListField = "staffname"
CBOSUPPORT2.BoundColumn = "staffcode"
'Dim rssupport3 As New ADODB.Recordset
Set rssupport3 = Nothing
If rssupport3.State = adStateOpen Then rssupport3.Close
rssupport3.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff  order by staffcode", db
Set CBOSUPPORT3.RowSource = rssupport3
CBOSUPPORT3.ListField = "staffname"
CBOSUPPORT3.BoundColumn = "staffcode"
'Dim rssupport4 As New ADODB.Recordset
Set rssupport4 = Nothing
If rssupport4.State = adStateOpen Then rssupport4.Close
rssupport4.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff  order by staffcode", db
Set CBOSUPPORT4.RowSource = rssupport4
CBOSUPPORT4.ListField = "staffname"
CBOSUPPORT4.BoundColumn = "staffcode"

Set rssupport5 = Nothing
If rssupport5.State = adStateOpen Then rssupport4.Close
rssupport5.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff  order by staffcode", db
Set CBOSUPPORT5.RowSource = rssupport5
CBOSUPPORT5.ListField = "staffname"
CBOSUPPORT5.BoundColumn = "staffcode"


'Dim rsind As New ADODB.Recordset
Set rsind = Nothing
If rsind.State = adStateOpen Then rsind.Close
rsind.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff  order by staffcode", db
Set cboindividual.RowSource = rsind
cboindividual.ListField = "staffname"
cboindividual.BoundColumn = "staffcode"

If mbypass = True Then

        TB.Buttons(3).Enabled = True
        Operation = "ADD"
       CLEARCONTROLL
        cbotrnid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select (max(trnid)+1) as maxid from tbllandreg", MHVDB
   If rs.EOF <> True Then
   cbotrnid.Text = rs!MaxId
   Else
   cbotrnid.Text = 1
   End If
   
   cbofarmerid.Text = mFARID
cbofarmerid.Enabled = False

End If
mbypass = False
mFARID = ""



Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub LBLDESC_Click()
LBLDESC.Caption = ""
End Sub

Private Sub Option2_Click()

End Sub

Private Sub OPTBYID_Click()
On Error GoTo err
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsfr = Nothing


cbofarmerid.Text = ""


If OPTNAME.Value = True Then
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(farmername  , ' ',idfarmer) as farmername,idfarmer  from tblfarmer order by FARMERNAME", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

Else
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub OPTNAME_Click()
On Error GoTo err
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsfr = Nothing

cbofarmerid.Text = ""



If OPTNAME.Value = True Then
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(farmername  , ' ',idfarmer) as farmername,idfarmer  from tblfarmer order by FARMERNAME", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

Else
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key
Case "ADD"
       cbofarmerid.Enabled = True
        TB.Buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
        cbotrnid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select (max(trnid)+1) as maxid from tbllandreg", MHVDB
   If rs.EOF <> True Then
   cbotrnid.Text = rs!MaxId
   Else
   cbotrnid.Text = 1
   End If
       
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
       cbotrnid.Enabled = True
      cbofarmerid.Enabled = False
      ' cbogewog.Enabled = True
      TB.Buttons(3).Enabled = True
       
       Case "SAVE"
       If Len(cborelation.Text) = 0 Then
       MsgBox "Please Select The Appropriate Relation."
       Exit Sub
       End If
       MNU_SAVE
        TB.Buttons(3).Enabled = False
        FillGrid
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub CLEARCONTROLL()
cbotrnid.Text = ""
cbofarmerid.Text = ""
txtregdate.Value = Format(Now, "dd/MM/yyyy")
txtthramno.Text = ""
txtthramholdername.Text = ""
cbosex.Text = ""
cborelation.Text = ""
txtregland.Text = ""
cbolandtype.Text = ""
txtagreement.Text = ""
cbocgid.Text = ""
cbomhvstaff.Text = ""
txtremarks.Text = ""
CBOMONITOR.Text = ""
cboindividual.Text = ""
CBOLEADSTAFF.Text = ""
CBOOUTREACH.Text = ""
CBOSUPPORT1.Text = ""
CBOSUPPORT2.Text = ""
CBOSUPPORT3.Text = ""
CBOSUPPORT4.Text = ""
TXTOTHER.Text = ""
TXTCIDOTHER.Text = ""

End Sub
Private Sub MNU_SAVE()
On Error GoTo err
Dim mlandconfirmed As String
Dim msex As Integer
Dim sharedmonitor, sharedoutreach, CG, cgmonitor, others, individual, leastaff, sup1, sup2, sp3, sup4, sup5 As Double
If Len(CBOMONITOR.Text) <> 0 Then
sharedmonitor = 50
sharedoutreach = 50
Else
sharedmonitor = 0
sharedoutreach = 0
End If
If Len(cbosessionid.Text) = 0 Then
MsgBox "Please select registration session."
Exit Sub
End If

If Len(txtagreement.Text) = 0 And Operation = "ADD" Then
MsgBox "Agreement no. is must."
Exit Sub
End If

If Val(txtagreement.Text) = 0 And Operation = "ADD" Then
MsgBox "Agreement no. is must."
Exit Sub
End If


If chkconfirmed.Value = 0 Then
mlandconfirmed = "N"
Else
mlandconfirmed = "Y"
End If


If Len(cbocgid.Text) <> 0 Then
CG = 50
cgmonitor = 50
Else
CG = 0
cgmonitor = 0
End If


If Len(TXTOTHER.Text) <> 0 Then
others = 100

Else
others = 0

End If

If Len(cboindividual.Text) <> 0 Then
individual = 100

Else
individual = 0

End If



Dim rs As New ADODB.Recordset
If Operation = "ADD" Then
   Set rs = Nothing
   rs.Open "select (max(trnid)) as maxid from tbllandreg", MHVDB
   If rs.EOF <> True Then
   cbotrnid.Text = rs!MaxId + 1
   Else
  cbotrnid.Text = 1
   End If
End If

If Mid(cbofarmerid.Text, 10, 1) <> "G" Then
If cbosex.Text = "Male" Then
msex = 0
ElseIf cbosex.Text = "Female" Then
msex = 1
Else
MsgBox "Please Select The appropriate Sex."
Exit Sub
End If
End If

If Len(cboplantedstatus.Text) = 0 Then

MsgBox "Select Planted Status."
Exit Sub
End If


If Len(cbofarmerid.Text) = 0 Then
MsgBox "Please Select The Appropriate Information From The Drop Down Controll.", , "MHV ERROR BOX"
Exit Sub
End If
MHVDB.BeginTrans
If Operation = "ADD" Then
MHVDB.Execute " insert into tbllandreg (trnid,farmerid,regdate,thramno,thramname,sex,relation,regland,landtype,cgid," _
& " CGMONITOR,remarks,monitOr,outreach,individual,leadstaff,support1,support2,support3,support4,other,cidother,SUPPORT5,SHAREDMONITORPERCENT,SHAREDOUTREACHPERCENT,CGIDPERCENT,CGMONITORPERCENT,OTHERSPERCENT,INDIVIDUALPERCENT,LEADSTAFFPERCENT,SUPPORT1PERCENT,SUPPORT2PERCENT,SUPPORT3PERCENT,SUPPORT4PERCENT,SUPPORT5PERCENT,status,insertedby,inserteddate,islandconfirmed,plantedstatus,sessionid) " _
& " values('" & cbotrnid.Text & "','" & cbofarmerid.BoundText & "','" & Format(txtregdate.Value, "yyyy-MM-dd") & "','" & txtthramno.Text & "'" _
& " ,'" & txtthramholdername.Text & "','" & msex & "','" & cborelation.Text & "','" & Val(txtregland.Text) & "'" _
& " ,'" & cbolandtype.Text & "','" & cbocgid.BoundText & "','" & cbomhvstaff.BoundText & "','" & txtremarks.Text & "'" _
& " ,'" & CBOMONITOR.BoundText & "','" & CBOOUTREACH.BoundText & "','" & cboindividual.BoundText & "','" & CBOLEADSTAFF.BoundText & "'" _
& ",'" & CBOSUPPORT1.BoundText & "','" & CBOSUPPORT2.BoundText & "','" & CBOSUPPORT3.BoundText & "','" & CBOSUPPORT4.BoundText & "'" _
& ",'" & TXTOTHER.Text & "','" & TXTCIDOTHER.Text & "','" & CBOSUPPORT5.BoundText & "','" & sharedmonitor & "','" & sharedoutreach & "','" & CG & "','" & cgmonitor & "','" & others & "','" & individual & "','" & Val(TXTLEADSTAFFPERCENT.Text) & "','" & Val(TXTSUPPORT1PERCENT.Text) & "','" & Val(TXTSUPPORT2PERCENT.Text) & "','" & Val(TXTSUPPORT3PERCENT.Text) & "','" & Val(TXTSUPPORT4PERCENT.Text) & "','" & Val(TXTSUPPORT5PERCENT.Text) & "','A','" & MUSER & "','" & Format(Now, "yyyy-MM-dd") & "','" & mlandconfirmed & "','" & Mid(cboplantedstatus.Text, 1, 1) & "','" & cbosessionid.BoundText & "')"

ElseIf Operation = "OPEN" Then
MHVDB.Execute " update  tbllandreg set farmerid='" & cbofarmerid.BoundText & "',regdate='" & Format(txtregdate.Value, "yyyy-MM-dd") & "', " _
                & "thramno='" & txtthramno.Text & "',thramname='" & txtthramholdername.Text & "',sex='" & msex & "'," _
                & "relation='" & cborelation.Text & "',regland='" & Val(txtregland.Text) & "'," _
                & "landtype='" & cbolandtype.Text & "',cgid='" & cbocgid.BoundText & "'," _
                & " cgmonitOr='" & cbomhvstaff.BoundText & "',remarks='" & txtremarks.Text & "'  ," _
                & "monitOr='" & CBOMONITOR.BoundText & "',outreach='" & CBOOUTREACH.BoundText & "',individual='" & cboindividual.BoundText & "'," _
                & "leadstaff='" & CBOLEADSTAFF.BoundText & "',support1='" & CBOSUPPORT1.BoundText & "',support2='" & CBOSUPPORT2.BoundText & "'," _
                & "support3='" & CBOSUPPORT3.BoundText & "',support4='" & CBOSUPPORT4.BoundText & "',plantedstatus='" & Mid(cboplantedstatus.Text, 1, 1) & "'," _
                & "other='" & TXTOTHER.Text & "',sessionid='" & cbosessionid.BoundText & "',cidother='" & TXTCIDOTHER.Text & "',regdate='" & Format(txtregdate.Value, "yyyy-MM-dd") & "',updatedby='" & MUSER & "',updateddate='" & Format(Now, "yyyy-MM-dd") & "',islandconfirmed='" & mlandconfirmed & "'" _
                & " where trnid='" & cbotrnid.Text & "' "




Else
MsgBox "OPERATION NOT SELECTED."
End If


If Len(txtagreement.Text) > 0 Or Val(txtagreement.Text) > 0 Then
MHVDB.Execute "UPDATE mhv.tblfarmer  SET agreementno ='" & txtagreement.Text & "' WHERE IDFARMER ='" & Mid(cbofarmerid.Text, 1, 14) & "'"
End If
MHVDB.Execute "UPDATE mhv.tblregistrationsettings SET updatetable ='Yes' WHERE tblid ='1'"
MHVDB.Execute "UPDATE mhv.tblregistrationsettings SET updatetable ='Yes' WHERE tblid ='5'"
MHVDB.Execute "UPDATE mhv.tblregistrationsettings SET updatetable ='Yes' WHERE tblid ='6'"
MHVDB.CommitTrans
Exit Sub
err:
MsgBox err.Description
MHVDB.RollbackTrans
End Sub

Private Sub FillGrid()

End Sub

Private Sub TXTCIDOTHER_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TXTLEADSTAFFPERCENT_LostFocus()
If Val(TXTLEADSTAFFPERCENT.Text) > 0 Then
Command2.Enabled = True
Else
Command2.Enabled = False
End If
End Sub

Private Sub TXTOTHER_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = TXTOTHER.SelStart + 1
    Dim sText As String
    sText = Left$(TXTOTHER.Text, iPos)
    If iPos = 1 Then GoTo Upit
    KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
    If iPos > 1 And _
     (InStr(iPos - 1, sText, " ") > 0 Or _
      InStr(iPos - 1, sText, "-") > 0 Or _
      InStr(iPos - 1, sText, ".") > 0 Or _
      InStr(iPos - 1, sText, "'") > 0) _
      Then GoTo Upit
    If iPos > 2 Then _
      If InStr(iPos - 2, sText, "Mc") > 0 _
        Then GoTo Upit
        
   End If
  Exit Sub
Upit:
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtregland_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtregland_LostFocus()
If Mid(UCase(cbofarmerid.Text), 10, 1) = "G" Or Mid(UCase(cbofarmerid.Text), 10, 1) = "C" Then
cboplantedstatus.Text = "Partial Plantation"
Else
cboplantedstatus.Text = "New"
End If
End Sub

Private Sub txtthramholdername_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtthramholdername.SelStart + 1
    Dim sText As String
    sText = Left$(txtthramholdername.Text, iPos)
    If iPos = 1 Then GoTo Upit
    KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
    If iPos > 1 And _
     (InStr(iPos - 1, sText, " ") > 0 Or _
      InStr(iPos - 1, sText, "-") > 0 Or _
      InStr(iPos - 1, sText, ".") > 0 Or _
      InStr(iPos - 1, sText, "'") > 0) _
      Then GoTo Upit
    If iPos > 2 Then _
      If InStr(iPos - 2, sText, "Mc") > 0 _
        Then GoTo Upit
        
   End If
  Exit Sub
Upit:
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Private Sub FindFarmer(ff As String)
On Error GoTo err
FrName = ""
magno = ""
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblfarmer where idfarmer='" & ff & "'", MHVDB
If rs.EOF <> True Then
FrName = rs!farmername
magno = rs!agreementno
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FindCG(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
CGname = ""
Set rs = Nothing
rs.Open "select * from tblfarmer where idfarmer='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
CGname = rs!farmername
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FindMHV(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
MHVName = "" = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
MHVName = rs!staffname
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FindMONITOR(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
MHVMONITOR = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
MHVMONITOR = rs!staffname
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FindINDIVIDUAL(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
MHVINDIVIDUAL = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
MHVINDIVIDUAL = rs!staffname
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FindLEADSTAFF(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
MHVLEADSTAFF = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
MHVLEADSTAFF = rs!staffname
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FindSPPORT1(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
SUPPORT1 = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
SUPPORT1 = rs!staffname
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FindSPPORT2(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
SUPPORT2 = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
SUPPORT2 = rs!staffname
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub FindSPPORT3(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
SUPPORT3 = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
SUPPORT3 = rs!staffname
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub FindSPPORT4(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
SUPPORT4 = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
SUPPORT4 = rs!staffname
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FindSPPORT5(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
SUPPORT5 = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
SUPPORT5 = rs!staffname
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FindOUTREACH(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
MHVOUTREACH = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
MHVOUTREACH = rs!staffname
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub


