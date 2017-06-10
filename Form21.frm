VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form21 
   BackColor       =   &H0000C000&
   Caption         =   "Laporan"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8310
   Icon            =   "Form21.frx":0000
   LinkTopic       =   "Form21"
   ScaleHeight     =   5760
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Cetak"
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
      Left            =   3000
      TabIndex        =   22
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cetak Lap. Stok"
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
      Left            =   3000
      TabIndex        =   21
      Top             =   3240
      Width           =   1575
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   7800
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6600
      Top             =   3720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=foodcourt"
      OLEDBString     =   "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=foodcourt"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbsuplier"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cetak Lap. Terlaris"
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
      Left            =   3000
      TabIndex        =   17
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cetak"
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
      Left            =   7080
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   495
      Left            =   5280
      TabIndex        =   13
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   131530755
      CurrentDate     =   42206
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   131530755
      CurrentDate     =   42206
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cetak"
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
      Left            =   7080
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   131530755
      CurrentDate     =   42206
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   131530755
      CurrentDate     =   42206
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2280
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4965
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel/Clear"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak"
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
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker9 
      Height          =   495
      Left            =   3000
      TabIndex        =   20
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   131530755
      CurrentDate     =   42206
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Laporan Hutang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   19
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Laporan Stok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   18
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Laporan Terlaris"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "S/D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4680
      TabIndex        =   14
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Laporan Pembelian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "S/D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Laporan Penjualan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Penanggung Jawab"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Laporan Harian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum LedgerColours
vbledgerWhite = &HF9FEFF
vbLedgerGreen = &HD0FFCC
vbLedgerYellow = &HE1FAFF
vbLedgerRed = &HE1E1FF
vbLedgerGrey = &HE0E0E0
vbLedgerBeige = &HD9F2F7
vbLedgerSoftWhite = &HF7F7F7
vbledgerPureWhite = &HFFFFFF
End Enum

Private Enum ImageSizingTypes
[sizeNone] = 0
[sizeCheckBox]
[sizeIcon]
End Enum

Private Sub kosong()
'Text3.Text = ""
'Text4.Text = ""
End Sub

Private Sub Command1_Click()
Dim pass As String, b As String
Dim rsSuplier As New ADODB.Recordset

pass = "Provider=MSDASQL.1;Pwd=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
cr.connect = pass
cr.ReportFileName = App.Path & "\laporanharian.rpt"
cr.SelectionFormula = "(day({tbjual.tglbukti}) = " & Day(DTPicker9.Value) & " and month({tbjual.tglbukti}) = " & Month(DTPicker9.Value) & " and year({tbjual.tglbukti}) = " & Year(DTPicker9.Value) & ")"
cr.WindowState = crptMaximized
cr.Formulas(0) = "tgl='" & Format(DTPicker9.Value, "dd/MM/yyyy") & "'"
cr.RetrieveDataFiles
cr.Action = 1
cr.reset
End Sub

Private Sub Command4_Click()
Dim pass As String, b As String
pass = "Provider=MSDASQL.1;Pwd=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
cr.connect = pass
cr.ReportFileName = App.Path & "\jual.rpt"
cr.SelectionFormula = "{tbjual.tglbukti} >= #" & Format(DTPicker1.Value, "yyyy-mm-dd") & "# and {tbjual.tglbukti} <= #" & Format(DTPicker2.Value, "yyyy-mm-dd") & "# "
cr.WindowState = crptMaximized
cr.Formulas(0) = "tgl='" & Format(DTPicker1.Value, "dd/MM/yyyy") & "'"
cr.Formulas(1) = "t2='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
cr.RetrieveDataFiles
cr.Action = 1
cr.reset
End Sub

Private Sub Command5_Click()
Dim pass As String, b As String
pass = "Provider=MSDASQL.1;Pwd=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
cr.connect = pass
cr.ReportFileName = App.Path & "\beli.rpt"
cr.SelectionFormula = "{tbbeli.tglbukti} >= #" & Format(DTPicker3.Value, "yyyy-mm-dd") & "# and {tbbeli.tglbukti} <= #" & Format(DTPicker4.Value, "yyyy-mm-dd") & "# "
cr.WindowState = crptMaximized
cr.Formulas(0) = "tgl='" & Format(DTPicker3.Value, "dd/MM/yyyy") & "'"
cr.Formulas(1) = "t2='" & Format(DTPicker4.Value, "dd/MM/yyyy") & "'"
cr.RetrieveDataFiles
cr.Action = 1
cr.reset
End Sub

Private Sub Command6_Click()
Dim pass As String, b As String
pass = "Provider=MSDASQL.1;Pwd=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
cr.connect = pass
cr.ReportFileName = App.Path & "\terlaris.rpt"
'Cr.SelectionFormula = "{tbbarang.kode} <> 0 "
cr.WindowState = crptMaximized
cr.RetrieveDataFiles
cr.Action = 1
cr.reset
End Sub

Private Sub Command7_Click()
Dim pass As String, b As String
pass = "Provider=MSDASQL.1;Pwd=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
cr.connect = pass
cr.ReportFileName = App.Path & "\laporanstok.rpt"
cr.WindowState = crptMaximized
cr.RetrieveDataFiles
cr.Action = 1
cr.reset
End Sub

Private Sub Command8_Click()
Dim pass As String, b As String
pass = "Provider=MSDASQL.1;Pwd=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
cr.connect = pass
cr.ReportFileName = App.Path & "\laporanhutang.rpt"
cr.WindowState = crptMaximized
cr.RetrieveDataFiles
cr.Action = 1
cr.reset
End Sub
Private Sub Form_Activate()
'Text3.SetFocus
End Sub
Private Sub Command2_Click()
kosong
End Sub

Private Sub Command3_Click()
Unload Me
If con.State = adStateOpen Then
con.Close
End If
End Sub

Private Sub Form_Load()
If con.State = adStateClosed Then
connect
End If
Dim strsql As String
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date
DTPicker4.Value = Date
DTPicker9.Value = Date
kosong
Text1 = username
End Sub
