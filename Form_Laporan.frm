VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form_Laporan 
   Caption         =   "Laporan"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_Absen 
      Caption         =   "Laporan Absen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Width           =   3135
   End
   Begin Crystal.CrystalReport cr 
      Left            =   3240
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton btn_hutang 
      Caption         =   "Laporan Hutang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   9
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CommandButton btn_pembayaran 
      Caption         =   "Laporan Pembayaran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CommandButton btn_stok 
      Caption         =   "Laporan Stok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   7
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton btn_pengeluaran 
      Caption         =   "Laporan Pengeluaran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton btn_penjualan 
      Caption         =   "Laporan Penjualan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   5
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton btn_harian 
      Caption         =   "Laporan Harian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker dt_start 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   112852993
      CurrentDate     =   42810
   End
   Begin MSComCtl2.DTPicker dt_end 
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   112852993
      CurrentDate     =   42810
   End
   Begin VB.Label Label2 
      Caption         =   "Sampai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form_Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NO_DATE As Integer = 0
Const ONE_DAY As Integer = 1
Const DURATION As Integer = 2

Private Sub btn_Absen_Click()
    Call openReport("laporanabsensi.rpt", "tbabsen.tanggal", DURATION)
End Sub

Private Sub btn_harian_Click()
    Call openReport("laporanharian.rpt", "bill.tanggal", ONE_DAY)
End Sub

Private Sub openReport(file_name As String, date_parameter As String, report_type As Integer)
    cr.connect = "Provider=MSDASQL.1;Pwd=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
    cr.ReportFileName = App.Path + "\" + file_name
    If report_type = ONE_DAY Then
        cr.SelectionFormula = "{" & date_parameter & "}= #" & Format(dt_start.Value, "yyyy-MM-dd") & "#"
        cr.Formulas(0) = "tgl='" & Format(dt_start.Value, "dd/MM/yyyy") & "'"
    ElseIf report_type = DURATION Then
        cr.SelectionFormula = "{" & date_parameter & "}>= #" & Format(dt_start.Value, "yyyy-MM-dd") & "# and {" & date_parameter & "}<= #" & Format(dt_end.Value, "yyyy-MM-dd") & "#"
        cr.Formulas(0) = "tgl='" & Format(dt_start.Value, "dd/MM/yyyy") & "'"
        cr.Formulas(1) = "tgl2='" & Format(dt_end.Value, "dd/MM/yyyy") & "'"
    End If
    cr.WindowState = crptMaximized
    cr.RetrieveDataFiles
    cr.Action = 1
    cr.reset
End Sub

Private Sub btn_hutang_Click()
    Call openReport("laporanhutang.rpt", "", NO_DATE)
End Sub

Private Sub btn_pembayaran_Click()
    Call openReport("laporanpembayaran.rpt", "", NO_DATE)
End Sub

Private Sub btn_pengeluaran_Click()
    Call openReport("laporanpengeluaran.rpt", "bill_beli.tanggal_lunas", DURATION)
End Sub

Private Sub btn_penjualan_Click()
     Call openReport("laporanpenjualan.rpt", "tbjual.tglbukti", DURATION)
End Sub

Private Sub btn_stok_Click()
    Call openReport("laporanstok.rpt", "", NO_DATE)
End Sub

Private Sub dt_end_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub dt_start_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub Form_Load()
    dt_start.Value = Now
    dt_end.Value = Now
End Sub
