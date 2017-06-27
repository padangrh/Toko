VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form_Laporan 
   Caption         =   "Laporan"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_kode_supplier 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      TabIndex        =   14
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox txt_nama_supplier 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   960
      TabIndex        =   13
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton btn_barang 
      Caption         =   "Laporan Barang"
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
      TabIndex        =   12
      Top             =   5520
      Width           =   3135
   End
   Begin MSComctlLib.ListView list_supplier 
      Height          =   2175
      Left            =   960
      TabIndex        =   11
      Top             =   3720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama"
         Object.Width           =   3176
      EndProperty
   End
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
      Format          =   93388801
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
      Format          =   93388801
      CurrentDate     =   42810
   End
   Begin VB.Label Label3 
      Caption         =   "Supplier"
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
      TabIndex        =   15
      Top             =   5400
      Width           =   1455
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
Dim txt_sup_toggle As Boolean

Private Sub btn_Absen_Click()
    Call openReport("laporanabsensi.rpt", "tbabsen.tanggal", DURATION, True)
End Sub

Private Sub btn_barang_Click()
    Call openReport("laporanbarang.rpt", "v_jual_supplier.tanggal", DURATION, False)
    cr.SelectionFormula = cr.SelectionFormula + " and {v_jual_supplier.kode_supplier}='" & txt_kode_supplier & "'"
    runCrystalReport
End Sub

Private Sub btn_harian_Click()
    Call openReport("laporanharian.rpt", "bill.tanggal", ONE_DAY, True)
End Sub

Private Sub openReport(file_name As String, date_parameter As String, report_type As Integer, auto_run As Boolean)
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
'    cr.WindowState = crptMaximized
'    cr.RetrieveDataFiles
'    cr.Action = 1
'    cr.reset
    If auto_run Then runCrystalReport
End Sub

Private Sub btn_hutang_Click()
    Call openReport("laporanhutang.rpt", "", NO_DATE, True)
End Sub

Private Sub btn_pembayaran_Click()
    Call openReport("laporanpembayaran.rpt", "", NO_DATE, True)
End Sub

Private Sub btn_pengeluaran_Click()
    Call openReport("laporanpengeluaran.rpt", "bill_beli.tanggal_lunas", DURATION, True)
End Sub

Private Sub btn_penjualan_Click()
     Call openReport("laporanpenjualan.rpt", "tbjual.tglbukti", DURATION, True)
End Sub

Private Sub btn_stok_Click()
    Call openReport("laporanstok.rpt", "", NO_DATE, True)
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
    
    list_supplier.Visible = False
    txt_sup_toggle = True
    btn_barang.Enabled = False
    
    btn_harian.Enabled = CheckPath("laporanharian.rpt")
    btn_penjualan.Enabled = CheckPath("laporanpenjualan.rpt")
    btn_pengeluaran.Enabled = CheckPath("laporanpengeluaran.rpt")
    btn_pembayaran.Enabled = CheckPath("laporanpembayaran.rpt")
    btn_hutang.Enabled = CheckPath("laporanhutang.rpt")
    btn_stok.Enabled = CheckPath("laporanstok.rpt")
    btn_Absen.Enabled = CheckPath("laporanabsensi.rpt")
End Sub

Private Sub runCrystalReport()
    cr.WindowState = crptMaximized
    cr.RetrieveDataFiles
    cr.Action = 1
    cr.reset
End Sub

Private Function CheckPath(strPath As String) As Boolean
    If Dir$(App.Path + "\" + strPath) <> "" Then
        CheckPath = True
    Else
        CheckPath = False
    End If
End Function

Private Sub reload_Supplier()
    'list_supplier.Visible = True
    list_supplier.ListItems.Clear
    Dim rsSup As ADODB.Recordset
    Set rsSup = con.Execute("select * from tbsuplier where nmsuplier like '%" & txt_nama_supplier & "%'")
    If rsSup.EOF Then
        list_supplier.Visible = False
        Exit Sub
    End If
    
    rsSup.MoveFirst
    Do While Not rsSup.EOF
        list_supplier.ListItems.Add(, , rsSup!kdsuplier).SubItems(1) = rsSup!nmsuplier
        rsSup.MoveNext
    Loop
    
    Set rsSup = Nothing
End Sub

Private Sub txt_kode_supplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_sup_toggle = True
        
        Set rsSupplier = con.Execute("select * from tbsuplier")
        
        If getSupplier(txt_kode_supplier) Then
            txt_nama_supplier.Text = rsSupplier!nmsuplier
            btn_barang.Enabled = True
            btn_barang.SetFocus
        Else
            MsgBox "Supplier tidak terdaftar"
            txt_kode_supplier.Text = ""
        End If
    Else
        txt_nama_supplier = ""
    End If
End Sub

Private Sub txt_nama_supplier_Change()
    If txt_nama_supplier.Text <> "" And txt_sup_toggle = False Then
        list_supplier.Visible = True
        reload_Supplier
    Else
        list_supplier.Visible = False
        txt_sup_toggle = False
    End If
End Sub

Private Sub txt_nama_supplier_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 38 Then
        list_supplier.Visible = True
        list_supplier.SetFocus
    ElseIf KeyCode = 13 And list_supplier.Visible = True Then
        list_supplier.SetFocus
    ElseIf KeyCode = 13 And list_supplier.Visible = False Then
        cb_bayar.SetFocus
    Else
        txt_kode_supplier = ""
        btn_barang.Enabled = False
    End If
    
End Sub

Private Sub txt_nama_supplier_LostFocus()
    If Not Me.ActiveControl.Name = "list_supplier" Then
        list_supplier.Visible = False
    End If
End Sub
