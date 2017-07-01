VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_Print_Beli 
   BackColor       =   &H00FF8080&
   Caption         =   "Cetak Bill"
   ClientHeight    =   4965
   ClientLeft      =   5760
   ClientTop       =   3585
   ClientWidth     =   7020
   ControlBox      =   0   'False
   Icon            =   "Form_Print_Beli.frx":0000
   LinkTopic       =   "Form13"
   ScaleHeight     =   4965
   ScaleWidth      =   7020
   Begin MSComctlLib.ListView list_supplier 
      Height          =   2055
      Left            =   3120
      TabIndex        =   11
      Top             =   2760
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama"
         Object.Width           =   4940
      EndProperty
   End
   Begin VB.ComboBox cb_bayar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form_Print_Beli.frx":628A
      Left            =   3120
      List            =   "Form_Print_Beli.frx":6297
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox txt_kode_supplier 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txt_nama_supplier 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox txt_total 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   7
      Text            =   "0"
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txt_bon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton btn_print 
      BackColor       =   &H0080FFFF&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton btn_batal 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "NOMOR BON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "SUPPLIER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "PEMBAYARAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
   End
End
Attribute VB_Name = "Form_Print_Beli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBill, rsbeli As ADODB.Recordset
Dim is_new As Boolean
Dim txt_sup_toggle As Boolean

Private Sub btn_batal_Click()
    Unload Me
End Sub

Private Sub cb_bayar_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub Form_Load()
    txt_sup_toggle = True
    reload_Supplier
End Sub

Private Sub Form_unload(cancel As Integer)
    If is_new Then
        Form_Pembelian.Enabled = True
    End If
End Sub

Public Sub Init(no_bon As String, total As String, new_bon As Boolean)
    txt_bon.Enabled = False
    txt_bon.Text = no_bon
    txt_total.Text = total
    is_new = new_bon
    list_supplier.Visible = False
    If priceToNum(total) > 600000 Then
        cb_bayar.ListIndex = 1
    Else
        cb_bayar.ListIndex = 0
    End If
    
    txt_sup_toggle = True
    If is_new Then
        If getSupplier(Form_Pembelian.kode_supplier) Then
            txt_kode_supplier = rsSupplier!kdsuplier
            txt_nama_supplier = rsSupplier!nmsuplier
            btn_print.SetFocus
        Else
            txt_kode_supplier.SetFocus
        End If
        
    Else
        Set rsBill = con.Execute("select * from bill_beli where nobukti = '" & no_bon & "'")
        btn_print.SetFocus
        txt_kode_supplier = rsBill!kode_supplier
        txt_total = Format(rsBill!total, "###,###,###")
        If getSupplier(rsBill!kode_supplier) Then
            txt_nama_supplier = rsSupplier!nmsuplier
        End If
        cb_bayar.ListIndex = rsBill!pembayaran
    End If
End Sub

Private Sub list_supplier_LostFocus()
    list_supplier.Visible = False
End Sub

Private Sub list_supplier_dblclick()
    If list_supplier.SelectedItem.index >= 0 Then
        txt_kode_supplier = list_supplier.SelectedItem.Text
        txt_nama_supplier = list_supplier.SelectedItem.SubItems(1)
        cb_bayar.SetFocus
    End If
End Sub

Private Sub list_supplier_keydown(key As Integer, Shift As Integer)
    If key = 13 Then
        list_supplier_dblclick
    End If
End Sub

Private Sub btn_print_click()
    If txt_kode_supplier = "" Or txt_nama_supplier = "" Then
        MsgBox "Data Supplier tidak valid"
        Exit Sub
    End If
    
    If is_new Then
        Dim i As Integer
        i = 1
        Dim tanggal As String
        tanggal = Format(Now, "yyyy-mm-dd")
        
        Do While i <= Form_Pembelian.lv_beli.ListItems.count
            Dim item As ListItem
            Set item = Form_Pembelian.lv_beli.ListItems(i)
            'editV2
            con.Execute ("insert into tbbeli (nobukti, tglbukti, kode, nama_barang, harga, jumlah, `return`) values('" & txt_bon & "', '" & tanggal & "', '" & item.Text & "', '" & item.SubItems(1) & "', " & priceToNum(item.SubItems(2)) & ", " & item.SubItems(3) & ", " & item.SubItems(4) & ")")
            con.Execute ("update tbbarang set jumlah_akhir = jumlah_akhir + " & (Val(item.SubItems(3)) - Val(item.SubItems(4))) & ", tgl_masuk='" & Format(Now, "yyyy-mm-dd") & "' where kode = '" & item.Text & "'")
            i = i + 1
        Loop
        'editV2
        con.Execute ("insert into bill_beli (nobukti, staff, tanggal, jam, total, kode_supplier, pembayaran, lunas, settled, tanggal_lunas) values('" & txt_bon & "','" & username & "', '" & tanggal & "', '" & Format(Now, "hh:mm:ss") & "', " & priceToNum(txt_total) & ", " & txt_kode_supplier & ", " & cb_bayar.ListIndex & ", 0,0,'1990-09-26')")
        Set rsBill = con.Execute("select * from bill_beli where nobukti = '" & txt_bon & "'")
    Else
        con.Execute ("update bill_beli set kode_supplier = " & txt_kode_supplier & ", pembayaran = " & cb_bayar.ListIndex & " where nobukti = '" & txt_bon & "'")
    End If
    
    Set rsbeli = con.Execute("select * from tbbeli where nobukti = '" & txt_bon & "'")
    If rsbeli.EOF Then
        MsgBox "data tidak ditemukan"
        Exit Sub
    End If
    rsbeli.MoveFirst
    
    Printer.Font = "times new roman"
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.Print Tab(2); " BON PEMBELIAN";
    Printer.Print Tab(2); "CHRISTINE HAKIM";
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print Tab(3); "                                                            "; 'baris kosong
    Printer.Print Tab(3); "Jl. Adinegoro No. 11A Padang";
    Printer.Print Tab(3); "                                                             ";
    Printer.Print Tab(3); "No. FAKTUR"; Tab(20); ": "; txt_bon.Text;
    Printer.Print Tab(3); "Staff"; Tab(20); ": "; rsBill!staff;
    Printer.Print Tab(3); "Supplier"; Tab(20); ": ["; txt_kode_supplier.Text; "] "; txt_nama_supplier.Text;
    Printer.Print Tab(3); "Pembayaran"; Tab(20); ": "; cb_bayar.Text;
    Printer.Print Tab(3); Format(rsBill!tanggal, "dd-MM-yyyy"); "  "; rsBill!jam;
    Printer.Print Tab(3); "---------------------------------------------------------------------------";
    Do While Not rsbeli.EOF
        Printer.Print Tab(3); rsbeli!nama_barang
        Dim bayar As Long
        bayar = (rsbeli!jumlah - rsbeli!return) * rsbeli!harga
        Printer.Print Tab(3); rsbeli!jumlah; "-"; rsbeli!return; "x"; Format(rsbeli!harga, "###,###,###"); Tab(35); Format(bayar, "###,###,###")
        rsbeli.MoveNext
    Loop
    Printer.FontSize = 14
    Printer.Print Tab(3); "                                                                         ";
    Printer.Print Tab(10); "Total: "; Format(txt_total, "###,###,###")
    Printer.CurrentX = 0
    Printer.FontSize = 10
    Printer.Print Tab(3); "                                                                         ";
    Printer.FontBold = True
    Printer.Print Tab(3); "Simpanlah faktur ini sebaik-baiknya";
    Printer.FontBold = False
    Printer.Print Tab(3); "Faktur tidak dapat dicetak ulang";
    Printer.Print Tab(3); "Penagihan wajib disertai faktur resmi";
    Printer.Print Tab(3); "----------------------------------------------------------------------------";
    Printer.Print Tab(3); "                                                       ";
    Printer.Font = "courier new"
    Printer.Print Tab(3); "Diterima oleh"; Tab(18); "Dibayar oleh";
    Printer.Print Tab(3); "                                                       ";
    Printer.Print Tab(3); "                                                       ";
    Printer.Print Tab(3); "                                                       ";
    Printer.Print Tab(3); "(___________)"; Tab(18); "(___________)";
    Printer.Print Tab(3); "                                                       ";
    Printer.Print Tab(3); "        *Terima Kasih*";
      
Close #1
Printer.EndDoc
    
    If is_new Then
        Form_Pembelian.nextFaktur
    Else
        Form_List_beli.refreshlist
    End If
    
    Unload Me
End Sub


Private Sub txt_kode_supplier_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        txt_sup_toggle = True
        
'        Set rsSupplier = con.Execute("select * from tbsuplier")
        
        If getSupplier(txt_kode_supplier) Then
            txt_nama_supplier.Text = rsSupplier!nmsuplier
            cb_bayar.SetFocus
        Else
            MsgBox "Supplier tidak terdaftar"
            txt_kode_supplier.Text = ""
        End If
    Else
        txt_nama_supplier = ""
    End If
End Sub

Private Sub txt_kode_supplier_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
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

Private Sub txt_nama_supplier_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 3)
End Sub

Private Sub txt_nama_supplier_LostFocus()
    If Not Me.ActiveControl.Name = "list_supplier" Then
        list_supplier.Visible = False
    End If
End Sub

Private Sub txt_nama_supplier_KeyDown(key As Integer, Shift As Integer)
    If key = 40 Then
        list_supplier.Visible = True
        list_supplier.SetFocus
    ElseIf key = 13 And list_supplier.Visible = True Then
        list_supplier.SetFocus
    ElseIf key = 13 And list_supplier.Visible = False Then
        cb_bayar.SetFocus
    Else
        txt_kode_supplier = ""
    End If
    
'    txt_kode_supplier = ""
'    list_supplier.Visible = True
'    list_supplier.ListItems.Clear
'    Dim rsSup As ADODB.Recordset
'    Set rsSup = con.Execute("select * from tbsuplier where nmsuplier like '%" & txt_nama_supplier & "%'")
'    If rsSup.EOF Or rsSup.BOF Then
'        Exit Sub
'    End If
'
'    rsSup.MoveFirst
'    Do While Not rsSup.EOF
'        list_supplier.ListItems.Add(, , rsSup!kdsuplier).SubItems(1) = rsSup!nmsuplier
'        rsSup.MoveNext
'    Loop
'
'    Set rsSup = Nothing
End Sub

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

