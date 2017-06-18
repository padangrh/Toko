VERSION 5.00
Begin VB.Form Form_Print 
   BackColor       =   &H000080FF&
   Caption         =   "Cetak Bill"
   ClientHeight    =   5835
   ClientLeft      =   5760
   ClientTop       =   3585
   ClientWidth     =   7020
   ControlBox      =   0   'False
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form13"
   ScaleHeight     =   5835
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_diskon 
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
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox txt_kembali 
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
      TabIndex        =   10
      Text            =   "0"
      Top             =   3840
      Width           =   3495
   End
   Begin VB.TextBox txt_uang 
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
      TabIndex        =   9
      Top             =   3000
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton btn_nontunai 
      BackColor       =   &H0080FF80&
      Caption         =   "Non Tunai"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton btn_tunai 
      BackColor       =   &H0080FFFF&
      Caption         =   "Tunai"
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
      Top             =   5040
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
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "DISKON"
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
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
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
      BackColor       =   &H000080FF&
      Caption         =   "TOTAL BELANJA"
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
      BackColor       =   &H000080FF&
      Caption         =   "JUMLAH UANG"
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
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "KEMBALIAN"
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
      Top             =   3960
      Width           =   2655
   End
End
Attribute VB_Name = "Form_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBill, rsJual, rsDis As ADODB.Recordset
Dim is_new As Boolean

Dim dis_spv As String
Dim dis_status As String
Dim dis_cust As String

Private Sub btn_batal_Click()
    Unload Me
End Sub

Private Sub btn_nontunai_Click()
    txt_uang = 0
    print_bon (0)
End Sub

Private Sub btn_tunai_Click()
    print_bon (1)
End Sub

Private Sub Form_unload(cancel As Integer)
    If is_new Then
        Form_Penjualan.Enabled = True
    End If
End Sub

Public Sub Init(no_bon As String, total As String, new_bon As Boolean)
    txt_bon.Enabled = False
    txt_bon.Text = no_bon
    is_new = new_bon
    
    If is_new Then
        txt_total.Text = total
        txt_uang.SetFocus
    Else
        Set rsBill = con.Execute("select * from bill where nobukti = '" & no_bon & "'")
        txt_uang = Format(rsBill!bayar, "###,###,##0")
        txt_diskon = Format(rsBill!diskon, "###,###,##0")
        txt_total = Format(rsBill!total, "###,###,##0")
        If rsBill!bayar = 0 Then
            btn_nontunai.SetFocus
        Else
            txt_kembali = Format(rsBill!bayar - rsBill!total + rsBill!diskon, "###,###,##0")
            btn_tunai.SetFocus
        End If
        If Val(txt_diskon) > 0 Then
            Set rsDis = con.Execute("select * from tbdiskon where nobukti = '" & no_bon & "'")
            dis_spv = rsDis!supervisor
            dis_cust = rsDis!customer
            dis_status = rsDis!status
        End If
    End If
End Sub


Private Sub print_bon(tunai As Integer)
   
    If tunai = 1 Then
        If Val(txt_uang) < (Val(txt_total) - Val(txt_diskon)) Then
            MsgBox "Jumlah Pembayaran Kurang (enter)", vbOKOnly + vbInformation, "Cek Jumlah Pembayaran"
            txt_uang.SetFocus
            Exit Sub
        Else
            txt_kembalian = Val(txt_uang) - (Val(txt_total) - Val(txt_diskon))
        End If
    End If
    
    If is_new Then
        Dim i As Integer
        i = 1
        Dim tanggal As String
        tanggal = Format(Now, "yyyy-mm-dd")
        
        Do While i <= Form_Penjualan.lv_jual.ListItems.count
            Dim item As ListItem
            Set item = Form_Penjualan.lv_jual.ListItems(i)
            con.Execute ("insert into tbjual values('" & txt_bon & "', '" & tanggal & "', '" & item.Text & "', '" & item.SubItems(1) & "', " & priceToNum(item.SubItems(2)) & ", " & item.SubItems(3) & ")")
            con.Execute ("update tbbarang set jumlah_akhir = jumlah_akhir - " & item.SubItems(3) & " where kode = '" & item.Text & "'")
            i = i + 1
        Loop
        
        con.Execute ("insert into bill values('" & txt_bon & "','" & username & "', '" & tanggal & "', '" & Format(Now, "hh:mm:ss") & "', " & priceToNum(txt_total) & ", " & Val(txt_uang) & ", " & tunai & ", " & priceToNum(txt_diskon) & ")")
        Set rsBill = con.Execute("select * from bill where nobukti = '" & txt_bon & "'")
        If (Val(txt_diskon.Text)) > 0 Then
            con.Execute ("insert into tbdiskon values('" & txt_bon & "', '" & dis_spv & "', '" & dis_status & "', '" & dis_cust & "', " & priceToNum(txt_diskon) & ")")
        End If
    Else
        con.Execute ("update bill set cash = " & tunai & ", bayar = " & priceToNum(txt_uang) & ", diskon = " & priceToNum(txt_diskon) & " where nobukti = '" & txt_bon & "'")
        Set rsDis = con.Execute("select * from tbdiskon where nobukti = '" & txt_bon & "'")
        If Val(txt_diskon) > 0 Then
            If rsDis.EOF = True Then
                con.Execute ("insert into tbdiskon values('" & txt_bon & "', '" & dis_spv & "', '" & dis_status & "', '" & dis_cust & "', " & priceToNum(txt_diskon) & ")")
            Else
                con.Execute ("update tbdiskon set supervisor = '" & dis_spv & "', customer = '" & dis_cust & "', status = '" & dis_status & "', nilai = " & priceToNum(txt_diskon) & " where nobukti = '" & txt_bon.Text & "'")
            End If
        Else
            If Not rsDis.EOF Then
                con.Execute ("Delete from tbdiskon where nobukti = '" & txt_bon & "'")
            End If
        End If
    End If
    
    '=========================================================
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0

    Set rsJual = con.Execute("select * from tbjual where nobukti = '" & txt_bon & "'")
    If rsJual.EOF Then
        MsgBox "data tidak ditemukan"
        Exit Sub
    End If
    rsJual.MoveFirst
    Dim cetak_Flag As Boolean
    cetak_Flag = False
    
    If is_new = False Then
        If MsgBox("Cetak struk pembelian?", vbYesNo) = vbYes Then cetak_Flag = True
    End If
    
    'Dim tempFont As String
    'tempFont = InputBox("Nama Font : ", "Font")
    
    If is_new = True Or cetak_Flag = True Then
        Printer.Font = "mingliu"
        'Printer.Font = tempFont
        Printer.FontSize = 18
        Printer.FontBold = True
        Printer.Print Tab(4); Setting_Object("Toko2"); '"KRIPIK BALADO";
        Printer.Print Tab(3); Setting_Object("Toko"); '"CHRISTINE HAKIM";
        'Printer.Print Tab(2); Printer.PaintPicture(App.Path & "\CHIP.jpg");
        'Printer.Print Tab(2); "CHRISTINE HAKIM";
        
'        Printer.PaintPicture LoadPicture(App.Path & "\chip.jpg"), 300, 0, 2774, 1510
'        Printer.Print Tab(1); "                                                                  ";
'        Printer.Print Tab(1); "                                                                  ";
'        Printer.Print Tab(1); "                                                                  ";
'        Printer.Print Tab(1); "                                                                  ";
'        Printer.Print Tab(1); "                                                                  ";
        Printer.Font = "dotumche"
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.Print Tab(1); "                                                            "; 'baris kosong
        Printer.Print Tab(1); Setting_Object("Alamat1");
        Printer.Print Tab(1); "No. FAKTUR : "; txt_bon.Text
        Printer.Print Tab(1); Format(rsBill!tanggal, "dd-MM-yyyy"); "  "; rsBill!jam;
        Printer.Print Tab(1); "Nama Kasir : "; rsBill!kasir;
        Printer.Print Tab(1); "                                                                  ";
        Printer.Print Tab(1); "------------------------------------------------------------------";
        Do While Not rsJual.EOF
            Printer.Print Tab(1); rsJual!nama_barang
            Dim bayar As Long
            bayar = Val(rsJual!jumlah_jual) * Val(rsJual!harga_jual)
            Printer.Print Tab(2); rsJual!jumlah_jual; Tab(9); "x"; Tab(21 - Len(Format(rsJual!harga_jual, "###,###,##0"))); Format(rsJual!harga_jual, "###,###,##0"); Tab(35 - Len(Format(bayar, "###,###,##0"))); Format(bayar, "###,###,##0")
            rsJual.MoveNext
        Loop
        Printer.Print Tab(1); "                                                                  ";
        Printer.FontSize = 10
        If priceToNum(txt_diskon) > 0 Then
            Printer.Print Tab(1); "Total"; Tab(20); "Rp."; Tab(35 - Len(Format(txt_total, "###,###,##0"))); Format(txt_total, "###,###,##0")
            Printer.Print Tab(1); "Diskon"; Tab(20); "Rp."; Tab(35 - Len(Format(txt_diskon, "###,###,##0"))); Format(txt_diskon, "###,###,##0")
        End If
        ''test
        'Printer.Print Tab(1); "Pajak Restoran 10%"; Tab(20); "Rp."; Tab(35 - Len(Format(txt_ppn, "###,###,##0"))); Format(txt_ppn, "###,###,##0")
        Printer.Print Tab(1); "------------------------------------------------------------------";
        Printer.FontSize = 12
        Printer.Print Tab(2); "Grand Total"; Tab(15); "Rp."; Tab(30 - Len(Format(priceToNum(txt_total) - priceToNum(txt_diskon), "###,###,##0"))); Format(priceToNum(txt_total) - priceToNum(txt_diskon), "###,###,##0")
    
    ''    Dim diskon_total As Long
    ''    diskon_total = priceToNum(txt_total) - priceToNum(txt_diskon)
    ''nyelip di sini
        Printer.CurrentX = 0
        Printer.FontSize = 10
        Printer.Print Tab(3); "                                                             ";
        If (tunai = 1) Then
            Printer.Print Tab(1); "Jumlah Uang"; Tab(15); "Rp. "; Tab(31 - Len(Format(txt_uang, "###,###,##0"))); Format(txt_uang, "###,###,##0");
            Printer.Print Tab(1); "Kembalian  "; Tab(15); "Rp. "; Tab(31 - Len(Format(txt_kembali, "###,###,##0"))); Format(txt_kembali, "###,###,##0");
        Else
            Printer.Print Tab(3); "-NON TUNAI-";
        End If
        Printer.Print Tab(1); "                                                             ";
        Printer.FontSize = 10
        
        Printer.Print Tab(2); "Customer Service: (0751)483518";
        Printer.Print Tab(2); "HP Pemesanan: 0811 668 5000";
        Printer.Print Tab(2); "www.christinehakimideapark.com";
        Printer.Print Tab(3); ""
        Printer.Print Tab(1); "                                                             ";
        
        'Printer.FontSize = 10
        Printer.Print Tab((38 - Len("*Periksalah uang kembalian anda*")) / 2); "*Periksalah uang kembalian anda*";
        Printer.Print Tab((38 - Len("*sebelum meninggalkan kasir*")) / 2); "*sebelum meninggalkan kasir*";
    
        Printer.EndDoc
    End If
    
    
    If priceToNum(txt_diskon.Text) > 0 Then
        'con.Execute ("insert into tbdiskon values('" & txt_bon & "', '" & dis_spv & "', '" & dis_status & "', '" & dis_cust & "', " & priceToNum(txt_diskon) & ")")
        If MsgBox("Cetak struk diskon?", vbYesNo) = vbYes Then
            Printer.Font = "dotumche"
            'Printer.Font = tempFont
            Printer.FontSize = 10
            Printer.Print Tab(5); Format(Now, "dd-MM-yyyy  hh:mm:ss");
            Printer.Print Tab(5); "No Faktur"; Tab(19); ":  "; txt_bon
            Printer.Print Tab(5); "Supervisor"; Tab(19); ":  "; dis_spv
            Printer.Print Tab(5); "Status"; Tab(19); ":  "; dis_status
            Printer.Print Tab(5); "Customer"; Tab(19); ":  "; dis_cust
            Printer.Print Tab(5); "Diskon"; Tab(19); ":  Rp. "; Format(txt_diskon, "###,###,##0")
            Print Tab(3); "                                                            ";
            Print Tab(3); "                                                            ";
            Print Tab(3); "                                                            ";
            Printer.EndDoc
        End If
    End If
    

      
    Close #1
    
    
    If is_new Then
        Form_Penjualan.nextFaktur
    Else
        Form_List_Jual.refreshlist
    End If

    Unload Me
    
'    ''pindahan print diskon
'    If priceToNum(txt_diskon.Text) > 0 Then
'       con.Execute ("insert into tbdiskon values('" & txt_bon & "', '" & dis_spv & "', '" & dis_status & "', '" & dis_cust & "', " & priceToNum(txt_diskon) & ")")
'
'        Printer.Font = "Times new roman"
'        Printer.FontSize = 12
'        Printer.Print Tab(4); Format(Now, "dd-MM-yyyy  hh:mm:ss");
'        Printer.Print Tab(4); "No Faktur"; Tab(18); ": "; txt_bon
'        Printer.Print Tab(4); "Supervisor"; Tab(18); ": "; dis_spv
'        Printer.Print Tab(4); "Status"; Tab(18); ": "; dis_status
'        Printer.Print Tab(4); "Customer"; Tab(18); ": "; dis_cust
'        Printer.Print Tab(4); "Diskon"; Tab(18); ": Rp."; txt_diskon
'        Printer.EndDoc
'    End If
'
'    ''end diskon
'
'    Set rsJual = con.Execute("select * from tbjual where nobukti = '" & txt_bon & "'")
'    If rsJual.EOF Then
'        MsgBox "data tidak ditemukan"
'        Exit Sub
'    End If
'    rsJual.MoveFirst
'
'    Printer.Font = "times new roman"
'    Printer.CurrentX = 0
'    Printer.CurrentY = 0
'    Printer.FontSize = 18
'    Printer.FontBold = True
'    Printer.Print Tab(3); " KRIPIK BALADO";
'    Printer.Print Tab(2); "CHRISTINE HAKIM";
'    Printer.FontSize = 10
'    Printer.FontBold = False
'    Printer.Print Tab(3); "                                                            "; 'baris kosong
'    Printer.Print Tab(3); "Jl. Adinegoro No. 11A Padang";
'    Printer.Print Tab(3); "                                                             ";
'    Printer.Print Tab(3); "No. FAKTUR :"; txt_bon.Text
'    Printer.Print Tab(3); "Nama Kasir :"; rsBill!kasir;
'    Printer.Print Tab(3); Format(rsBill!tanggal, "dd-MM-yyyy"); "  "; rsBill!jam;
'    Printer.Print Tab(3); "                                                             ";
'    Do While Not rsJual.EOF
'        Printer.Print Tab(3); rsJual!nama_barang
'        Dim bayar As Long
'        bayar = Val(rsJual!jumlah_jual) * Val(rsJual!harga_jual)
'        Printer.Print Tab(3); rsJual!jumlah_jual; "x"; Format(rsJual!harga_jual, "###,###,###"); Tab(35); Format(bayar, "###,###,###")
'        rsJual.MoveNext
'    Loop
'    Printer.Print Tab(30); "--------------------------";
'    Printer.FontSize = 14
'    If priceToNum(txt_diskon) > 0 Then
'        Printer.Print Tab(8); "Diskon"; Tab(16); ": Rp."; txt_diskon
'    End If
'    Dim diskon_total As Long
'    diskon_total = priceToNum(txt_total) - priceToNum(txt_diskon)
'    Printer.Print Tab(8); "Total"; Tab(16); ": Rp."; Format(diskon_total, "###,###,###")
'    Printer.CurrentX = 0
'    Printer.FontSize = 10
'    Printer.Print Tab(3); "                                                             ";
'    If (tunai = 1) Then
'        Printer.Print Tab(3); "Jumlah Uang  "; Tab(25); Format(txt_uang, "###,###,##0");
'        Printer.Print Tab(3); "Kembalian    "; Tab(25); Format(txt_kembali, "###,###,##0");
'    Else
'        Printer.Print Tab(3); "-NON TUNAI-";
'    End If
'    Printer.Print Tab(3); "                                                             ";
'    Printer.FontSize = 10
'    Printer.Print Tab(3); "                                                             ";
'    Printer.Print Tab(3); "Customer Service: (0751)483518";
'    Printer.Print Tab(3); "HP Pemesanan: 0811 668 5000";
'    Printer.Print Tab(3); "Website: www.christinehakimideapark.com";
'    Printer.Print Tab(3); "                                                             ";
'    Printer.FontSize = 8
'    Printer.Print Tab(3); "*Barang yang sudah dibeli tidak dapat dikembalikan*";
'    Printer.Print Tab(25); "*Terima Kasih*";
'
'Close #1
'Printer.EndDoc
'
'    If is_new Then
'        Form_Penjualan.nextFaktur
'    Else
'        Form_List_Jual.refreshlist
'    End If
'
'    Unload Me
End Sub

Private Sub txt_diskon_Click()
    Me.Enabled = False
    Form_Diskon.Show 1
    Form_Diskon.Top = Me.Top + txt_diskon.Top + 200
    Form_Diskon.Left = Me.Left + txt_diskon.Left
End Sub

Private Sub txt_uang_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        If Len(txt_uang) > 8 Then
            txt_uang = ""
            Exit Sub
        End If
        Dim kembalian As Long
        kembalian = priceToNum(txt_uang) - priceToNum(txt_total.Text) + priceToNum(txt_diskon)
        If kembalian < 0 Then
            MsgBox "Uang tidak cukup"
        Else
            txt_kembali.Text = Format(kembalian, "###,###,##0")
            btn_tunai.SetFocus
        End If
    End If
End Sub

Public Sub diskon_query()
    dis_spv = Form_Diskon.txt_spv.Text
    dis_status = Form_Diskon.cb_status.Text
    dis_cust = Form_Diskon.txt_customer.Text
End Sub

Private Sub txt_uang_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 1)
End Sub
