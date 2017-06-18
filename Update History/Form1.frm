VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Update"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Buat tabel absensi"
      Enabled         =   0   'False
      Height          =   735
      Left            =   480
      TabIndex        =   15
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Buat Database Backup"
      Enabled         =   0   'False
      Height          =   735
      Left            =   480
      TabIndex        =   14
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Perbaiki Database Warning"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2640
      TabIndex        =   13
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buat Tabel dan Views"
      Enabled         =   0   'False
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txt_Pass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   6120
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "TrimData"
      Enabled         =   0   'False
      Height          =   735
      Left            =   10800
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Update4"
      Enabled         =   0   'False
      Height          =   735
      Left            =   8760
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update3"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6720
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update2"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4680
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update1"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lbl_Pass 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   12
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "TrimData - Menghapus data history barang yg telah dihapus"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   5400
      Width           =   10095
   End
   Begin VB.Label Label4 
      Caption         =   "Update4 - Melengkapi data2 yg belum terisi dengan hasil penjualan berikutnya atau dengan harga di tabel barang "
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   4920
      Width           =   10095
   End
   Begin VB.Label Label3 
      Caption         =   "Update3 - Melengkapi data2 yg belum terisi berdasarkan penjualan terakhir"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   4440
      Width           =   10095
   End
   Begin VB.Label Label2 
      Caption         =   "Update2 - Memasukkan data perubahan harga modal dari tbbeli"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3960
      Width           =   10095
   End
   Begin VB.Label Label1 
      Caption         =   "Update1 - Memasukkan data perubahan harga jual dari tbjual"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   3480
      Width           =   10095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection

Private Sub command1_Click()
    If MsgBox("Jalankan update 1?", vbYesNo, "Update 1") = vbYes Then
        Dim rsHistory As ADODB.Recordset
        Dim tempHarga As Long
        Dim tempKode As String
        tempHarga = 0
        tempKode = ""
        Set rsHistory = con.Execute("select * from v_update1")
        Do While Not rsHistory.EOF
            If rsHistory!kode = tempKode And rsHistory!harga_jual <> tempHarga And tempKode <> "" Then
                con.Execute ("insert into tbbarang_history values ('" & tempKode & "','" & rsHistory!nama_barang & "','" & Format(rsHistory!tglbukti, "yyyy-MM-dd") & "','0','" & rsHistory!harga_jual & "')")
                
            End If
            tempKode = rsHistory!kode
            tempHarga = rsHistory!harga_jual
            rsHistory.MoveNext
        Loop
        Set rsHistory = Nothing
    '    rsHistory.Close
        MsgBox "Update 1 berhasil"
    End If
End Sub

Private Sub command2_Click()
    If MsgBox("Jalankan update 2?", vbYesNo, "Update 2") = vbYes Then
        Dim rsHistory As ADODB.Recordset
        Dim rsCekTb As ADODB.Recordset
        Dim tempHarga As Long
        Dim tempKode As String
        Dim tempTanggal As Date
        tempHarga = 0
        tempKode = ""
        tempTanggal = 0
        Set rsHistory = con.Execute("select * from v_update2")
        Do While Not rsHistory.EOF
            If rsHistory!kode = tempKode And rsHistory!harga <> tempHarga And tempKode <> "" Then
                Set rsCekTb = con.Execute("Select * from tbbarang_history where tanggal = '" & Format(rsHistory!tglbukti, "yyyy-MM-dd") & "' and kode = '" & tempKode & "'")
                If rsCekTb.EOF Then
                    con.Execute ("insert into tbbarang_history values ('" & tempKode & "','" & rsHistory!nama_barang & "','" & Format(rsHistory!tglbukti, "yyyy-MM-dd") & "','" & rsHistory!harga & "','0')")
                Else
                    If rsCekTb!harga_modal = 0 Then
                        con.Execute ("update tbbarang_history set harga_modal =  '" & rsHistory!harga & "' where tanggal = '" & Format(rsHistory!tglbukti, "yyyy-MM-dd") & "' and kode = '" & tempKode & "'")
                    Else
                        con.Execute ("insert into tbbarang_history values ('" & tempKode & "','" & rsHistory!nama_barang & "','" & Format(rsHistory!tglbukti, "yyyy-MM-dd") & "','" & rsHistory!harga & "','0')")
                    End If
                End If
            End If
            tempKode = rsHistory!kode
            tempHarga = rsHistory!harga
            rsHistory.MoveNext
        Loop
        Set rsCekTb = Nothing
        Set rsHistory = Nothing
    '    rsHistory.Close
        MsgBox "Update 2 berhasil"
    End If
End Sub

Private Sub command3_Click()
    If MsgBox("Jalankan update 3?", vbYesNo, "Update 3") = vbYes Then
        Dim rsHistory As ADODB.Recordset
        Dim rsTB As ADODB.Recordset
        Set rsHistory = con.Execute("select * from tbbarang_history")
        Do While Not rsHistory.EOF
'            Set rsTB = con.Execute("select * from tbbeli where kode = '" & rsHistory!kode & "' and tglbukti <= '" & rsHistory!tanggal & "' order by tglbukti desc limit 1")
'            If Not rsTB.EOF Then
'                If rsHistory!harga_modal = 0 Then
'                    con.Execute ("update tbbarang_history set harga_modal = '" & rsTB!harga_modal & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
'                ElseIf rsHistory!harga_jual = 0 Then
'                    con.Execute ("update tbbarang_history set harga_jual = '" & rsTB!harga_jual & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
'                End If
'            End If
            If rsHistory!harga_modal = 0 Then
                Set rsTB = con.Execute("select * from tbbeli where kode = '" & rsHistory!kode & "' and tglbukti <= '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' order by tglbukti desc limit 1")
                If Not rsTB.EOF Then
                    con.Execute ("update tbbarang_history set harga_modal = '" & rsTB!harga & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                End If
            ElseIf rsHistory!harga_jual = 0 Then
                Set rsTB = con.Execute("select * from tbjual where kode = '" & rsHistory!kode & "' and tglbukti <= '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' order by tglbukti desc limit 1")
                If Not rsTB.EOF Then
                    con.Execute ("update tbbarang_history set harga_jual = '" & rsTB!harga_jual & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                End If
            End If
            rsHistory.MoveNext
        Loop
        Set rsTB = Nothing
        Set rsHistory = Nothing
        MsgBox "Update 3 berhasil"
    End If
End Sub

Private Sub command4_Click()
    If MsgBox("Jalankan update 4?", vbYesNo, "Update 4") = vbYes Then
        Dim rsHistory As ADODB.Recordset
        Dim rsTB As ADODB.Recordset
        Set rsHistory = con.Execute("select * from tbbarang_history")
        Do While Not rsHistory.EOF
'            Set rsTB = con.Execute("select * from tbbeli where kode = '" & rsHistory!kode & "' and tglbukti <= '" & rsHistory!tanggal & "' order by tglbukti desc limit 1")
'            If Not rsTB.EOF Then
'                If rsHistory!harga_modal = 0 Then
'                    con.Execute ("update tbbarang_history set harga_modal = '" & rsTB!harga_modal & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
'                ElseIf rsHistory!harga_jual = 0 Then
'                    con.Execute ("update tbbarang_history set harga_jual = '" & rsTB!harga_jual & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
'                End If
'            End If
            If rsHistory!harga_modal = 0 Then
                Set rsTB = con.Execute("select * from tbbeli where kode = '" & rsHistory!kode & "' and tglbukti >= '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' order by tglbukti asc limit 1")
                If Not rsTB.EOF Then
                    con.Execute ("update tbbarang_history set harga_modal = '" & rsTB!harga & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                Else
                    Set rsTB = con.Execute("Select * from tbbarang where kode = '" & rsHistory!kode & "'")
                    If Not rsTB.EOF Then
                        con.Execute ("update tbbarang_history set harga_modal = '" & rsTB!harga_modal & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                    End If
                End If
            ElseIf rsHistory!harga_jual = 0 Then
                Set rsTB = con.Execute("select * from tbjual where kode = '" & rsHistory!kode & "' and tglbukti >= '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' order by tglbukti asc limit 1")
                If Not rsTB.EOF Then
                    con.Execute ("update tbbarang_history set harga_jual = '" & rsTB!harga_jual & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                Else
                    Set rsTB = con.Execute("Select * from tbbarang where kode = '" & rsHistory!kode & "'")
                    If Not rsTB.EOF Then
                        con.Execute ("update tbbarang_history set harga_jual = '" & rsTB!harga_jual & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                    End If
                End If
            End If
            rsHistory.MoveNext
        Loop
        Set rsTB = Nothing
        Set rsHistory = Nothing
        MsgBox "Update 4 berhasil"
    End If
End Sub


Private Sub command5_Click()
    If MsgBox("Jalankan Trim Data?", vbYesNo, "Trim Data") = vbYes Then
        Dim rsHistory As ADODB.Recordset
        Dim rsTB As ADODB.Recordset
        Dim countDeleted As Integer
        countDeleted = 0
        Set rsHistory = con.Execute("Select * from tbbarang_history")
        Do While Not rsHistory.EOF
            Set rsTB = con.Execute("Select * from tbbarang where kode = '" & rsHistory!kode & "'")
            If rsTB.EOF Then
                con.Execute ("Delete from tbbarang_history where kode = '" & rsHistory!kode & "'")
                countDeleted = countDeleted + 1
            End If
            rsHistory.MoveNext
        Loop
        Set rsHistory = Nothing
        Set rsTB = Nothing
        MsgBox (countDeleted & " data dihapus")
    End If
End Sub

Private Sub Command6_Click()
    con.Execute ("CREATE TABLE `tbbarang_history` (`kode` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,`nama` CHAR(50) COLLATE utf8_general_ci DEFAULT NULL,`tanggal` DATE DEFAULT NULL,`harga_modal` INTEGER(11) DEFAULT NULL,`harga_jual` INTEGER(11) DEFAULT NULL) ENGINE=InnoDB CHARACTER SET 'utf8' COLLATE 'utf8_general_ci'")
    con.Execute ("CREATE ALGORITHM=UNDEFINED DEFINER='root'@'localhost' SQL SECURITY DEFINER VIEW `v_update1` AS select `a`.`nobukti` AS `nobukti`, `a`.`tglbukti` AS `tglbukti`, `a`.`kode` AS `kode`, `a`.`nama_barang` AS `nama_barang`, `a`.`harga_jual` AS `harga_jual`, `a`.`jumlah_jual` AS `jumlah_jual` From `tbjual` `a` Where (left (`a`.`nobukti`, 2) <> 'RS') Group By `a`.`kode`,`a`.`tglbukti`,`a`.`harga_jual`")
    con.Execute ("CREATE ALGORITHM=UNDEFINED DEFINER='root'@'localhost' SQL SECURITY DEFINER VIEW `v_update2` AS select `a`.`nobukti` AS `nobukti`, `a`.`tglbukti` AS `tglbukti`, `a`.`kode` AS `kode`, `a`.`nama_barang` AS `nama_barang`, `a`.`harga` AS `harga`, `a`.`jumlah` AS `jumlah`, `a`.`return` AS `return` From `tbbeli` `a` Group By `a`.`kode`, `a`.`tglbukti`, `a`.`harga`")
    MsgBox ("Database dan Views berhasil dibuat")
End Sub

Private Sub Command7_Click()
    Dim rsBarang, rsBeli As ADODB.Recordset
    Set rsBeli = con.Execute("select * from tbbeli where tglbukti >= '2017-05-01' order by tglbukti desc")
    If Not rsBeli.EOF Then
        Do While Not rsBeli.EOF
        Set rsBarang = con.Execute("Select * from tbbarang where kode = '" & rsBeli!kode & "'")
        If Not rsBarang.EOF Then
            If rsBarang!tgl_masuk < rsBeli!tglbukti Then
                If MsgBox("Replace tgl_masuk : " & rsBarang!tgl_masuk & vbNewLine & "Dengan tglbukti : " & rsBeli!tglbukti & vbNewLine & "Untuk Barang : " & rsBarang!kode & vbNewLine & rsBarang!nama, vbYesNo, "Konfirmasi") = vbYes Then
                    con.Execute ("Insert into backup values('" & rsBarang!kode & "','" & rsBarang!nama & "','" & Format(rsBarang!tgl_masuk, "yyyy-mm-dd") & "','" & Format(rsBeli!tglbukti, "yyyy-mm-dd") & "')")
                    con.Execute ("Update tbbarang set tgl_masuk = '" & Format(rsBeli!tglbukti, "yyyy-mm-dd") & "' where kode = '" & rsBarang!kode & "'")
                End If
            End If
        End If
        rsBeli.MoveNext
        Loop
    End If
    Set rsBarang = Nothing
    Set rsBeli = Nothing
    MsgBox ("Done")
End Sub

Private Sub Command8_Click()
    con.Execute ("CREATE TABLE `backup` (`kode` CHAR(20) COLLATE utf8_general_ci NOT NULL, `nama` CHAR(100) COLLATE utf8_general_ci DEFAULT NULL, `tgl_masuk` DATE DEFAULT NULL, `tglbukti` DATE DEFAULT NULL ) ENGINE=InnoDB CHARACTER SET 'utf8' COLLATE 'utf8_general_ci';")
    MsgBox ("Done")
End Sub

Private Sub Command9_Click()
    con.Execute ("CREATE TABLE `tbabsen` ( `userid` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL, `tanggal` DATE DEFAULT NULL, `jam_masuk` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL, `jam_keluar` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL ) ENGINE=InnoDB CHARACTER SET 'utf8' COLLATE 'utf8_general_ci' ;")
    MsgBox ("Done")
End Sub

Private Sub Form_Load()
    con.ConnectionString = "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
    con.Open
End Sub

Private Sub Form_Unload(Cancel As Integer)
    con.Close
End Sub

Private Sub txt_Pass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txt_Pass.Text = "chip456" Then
            Command1.Enabled = True
            Command2.Enabled = True
            Command3.Enabled = True
            Command4.Enabled = True
            Command5.Enabled = True
            Command6.Enabled = True
            Command7.Enabled = True
            Command8.Enabled = True
            Command9.Enabled = True
            txt_Pass.Visible = False
            lbl_Pass.Visible = False
            Command6.SetFocus
        End If
    End If
End Sub
