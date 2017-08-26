VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Tambah_Retur 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form Pengembalian"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   15720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_kode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Text            =   "12345678901234"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txt_nama 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Text            =   "12345678901234"
      Top             =   1200
      Width           =   8415
   End
   Begin VB.TextBox txt_jumlah 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   614
      Left            =   13440
      TabIndex        =   4
      Text            =   "12345678901234"
      Top             =   1200
      Width           =   1239
   End
   Begin MSComctlLib.ListView list_nama 
      Height          =   2295
      Left            =   4440
      TabIndex        =   3
      Top             =   1800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4048
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode"
         Object.Width           =   2976
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama"
         Object.Width           =   7440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Harga"
         Object.Width           =   2976
      EndProperty
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   615
      Left            =   14640
      TabIndex        =   2
      Top             =   1200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "list_nama"
      BuddyDispid     =   196615
      OrigLeft        =   19062
      OrigTop         =   2863
      OrigRight       =   19364
      OrigBottom      =   3596
      Max             =   9999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ListView lv_Retur 
      Height          =   5775
      Left            =   600
      TabIndex        =   5
      Top             =   2400
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode Barang"
         Object.Width           =   4464
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   16140
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Jumlah"
         Object.Width           =   4464
      EndProperty
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form_Tambah_Retur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsbarang As ADODB.Recordset
Dim txt_nama_Toggle As Boolean

'Private Sub Command1_Click()
'    Dim i As Integer
'    For i = 1 To lv_Retur.ColumnHeaders.count
'        MsgBox lv_Retur.ColumnHeaders(i).Width
'    Next
'End Sub

Private Sub Form_Load()
    kosongkan
    txt_nama_Toggle = False
    Set rsbarang = con.Execute("select * from tbbarang")
    reload_List
End Sub

Private Sub Form_KeyDown(key As Integer, Shift As Integer)
    If key = 112 Then
        If lv_Retur.ListItems.count > 0 Then
'            Form_Print.Show
'            Form_Print.Init lbl_faktur, txt_total, True
'            Me.Enabled = False
            If MsgBox("Simpan data?", vbYesNo, "Konfirmasi") = vbYes Then
                simpanData
                Form_Retur.refreshlist
                Unload Me
            End If
        Else
            MsgBox "Faktur masih kosong"
        End If
    End If
    
    If key = 46 Then
        If Shift = 1 Then
'            txt_total = "0"
            lv_Retur.ListItems.Clear
        Else
'            txt_total = Format(priceToNum(txt_total) - priceToNum(lv_Retur.SelectedItem.SubItems(4)), "###,###,##0")
            lv_Retur.ListItems.Remove (lv_Retur.SelectedItem.index)
        End If
    End If
    If key = 115 Then
        If MsgBox("Tutup form transaksi?", vbYesNo) = vbYes Then
            Unload Me
        End If
    End If
End Sub

Private Sub kosongkan()
    txt_kode.Text = ""
    txt_nama.Text = ""
    txt_jumlah.Text = 1
    list_nama.Visible = False
End Sub

Private Sub list_nama_lostfocus()
    list_nama.Visible = False
End Sub

Private Sub list_nama_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        list_nama_DblClick
    End If
End Sub

Private Sub txt_jumlah_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 1)
End Sub

Private Sub txt_kode_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub txt_nama_Change()
    If txt_nama.Text <> "" And txt_nama_Toggle = False Then
        list_nama.Visible = True
        reload_List
    Else
        list_nama.Visible = False
        txt_nama_Toggle = False
    End If
End Sub

Private Sub txt_nama_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 3)
End Sub

Private Sub txt_nama_LostFocus()
    If Not Me.ActiveControl Is Nothing Then
        If Not Me.ActiveControl.Name = "list_nama" Then
            list_nama.Visible = False
        End If
    End If
End Sub

Private Sub list_nama_DblClick()
    If getItemByID(list_nama.SelectedItem.Text) Then
        txt_kode.Text = rsbarang!kode
        txt_nama.Text = rsbarang!Nama
        list_nama.Visible = False
        txt_jumlah.SetFocus
        txt_jumlah.SelLength = Len(txt_jumlah.Text)
    End If
End Sub


Private Sub txt_jumlah_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        If Len(txt_jumlah) > 4 Then
            txt_jumlah = ""
            Exit Sub
        End If
    
'        If txt_harga = "" Then
'            MsgBox "Barang tidak valid"
'            Exit Sub
'        End If
        
        If Val(txt_jumlah.Text) < 1 Then
            MsgBox "Jumlah tidak valid"
            Exit Sub
        End If
        
        Dim found As Boolean
        Dim i As Integer
        found = False
        i = 1
        
        Do While i <= lv_Retur.ListItems.count
            If lv_Retur.ListItems(i).Text = rsbarang!kode Then
                found = True
                lv_Retur.ListItems(i).SubItems(2) = Val(lv_Retur.ListItems(i).SubItems(2)) + Val(txt_jumlah.Text)
                Exit Do
            End If
            i = i + 1
        Loop
        
        If found = False Then
            Dim item As ListItem
            Set item = lv_Retur.ListItems.Add(, , rsbarang!kode)
            item.SubItems(1) = rsbarang!Nama
            item.SubItems(2) = txt_jumlah.Text
        End If
        
        kosongkan
        reload_List
        txt_kode.SetFocus
    End If
End Sub

Private Sub txt_kode_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        txt_nama_Toggle = True
        Dim kode As String
        kode = Trim(txt_kode.Text)
        If getItemByID(kode) Then
            txt_nama.Text = rsbarang!Nama
            txt_jumlah.SetFocus
            txt_jumlah.SelLength = Len(txt_jumlah.Text)
        Else
            MsgBox ("Kode ini tidak terdaftar")
        End If
    ElseIf Len(txt_nama) > 0 Then
        txt_nama = ""
    End If
End Sub

Private Function getItemByID(kode As String) As Boolean
    rsbarang.MoveFirst
    Do While Not rsbarang.EOF
        If rsbarang!kode = kode Then
            getItemByID = True
            Exit Function
        End If
        rsbarang.MoveNext
    Loop
    rsbarang.MoveFirst
    getItemByID = False
End Function

Private Sub txt_nama_KeyDown(key As Integer, Shift As Integer)
    If key = 40 Then
        list_nama.Visible = True
        list_nama.SetFocus
        'Exit Sub
    ElseIf key = 13 And list_nama.Visible = True Then
        list_nama.SetFocus
    End If
    
    'pindah ke txt_nama_change
'    list_nama.ListItems.Clear
'    list_nama.Visible = True
'    Dim rsFilter As ADODB.Recordset
'    Set rsFilter = con.Execute("select * from tbbarang where nama like '%" & txt_nama.Text & "%'")
'
'    If rsFilter.EOF Then
'        list_nama.Visible = False
'        Exit Sub
'    End If
'
'    rsFilter.MoveFirst
'    Do While Not rsFilter.EOF
'        Dim mitem As ListItem
'        Set mitem = list_nama.ListItems.Add(, , rsFilter!kode)
'        mitem.SubItems(1) = rsFilter!nama
'        mitem.SubItems(2) = "Rp. " + Format(rsFilter!harga_jual, "###,###,##0")
'        rsFilter.MoveNext
'    Loop
'
'    Set rsFilter = Nothing
End Sub

Public Sub reload_List()
'pindahan generate list barang
    list_nama.ListItems.Clear
    'list_nama.Visible = True
    Dim rsFilter As ADODB.Recordset
    Set rsFilter = con.Execute("select * from tbbarang where nama like '%" & txt_nama.Text & "%'")
    
    If rsFilter.EOF Then
        list_nama.Visible = False
        Exit Sub
    End If
    
    rsFilter.MoveFirst
    Do While Not rsFilter.EOF
        Dim mitem As ListItem
        Set mitem = list_nama.ListItems.Add(, , rsFilter!kode)
        mitem.SubItems(1) = rsFilter!Nama
        mitem.SubItems(2) = "Rp. " + Format(rsFilter!harga_jual, "###,###,##0")
        rsFilter.MoveNext
    Loop
    
    Set rsFilter = Nothing
'end pindahan list barang
End Sub

Private Sub simpanData()
    Dim i  As Integer
    For i = 1 To lv_Retur.ListItems.count
        Call prosesSimpan(lv_Retur.ListItems(i).Text, lv_Retur.ListItems(i).SubItems(1), lv_Retur.ListItems(i).SubItems(2))
    Next
End Sub

Private Sub prosesSimpan(inKode As String, inNama As String, inJumlah As Integer)
    Dim rsRetur As ADODB.Recordset
    Set rsRetur = con.Execute("select * from tbretur where kode = '" & inKode & "'")
    'cek database
    If rsRetur.EOF Then
    'not exist insert
        con.Execute ("insert into tbretur(kode,nama,tgl_retur,userid,jumlah) values ('" & inKode & "','" & inNama & "','" & Format(Date, "yyyy-mm-dd") & "','" & username & "','" & inJumlah & "')")
    Else
    'exist update
        'con.Execute ("update tbretur set tgl_retur = '" & Format(Date, "yyyy-mm-dd") & "', userid = '" & username & "', jumlah = '" & inJumlah + rsRetur!jumlah & "' where kode = '" & inKode & "'")
        If Len(inJumlah + rsRetur!jumlah) < 5 Then
            con.Execute ("update tbretur set jumlah = '" & inJumlah + rsRetur!jumlah & "' where kode = '" & inKode & "'")
        Else
            MsgBox "Barang dengan kode " & inKode & " melebihi 9999.", vbOKOnly, "Update gagal"
        End If
    End If
    Set rsRetur = Nothing
End Sub
