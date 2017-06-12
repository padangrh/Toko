VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Pembelian 
   BackColor       =   &H0080FFFF&
   Caption         =   "Transaksi Pembelian"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form11"
   ScaleHeight     =   13064.73
   ScaleMode       =   0  'User
   ScaleWidth      =   42894.2
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView list_nama 
      Height          =   2295
      Left            =   4440
      TabIndex        =   22
      Top             =   3000
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
   Begin VB.TextBox txt_return 
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
      Left            =   18000
      TabIndex        =   21
      Text            =   "12345678901234"
      Top             =   2400
      Width           =   1023
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   615
      Left            =   17280
      TabIndex        =   16
      Top             =   2400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt_jumlah"
      BuddyDispid     =   196610
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
   Begin MSComctlLib.ListView lv_beli 
      Height          =   6855
      Left            =   360
      TabIndex        =   14
      Top             =   3360
      Width           =   19275
      _ExtentX        =   33999
      _ExtentY        =   12091
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode Barang"
         Object.Width           =   4464
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   14879
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Harga"
         Object.Width           =   4464
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Jumlah"
         Object.Width           =   2232
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Return"
         Object.Width           =   2232
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   5208
      EndProperty
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
      Left            =   16080
      TabIndex        =   13
      Text            =   "12345678901234"
      Top             =   2400
      Width           =   1453
   End
   Begin VB.TextBox txt_harga 
      Enabled         =   0   'False
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
      Left            =   13320
      TabIndex        =   12
      Text            =   "12345678901234"
      Top             =   2400
      Width           =   2295
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
      TabIndex        =   9
      Text            =   "12345678901234"
      Top             =   2400
      Width           =   8415
   End
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
      Left            =   600
      TabIndex        =   6
      Text            =   "12345678901234"
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox txt_total 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15360
      TabIndex        =   3
      Text            =   "00.000.000"
      Top             =   240
      Width           =   4095
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   614
      Left            =   19020
      TabIndex        =   20
      Top             =   2400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt_return"
      BuddyDispid     =   196609
      OrigLeft        =   19062
      OrigTop         =   2863
      OrigRight       =   19364
      OrigBottom      =   3596
      Max             =   9999
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000C0C0&
      Caption         =   "Return"
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
      Left            =   18000
      TabIndex        =   19
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "[F1 -> Print]  [Delete -> Hapus 1 baris]  [Shift + Delete -> Hapus Semua]  [F4 -> Tutup]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   10320
      Width           =   15015
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
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
      Left            =   16080
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "Harga"
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
      Left            =   13320
      TabIndex        =   10
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
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
      Left            =   4440
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
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
      Left            =   600
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lbl_user 
      BackColor       =   &H0000C0C0&
      Caption         =   "Richard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   5
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "Staff :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14280
      TabIndex        =   2
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label lbl_faktur 
      BackColor       =   &H0000C0C0&
      Caption         =   "A123456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "No. Faktur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Height          =   1695
      Left            =   360
      TabIndex        =   17
      Top             =   1560
      Width           =   19275
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C0C0&
      Height          =   1095
      Left            =   360
      TabIndex        =   18
      Top             =   120
      Width           =   19275
   End
End
Attribute VB_Name = "Form_Pembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsbarang As ADODB.Recordset
Public kode_supplier As String
Dim txt_nama_Toggle As Boolean

Private Sub Form_Load()
    lbl_user = username
    txt_total = 0
    kode_supplier = ""
    resetFaktur
    txt_nama_Toggle = False
    
    Set rsbarang = con.Execute("select * from tbbarang")

    Dim namafile, file_data, huruf As String
    Dim angka As Long
    namafile = App.Path & "\fakturbeli.txt"
    Open namafile For Input As #1
    While Not EOF(1)
        Input #1, data
        file_data = data
        huruf = Left(file_data, 1)
        angka = Val(Mid(file_data, 2, 20))
        lbl_faktur = huruf + CStr(angka + 1)
    Wend
Close #1

End Sub


Private Sub Form_KeyDown(key As Integer, Shift As Integer)
    If key = 112 Then
        If lv_beli.ListItems.count > 0 Then
            Form_Print_Beli.Show
            Form_Print_Beli.Init lbl_faktur, txt_total, True
            Me.Enabled = False
        Else
            MsgBox "Faktur masih kosong"
        End If
    End If
    If key = 46 Then
        If Shift = 1 Then
            txt_total = "0"
            lv_beli.ListItems.Clear
        Else
            txt_total = Format(priceToNum(txt_total) - priceToNum(lv_beli.SelectedItem.SubItems(5)), "###,###,##0")
            lv_beli.ListItems.Remove (lv_beli.SelectedItem.index)
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
    txt_harga.Text = ""
    txt_return.Text = 0
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
    Select Case KeyAscii
        Case 48 To 57, 45, 8 ' 0-9, minus and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub txt_kode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
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
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 45, 47, 8 ' A-Z, 0-9, a-z, minus, slash and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub txt_nama_LostFocus()
     If Not Me.ActiveControl.Name = "list_nama" Then
        list_nama.Visible = False
    End If
End Sub

Private Sub list_nama_DblClick()
    If getItemByID(list_nama.SelectedItem.Text) Then
        txt_kode.Text = rsbarang!kode
        txt_nama.Text = rsbarang!Nama
        txt_harga.Text = Format(rsbarang!harga_modal, "###,###,###")
        list_nama.Visible = False
        txt_jumlah.SetFocus
        txt_jumlah.SelLength = Len(txt_jumlah.Text)
    End If
End Sub

Private Sub txt_jumlah_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        If Val(txt_jumlah.Text) < 0 Then
            MsgBox "Jumlah tidak valid"
            txt_jumlah = ""
            Exit Sub
        End If
        
        txt_return.SetFocus
        txt_return.SelLength = Len(txt_return)
    End If
End Sub

Private Sub txt_return_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        If txt_harga = "" Then
            MsgBox "Barang tidak valid"
            Exit Sub
        End If
        
        Dim found As Boolean
        Dim i As Integer
        found = False
        i = 1
        
        Do While i <= lv_beli.ListItems.count
            If lv_beli.ListItems(i).Text = rsbarang!kode Then
                found = True
                lv_beli.ListItems(i).SubItems(3) = Val(lv_beli.ListItems(i).SubItems(3)) + Val(txt_jumlah.Text)
                lv_beli.ListItems(i).SubItems(4) = Val(lv_beli.ListItems(i).SubItems(4)) + Val(txt_return.Text)
                lv_beli.ListItems(i).SubItems(5) = priceToNum(lv_beli.ListItems(i).SubItems(5)) + (Val(txt_jumlah.Text) - Val(txt_return.Text)) * priceToNum(txt_harga)
                lv_beli.ListItems(i).SubItems(5) = Format(lv_beli.ListItems(i).SubItems(5), "###,###,###")
                Exit Do
            End If
            i = i + 1
        Loop
        
        Dim subtotal As String
        subtotal = Format(rsbarang!harga_modal * (Val(txt_jumlah) - Val(txt_return)), "###,###,###")
        
        If found = False Then
            Dim item As ListItem
            Set item = lv_beli.ListItems.Add(, , rsbarang!kode)
            item.SubItems(1) = rsbarang!Nama
            item.SubItems(2) = Format(rsbarang!harga_modal, "###,###,###")
            item.SubItems(3) = txt_jumlah.Text
            item.SubItems(4) = txt_return.Text
            item.SubItems(5) = subtotal
        End If
        
        txt_total.Text = Format(priceToNum(txt_total) + priceToNum(subtotal), "###,###,###")
        If kode_supplier = "" Then
            kode_supplier = rsbarang!kdsuplier
        End If
        kosongkan
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
            txt_harga.Text = Format(rsbarang!harga_modal, "###,###,###")
            txt_jumlah.SetFocus
            txt_jumlah.SelLength = Len(txt_jumlah.Text)
        Else
            MsgBox ("Kode ini tidak terdaftar")
        End If
    ElseIf Len(txt_nama) > 0 Then
        txt_nama = ""
        txt_harga = ""
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
    getItemByID = False
End Function

Private Sub txt_nama_KeyDown(key As Integer, Shift As Integer)
     If key = 40 Then
        list_nama.Visible = True
        list_nama.SetFocus
    ElseIf key = 13 And list_nama.Visible = True Then
        list_nama.SetFocus
    End If

'    If Len(txt_kode) > 0 Then
'        txt_kode = ""
'        txt_harga = ""
'    End If
'
'    list_nama.ListItems.Clear
'    list_nama.Visible = True
'    Dim rsfilter As ADODB.Recordset
'    Set rsfilter = con.Execute("select * from tbbarang where nama like '%" & txt_nama.Text & "%'")
'
'    If rsfilter.EOF Then
'        Exit Sub
'    End If
'
'    rsfilter.MoveFirst
'    Do While Not rsfilter.EOF
'        Dim mitem As ListItem
'        Set mitem = list_nama.ListItems.Add(, , rsfilter!kode)
'        mitem.SubItems(1) = rsfilter!nama
'        mitem.SubItems(2) = "Rp. " + Format(rsfilter!harga_modal, "###,###,###")
'        rsfilter.MoveNext
'    Loop
'
'    Set rsfilter = Nothing
End Sub

Public Sub nextFaktur()
    Me.Enabled = True
    Dim namafile, huruf As String
    Dim angka As Integer
    
    huruf = Left(lbl_faktur, 1)
    angka = Val(Mid(lbl_faktur, 2, 20))
    
    namafile = App.Path & "\fakturbeli.txt"
    Open namafile For Output As #1
    Print #1, lbl_faktur
    Close #1
    
    lbl_faktur = huruf + CStr(angka + 1)
    resetFaktur
    txt_kode.SetFocus
    Form_List_beli.refreshlist
End Sub

Private Sub resetFaktur()
    txt_total = "0"
    lv_beli.ListItems.Clear
    kode_supplier = ""
    kosongkan
End Sub

Public Sub reload_List()
    list_nama.ListItems.Clear
    Dim rsFilter As ADODB.Recordset
    Set rsFilter = con.Execute("Select * from tbbarang where nama like '%" & txt_nama.Text & "%'")
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
End Sub

Private Sub txt_return_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 45, 8 ' 0-9, minus and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub
