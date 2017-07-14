VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form_Penjualan 
   BackColor       =   &H0000C000&
   Caption         =   "Transaksi Penjualan"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form11"
   ScaleHeight     =   13046.83
   ScaleMode       =   0  'User
   ScaleWidth      =   24147.85
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_Diskon 
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
      Left            =   17880
      TabIndex        =   14
      Text            =   "12345678901234"
      Top             =   2400
      Width           =   1239
   End
   Begin MSComctlLib.ListView list_nama 
      Height          =   2295
      Left            =   4440
      TabIndex        =   20
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
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   615
      Left            =   17040
      TabIndex        =   17
      Top             =   2400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt_jumlah"
      BuddyDispid     =   196609
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
   Begin MSCommLib.MSComm MSComm1 
      Left            =   12360
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.ListView lv_jual 
      Height          =   6855
      Left            =   360
      TabIndex        =   15
      Top             =   3360
      Width           =   19515
      _ExtentX        =   34422
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
         Object.Width           =   13018
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
         Object.Width           =   4464
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Diskon"
         Object.Width           =   2130
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
      Left            =   15840
      TabIndex        =   13
      Text            =   "12345678901234"
      Top             =   2400
      Width           =   1239
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
      Left            =   13200
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
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "Diskon"
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
      Left            =   17880
      TabIndex        =   21
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FF80&
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
      TabIndex        =   16
      Top             =   10320
      Width           =   15015
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
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
      Left            =   15960
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
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
      Left            =   13200
      TabIndex        =   10
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      Left            =   8400
      TabIndex        =   5
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "Kasir:"
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
      Left            =   6960
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
      Caption         =   "No. Faktur:"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FF80&
      Height          =   1695
      Left            =   360
      TabIndex        =   18
      Top             =   1560
      Width           =   19515
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Height          =   1095
      Left            =   360
      TabIndex        =   19
      Top             =   120
      Width           =   19275
   End
End
Attribute VB_Name = "Form_Penjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsbarang As ADODB.Recordset
Dim txt_nama_Toggle As Boolean
Public sumDiskon As Long

Private Sub Form_Load()
    lbl_user = username
    txt_total = 0
    sumDiskon = 0
    kosongkan
    txt_nama_Toggle = False
    Set rsbarang = con.Execute("select * from tbbarang")
        
    Dim namafile, file_data, huruf As String
    Dim angka As Long
    namafile = App.Path & "\faktur.txt"
    Open namafile For Input As #1
    While Not EOF(1)
        Input #1, data
        file_data = data
        huruf = Left(file_data, 1)
        angka = Val(Mid(file_data, 2, 20))
        lbl_faktur = huruf + CStr(angka + 1)
    Wend
Close #1

    On Error GoTo ErrorFound
    MSComm1.CommPort = 3
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.PortOpen = True
    
    If MSComm1.PortOpen Then
        MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
        MSComm1.Output = "Selamat Datang      Kasir: " + username
    End If
ErrorFound:
        'nothing happens
    On Error GoTo 0
    
    reload_List

End Sub

Private Sub Form_unload(cancel As Integer)
     If MSComm1.PortOpen = True Then
      Do While MSComm1.OutBufferCount > 0
          DoEvents
       Loop
       MSComm1.PortOpen = False
    End If
End Sub

Private Sub Form_KeyDown(key As Integer, Shift As Integer)
    If key = 112 Then
        If lv_jual.ListItems.count > 0 Then
            Form_Print.Show
            Form_Print.Init lbl_faktur, txt_total, True
            Me.Enabled = False
            If MSComm1.PortOpen Then
                MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
                MSComm1.Output = "Total Belanja:      " + txt_total.Text
            End If
        Else
            MsgBox "Faktur masih kosong"
        End If
    End If
    
    If key = 46 Then
        If Shift = 1 Then
            txt_total = "0"
            sumDiskon = 0
            lv_jual.ListItems.Clear
        Else
            txt_total = Format(priceToNum(txt_total) - priceToNum(lv_jual.SelectedItem.SubItems(5)), "###,###,##0")
            sumDiskon = sumDiskon - priceToNum(lv_jual.SelectedItem.SubItems(4))
            lv_jual.ListItems.Remove (lv_jual.SelectedItem.index)
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
    txt_jumlah.Text = 1
    txt_Diskon.Text = 0
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

Private Sub txt_Diskon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Len(txt_jumlah) > 4 Then
            txt_jumlah = ""
            Exit Sub
        End If
        
        If txt_Diskon.Text > 100 Or txt_Diskon.Text < 0 Then
            txt_Diskon.Text = "0"
            Exit Sub
        End If
    
        If txt_harga = "" Then
            MsgBox "Barang tidak valid"
            Exit Sub
        End If
        
        If Val(txt_jumlah.Text) < 1 Then
            MsgBox "Jumlah tidak valid"
            Exit Sub
        End If
        
        Dim found As Boolean
        Dim i As Integer
        found = False
        i = 1
        
        Dim subtotal, diskon As String
        diskon = rsbarang!harga_jual * Val(txt_jumlah) * Val(txt_Diskon) / 100
        sumDiskon = sumDiskon + CLng(diskon)
        subtotal = rsbarang!harga_jual * Val(txt_jumlah) - CLng(diskon)
        
        Do While i <= lv_jual.ListItems.count
            If lv_jual.ListItems(i).Text = rsbarang!kode Then
                found = True
                lv_jual.ListItems(i).SubItems(3) = Val(lv_jual.ListItems(i).SubItems(3)) + Val(txt_jumlah.Text)
                lv_jual.ListItems(i).SubItems(4) = priceToNum(lv_jual.ListItems(i).SubItems(4)) + diskon
                lv_jual.ListItems(i).SubItems(4) = Format(lv_jual.ListItems(i).SubItems(4), "###,###,##0")
                lv_jual.ListItems(i).SubItems(5) = priceToNum(lv_jual.ListItems(i).SubItems(5)) + subtotal
                lv_jual.ListItems(i).SubItems(5) = Format(lv_jual.ListItems(i).SubItems(5), "###,###,##0")
                Exit Do
            End If
            i = i + 1
        Loop
        
        If found = False Then
            Dim item As ListItem
            Set item = lv_jual.ListItems.Add(, , rsbarang!kode)
            item.SubItems(1) = rsbarang!Nama
            item.SubItems(2) = Format(rsbarang!harga_jual, "###,###,##0")
            item.SubItems(3) = txt_jumlah.Text
            item.SubItems(4) = Format(diskon, "###,###,##0")
            item.SubItems(5) = Format(subtotal, "###,###,##0")
        End If
        
        txt_total.Text = Format(priceToNum(txt_total) + CLng(subtotal), "###,###,##0")
        
        If MSComm1.PortOpen Then
            MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
            Dim baris1, baris2 As String
            baris1 = txt_jumlah.Text + " " + txt_nama.Text
            If Len(baris1) < 20 Then
               Do While (Len(baris1) < 20)
                baris1 = baris1 + " "
               Loop
            Else
                baris1 = Left$(baris1, 20)
            End If
            
            MSComm1.Output = baris1
              
            Dim spaces As Integer
            spaces = 20 - (Len(subtotal) + Len(txt_total.Text) + 2)
            Do While (Len(baris2) < spaces)
                baris2 = baris2 + " "
            Loop
            baris2 = subtotal + baris2 + "(" + txt_total.Text + ")"
            MSComm1.Output = baris2
        End If
        
        kosongkan
        reload_List
        txt_kode.SetFocus
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
        txt_harga.Text = Format(rsbarang!harga_jual, "###,###,##0")
        list_nama.Visible = False
        txt_jumlah.SetFocus
        txt_jumlah.SelLength = Len(txt_jumlah.Text)
    End If
End Sub


Private Sub txt_jumlah_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        txt_Diskon.SetFocus
    End If
End Sub

Private Sub txt_kode_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        txt_nama_Toggle = True
        Dim kode As String
        kode = Trim(txt_kode.Text)
        If getItemByID(kode) Then
            txt_nama.Text = rsbarang!Nama
            txt_harga.Text = Format(rsbarang!harga_jual, "###,###,##0")
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

Public Sub nextFaktur()
    Dim namafile, huruf As String
    Dim angka As Long
    Me.Enabled = True
    huruf = Left(lbl_faktur, 1)
    angka = Val(Mid(lbl_faktur, 2, 20))
    
    namafile = App.Path & "\faktur.txt"
    Open namafile For Output As #1
    Print #1, lbl_faktur
    Close #1
    
    lbl_faktur = huruf + CStr(angka + 1)
    lv_jual.ListItems.Clear
    txt_total = "0"
    sumDiskon = 0
    kosongkan
    txt_kode.SetFocus
    Form_List_Jual.refreshlist
    
    If MSComm1.PortOpen Then
        MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
        MSComm1.Output = "Selamat Datang      Kasir: " + username
    End If
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
