VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form_List_beli 
   Caption         =   "Pembelian"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11940
   ControlBox      =   0   'False
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   108
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":701E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":8052
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":8F37
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":9BB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":A590
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":B056
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20010
      _ExtentX        =   35295
      _ExtentY        =   1720
      BandCount       =   5
      _CBWidth        =   20010
      _CBHeight       =   975
      _Version        =   "6.0.8169"
      Caption1        =   "Filter:"
      Child1          =   "txt_filter"
      MinHeight1      =   600
      Width1          =   2925
      NewRow1         =   0   'False
      Caption2        =   "Tanggal"
      Child2          =   "tgl"
      MinHeight2      =   600
      Width2          =   3495
      NewRow2         =   0   'False
      Child3          =   "chk_Sampai"
      MinHeight3      =   255
      Width3          =   405
      NewRow3         =   0   'False
      Caption4        =   "Sampai"
      Child4          =   "dtp_Sampai"
      MinHeight4      =   615
      Width4          =   3495
      NewRow4         =   0   'False
      Child5          =   "Toolbar1"
      MinHeight5      =   915
      Width5          =   9000
      NewRow5         =   0   'False
      Begin MSComCtl2.DTPicker dtp_Sampai 
         Height          =   615
         Left            =   7725
         TabIndex        =   6
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   92078081
         CurrentDate     =   39459
      End
      Begin VB.CheckBox chk_Sampai 
         Height          =   255
         Left            =   6645
         TabIndex        =   5
         Top             =   360
         Width           =   210
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   915
         Left            =   10605
         TabIndex        =   4
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   1614
         ButtonWidth     =   3043
         ButtonHeight    =   1455
         Appearance      =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   7
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.TextBox txt_filter 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   660
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   180
         Width           =   2235
      End
      Begin MSComCtl2.DTPicker tgl 
         Height          =   600
         Left            =   3825
         TabIndex        =   1
         Top             =   180
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1058
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   102891521
         CurrentDate     =   39459
      End
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageListOrder"
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No Faktur"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Jam"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bayar"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Supplier"
         Object.Width           =   9701
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Lunas"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Settled"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListOrder 
      Left            =   1320
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":BF7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":C2F9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":C675
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":CA06
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":D6E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":DA94
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":DE1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":E1C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":E518
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":E8E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":ECCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":F0A4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form_List_beli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FilterCat, x, strsql As String
Dim a, ha, hut As Double

Public Sub refreshlist()
    Dim query As String
    Dim rsbeli As ADODB.Recordset
    LV1.Sorted = False
    query = "Select a.nobukti, a.tanggal, a.jam, a.pembayaran, a.kode_supplier, b.nmsuplier, a.total, a.lunas, a.settled, a.tanggal_lunas from bill_beli a, tbsuplier b where a.kode_supplier = b.kdsuplier"
    If dtp_Sampai.Enabled Then
        'Set rsbeli = con.Execute("Select * from bill_beli where tanggal>='" & Format(tgl, "yyyy-mm-dd") & "' and tanggal <='" & Format(dtp_Sampai, "yyyy-mm-dd") & "' and nobukti like '%" & txt_filter & "%'")
        query = query & " and a.tanggal>='" & Format(tgl, "yyyy-mm-dd") & "' and a.tanggal <='" & Format(dtp_Sampai, "yyyy-mm-dd") & "'"
    Else
        'Set rsbeli = con.Execute("SELECT * from bill_beli where tanggal='" & Format(tgl, "yyyy-mm-dd") & "' and nobukti like '%" & txt_filter & "%' ")
        query = query & " and a.tanggal='" & Format(tgl, "yyyy-mm-dd") & "'"
    End If
    
    query = query & " and (a.nobukti like '%" & txt_filter & "%' or b.nmsuplier like '%" & txt_filter & "%' or a.total like '%" & txt_filter & "%')"
    
    Set rsbeli = con.Execute(query)
    
    LV1.ListItems.Clear
    CoolBar1.Bands(4).Caption = "Record : 0"
    
'    If rsbeli.RecordCount = 0 Then
'        Toolbar1.Buttons(2).Enabled = False
'        Toolbar1.Buttons(3).Enabled = False
'        Exit Sub
'    End If
'
'    Toolbar1.Buttons(2).Enabled = True
'    Toolbar1.Buttons(3).Enabled = True
    
    If rsbeli.EOF Then
        rsbeli.Close
        Exit Sub
    End If
    
    rsbeli.MoveFirst
    
    Do While Not rsbeli.EOF
        Dim bayar_string As String
        If rsbeli!pembayaran = 0 Then
            bayar_string = "Tunai"
        ElseIf rsbeli!pembayaran = 1 Then
            bayar_string = "Transfer"
        Else
            bayar_string = "Cek"
        End If
        
        Dim mitem As ListItem
        Set mitem = LV1.ListItems.Add(, , rsbeli.Fields("nobukti"))
        mitem.SubItems(1) = rsbeli!jam
        mitem.SubItems(2) = bayar_string
        If getSupplier(rsbeli!kode_supplier) Then
            mitem.SubItems(3) = rsSupplier!nmsuplier
        Else
            mitem.SubItems(3) = ""
        End If
        
        mitem.SubItems(4) = Format(rsbeli!total, "###,###,##0")
        
        If rsbeli.Fields("lunas") = 0 Then
          mitem.ForeColor = vbRed
          mitem.SubItems(5) = ""
        Else
          mitem.ForeColor = vbGreen
          mitem.SubItems(5) = "v"
        End If
      
        If rsbeli.Fields("settled") = 1 Then
          mitem.SubItems(6) = Format(rsbeli!tanggal_lunas, "dd/MM/yyyy")
          mitem.ForeColor = vbBlack
        Else
          mitem.SubItems(6) = ""
        End If
        
        rsbeli.MoveNext
      Loop
    CoolBar1.Bands(4).Caption = "Record : " & LV1.ListItems.count
    rsbeli.Close
    Set rsbeli = Nothing
End Sub

Private Sub btn_export_Click()
    Dim data_beli, bill_beli As ADODB.Recordset
    Set data_beli = con.Execute("select * from tbbeli where tglbukti = '" & Format(tgl.Value, "yyyy-mm-dd") & "'")
    Set bill_beli = con.Execute("select * from bill_beli where tanggal = '" & Format(tgl.Value, "yyyy-mm-dd") & "'")
    If data_beli.EOF Or data_beli.BOF Then
        MsgBox "Data Kosong"
        Exit Sub
    End If
    
    Dim export_data, keyword1, keyword2, keyword3, quote, comma, tanggal As String
    quote = "'"
    comma = ","
    tanggal = quote + Format(tgl.Value, "yyyy-mm-dd") + quote
    'editV2
    keyword1 = "insert into tbbeli (nobukti, tglbukti, kode, nama_barang, harga, jumlah, `return`) values("
    keyword2 = "update tbbarang set tgl_masuk=" + tanggal + " where kode="
    'editV2
    keyword3 = "insert into bill_beli (nobukti, staff, tanggal, jam, total, kode_supplier, pembayaran, lunas, settled, tanggal_lunas) values("
    
    data_beli.MoveFirst
    Do While Not data_beli.EOF
        export_data = export_data + keyword1 + quote + data_beli!nobukti + quote + comma + tanggal + comma + quote + data_beli!kode + quote + comma + quote + data_beli!nama_barang + quote + comma + CStr(data_beli!harga) + comma + CStr(data_beli!jumlah) + comma + CStr(data_beli!return) + ");"
        'export_data = export_data + keyword1 + quote + data_beli!nobukti + quote + comma + tanggal + comma + CStr(data_beli!kdsuplier) + comma + quote + data_beli!kode + quote + comma + CStr(data_beli!jumlah) + comma + CStr(data_beli!harga) + comma + CStr(data_beli!bayar) + comma + "'L'" + comma + tanggal + comma + "0" + comma + tanggal + ",0,0,0," + quote + data_beli!nama_barang + quote + ",0,0);"
        export_data = export_data + keyword2 + quote + data_beli!kode + quote + ";"
        data_beli.MoveNext
    Loop
    
    bill_beli.MoveFirst
    Do While Not bill_beli.EOF
        export_data = export_data + keyword3 + quote + bill_beli!nobukti + quote + comma + quote + bill_beli!staff + quote + comma + tanggal + comma + quote + bill_beli!jam + quote + comma + CStr(bill_beli!total) + comma + CStr(bill_beli!kode_supplier) + comma + CStr(bill_beli!pembayaran) + comma + CStr(bill_beli!lunas) + comma + CStr(bill_beli!settled) + comma + quote + Format(bill_beli!tanggal_lunas, "yyyy-mm-dd") + quote + ");"
        bill_beli.MoveNext
    Loop
    
    On Error GoTo errorHandler
    Dim httpReq As XMLHTTP60
    Set httpReq = New XMLHTTP60
    httpReq.Open "POST", "http://www.chip-padang.com/export_data.php", False
    httpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    httpReq.setRequestHeader "User-Agent", "Firefox 3.6.4"
    Dim post_data As String
    post_data = "tanggal=" + Format(tgl.Value, "yyyy-mm-dd") + "&data=" + Base64EncodeString(export_data)
    httpReq.send post_data
    
    MsgBox httpReq.responseText
    
'    CommonDialog1.Filter = "Apps (*.sql)|*.sql|All files (*.*)|*.*"
'    CommonDialog1.DefaultExt = "sql"
'    CommonDialog1.DialogTitle = "Select File"
'    CommonDialog1.ShowSave
'
'    Open CommonDialog1.FileName For Output As #1
'    'Menyimpan semua data
'    For i = 1 To Ndata
'        Print #1, export_data
'    Next i
'    Menutup File
'    Close #1
    Exit Sub
errorHandler:
    MsgBox "Export Data Gagal"
End Sub

Private Sub btn_import_Click()
    'CommonDialog1.Filter = "Apps (*.sql)|*.sql|All files (*.*)|*.*"
    'CommonDialog1.DefaultExt = "sql"
    'CommonDialog1.DialogTitle = "Select File"
    'CommonDialog1.ShowOpen
    
    On Error GoTo errorHandler
    Dim base64_data, import_data, import_url As String
    Dim httpReq As XMLHTTP60
    Set httpReq = New XMLHTTP60
    import_url = "http://www.chip-padang.com/import_data.php"
    httpReq.Open "POST", import_url, False
    httpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    httpReq.setRequestHeader "User-Agent", "Firefox 3.6.4"
    httpReq.send "tanggal=" + Format(tgl.Value, "yyyy-mm-dd")
    
    base64_data = httpReq.responseText
    import_data = Base64DecodeString(base64_data)

    'Dim FSO As FileSystemObject
    'Dim TS As TextStream
    'Set FSO = New FileSystemObject
    'Set TS = FSO.OpenTextFile(CommonDialog1.FileName, ForReading)
    'import_data = TS.ReadAll
    'TS.Close
    
    Dim sql_query() As String
    sql_query = Split(import_data, ";")
    
    Dim i As Integer
    For i = 0 To UBound(sql_query) - 1
        con.Execute (sql_query(i))
    Next
    MsgBox "Import Success"
    Exit Sub
errorHandler:
    MsgBox "Import Gagal"
    Err.Clear
End Sub

Private Sub chk_Sampai_Click()
    If dtp_Sampai.Enabled Then
        dtp_Sampai.Enabled = False
    Else
        dtp_Sampai.Enabled = True
    End If
    Call refreshlist
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
  Form_Resize
End Sub

Private Sub dtp_Sampai_Change()
    Call refreshlist
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyInsert
             tambah
        Case vbKeyEscape
            Unload Me
        Case vbKeyF5
            refreshlist
    End Select
End Sub

Private Sub Form_Load()
  tgl = Date
  dtp_Sampai = Date
  Dim i As Integer
  For i = 1 To LV1.ColumnHeaders.count
    LV1.ColumnHeaders.item(i).Icon = 0
  Next
  LV1.ColumnHeaders.item(1).Icon = 1
  txt_filter.Text = ""
'  Toolbar1.Buttons(4).Visible = isMaster
'  Toolbar1.Buttons(5).Visible = isMaster
'    Toolbar1.Buttons(4).Visible = False
'    Toolbar1.Buttons(5).Visible = False
End Sub

Private Sub Form_Resize()
  CoolBar1.Width = Me.ScaleWidth
  LV1.Top = Me.ScaleTop + CoolBar1.Height
  LV1.Left = Me.ScaleLeft
  LV1.Width = Me.ScaleWidth
  LV1.Height = IIf(Me.ScaleHeight - CoolBar1.Height > 0, Me.ScaleHeight - CoolBar1.Height, 0)
End Sub
Private Sub LV1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LV1.Sorted = True
    Dim i As Byte
    For i = 1 To LV1.ColumnHeaders.count
        LV1.ColumnHeaders.item(i).Icon = 0
    Next
    If LV1.SortKey <> ColumnHeader.index - 1 Then
        LV1.SortOrder = lvwAscending
        ColumnHeader.Icon = 1
        LV1.SortKey = ColumnHeader.index - 1
    Else
        If LV1.SortOrder = lvwAscending Then
            LV1.SortOrder = lvwDescending
            ColumnHeader.Icon = 2
        Else
            LV1.SortOrder = lvwAscending
            ColumnHeader.Icon = 1
        End If
    End If
End Sub

Private Sub tgl_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub txt_filter_change()
    refreshlist
End Sub

Private Sub LV1_DblClick()
  If LV1.ListItems.count = 0 Then
    tambah
  ElseIf LV1.SelectedItem.SubItems(6) = "" Then
    Dim no_bukti As String
    no_bukti = LV1.SelectedItem.Text
    con.Execute ("update bill_beli set lunas = case when lunas = 1 then 0 else 1 end where nobukti='" & no_bukti & "' and settled = 0")
    If LV1.SelectedItem.SubItems(5) = "v" Then
        LV1.SelectedItem.ForeColor = vbRed
        LV1.SelectedItem.SubItems(5) = ""
    Else
        LV1.SelectedItem.ForeColor = vbGreen
        LV1.SelectedItem.SubItems(5) = "v"
    End If
  End If
End Sub

Private Sub LV1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then LV1_DblClick
End Sub

Private Sub tgl_Change()
    Call refreshlist
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1
            tambah
        Case 2
            Call refreshlist
        Case 3
            Settlement
        Case 4
            Form_Print_Beli.Show
            Form_Print_Beli.Init LV1.SelectedItem.Text, 0, False
        Case 5
            deleteRecord
    End Select
End Sub

Private Sub deleteRecord()
    Dim no_bon As String
    no_bon = LV1.SelectedItem.Text
    
    If MsgBox("Hapus faktur No. " + no_bon + "?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Dim rsbeli As ADODB.Recordset
    Set rsbeli = con.Execute("select * from tbbeli where nobukti = '" & no_bon & "'")
    If rsbeli.EOF Or rsbeli.BOF Then
        Exit Sub
    End If
    
    rsbeli.MoveFirst
    Do While Not rsbeli.EOF
        con.Execute ("update tbbarang set jumlah_akhir = jumlah_akhir - " & CStr(rsbeli!jumlah - rsbeli!return) & " where kode = '" & rsbeli!kode & "'")
        rsbeli.MoveNext
    Loop
    con.Execute ("delete from tbbeli where nobukti='" & no_bon & "'")
    con.Execute ("delete from bill_beli where nobukti='" & no_bon & "'")
    LV1.ListItems.Remove (LV1.SelectedItem.index)
End Sub

Private Sub Settlement()
    Dim rsbeli As ADODB.Recordset
    Dim message, no_bon As String
    Dim nominal As Long
    Dim count As Integer
    
    Set rsbeli = con.Execute("select * from bill_beli where lunas=1 and settled=0")
    If rsbeli.EOF Or rsbeli.BOF Then
        MsgBox "Settlement Kosong"
        rsbeli.Close
        Exit Sub
    End If
    
    nominal = 0
    count = 0
    no_bon = ""
    message = "SETTLEMENT REPORT" + vbCrLf
    rsbeli.MoveFirst
    Do While Not rsbeli.EOF
        nominal = nominal + rsbeli.Fields("total")
        count = count + 1
        no_bon = no_bon + CStr(rsbeli.Fields("nobukti")) + " "
        If count Mod 5 = 0 Then
            no_bon = no_bon + vbCrLf
        End If
        rsbeli.MoveNext
    Loop
    
    message = message + "Jumlah pembayaran = " + Format(CStr(nominal), "###,###,##0") + vbCrLf + "Jumlah Bon = " + CStr(count) + " bon" + vbCrLf + no_bon
    If MsgBox(message, vbOKCancel, "Settlement Pembelian") = vbOK Then
        con.Execute ("update bill_beli set settled = 1, tanggal_lunas='" & Format(Now, "yyyy-mm-dd") & "' where lunas = 1 and settled = 0")
        MsgBox "Settlement Sukses"
        Call refreshlist
    End If
    
    rsbeli.Close
End Sub

Private Sub tambah()
  Form_Pembelian.Show
End Sub

Private Sub txt_filter_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub
