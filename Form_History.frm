VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_History 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lv_history 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageListOrder"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tanggal"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Modal"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Jual"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListOrder 
      Left            =   7800
      Top             =   1680
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
            Picture         =   "Form_History.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_History.frx":037B
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kode_barang As String

Public Sub Init(kode As String)
    kode_barang = kode
End Sub


Private Sub Form_Load()
    Dim rsHistory As ADODB.Recordset
    Set rsHistory = con.Execute("select * from tbbarang_history where kode='" & kode_barang & "' order by tanggal desc")
    If rsHistory.EOF Or rsHistory.BOF Then
        Exit Sub
    End If
    
    Dim l_item As ListItem
    rsHistory.MoveFirst
    Do While Not rsHistory.EOF
        Set l_item = lv_history.ListItems.Add(, , Format(rsHistory!tanggal, "yyyy-mm-dd"))
        l_item.SubItems(1) = rsHistory!Nama
        l_item.SubItems(2) = Format(rsHistory!harga_modal, "###,###,##0")
        l_item.SubItems(3) = Format(rsHistory!harga_jual, "###,###,##0")
        rsHistory.MoveNext
    Loop
    rsHistory.Close
    For i = 1 To lv_history.ColumnHeaders.count
      lv_history.ColumnHeaders.item(i).Icon = 0
    Next
    lv_history.SortOrder = lvwDescending
    lv_history.ColumnHeaders.item(1).Icon = 2
End Sub

Private Sub lv_history_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim i As Integer
    Dim rsbarang As New ADODB.Recordset
    Dim mitem As ListItem
    Dim strQuery As String
    strQuery = "SELECT * from v_barang where nama like '%" & txt_filter & "%'"
    
    Select Case ColumnHeader.index
        Case 1
            strQuery = strQuery & " order by kode"
        Case 2
            strQuery = strQuery & " order by nama"
        Case 3
            strQuery = strQuery & " order by kategori"
        Case 4
            strQuery = strQuery & " order by harga_modal"
        Case 5
            strQuery = strQuery & " order by harga_jual"
        Case 6
            strQuery = strQuery & " order by jumlah_akhir"
        Case 7
            strQuery = strQuery & " order by nmsuplier"
        Case 8
            strQuery = strQuery & " order by tgl_masuk"
    End Select
    
    If ColumnHeader.Icon = 1 Then
        strQuery = strQuery & " desc"
        For i = 1 To LV1.ColumnHeaders.count
            LV1.ColumnHeaders.item(i).Icon = 0
        Next
        ColumnHeader.Icon = 2
    Else
        strQuery = strQuery & " asc"
        For i = 1 To LV1.ColumnHeaders.count
            LV1.ColumnHeaders.item(i).Icon = 0
        Next
        ColumnHeader.Icon = 1
    End If
    
    Set rsbarang = con.Execute(strQuery)
    LV1.ListItems.Clear
    If rsbarang.RecordCount = 0 Then
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
    Else
        Toolbar1.Buttons(2).Enabled = True
        Toolbar1.Buttons(3).Enabled = True
        If Not rsbarang.EOF Then
            rsbarang.MoveFirst
        
            Do While Not rsbarang.EOF
                Set mitem = LV1.ListItems.Add(, , rsbarang.Fields("kode"))
                mitem.SubItems(1) = rsbarang.Fields("Nama")
                mitem.SubItems(2) = rsbarang.Fields("kategori")
                mitem.SubItems(3) = Format(rsbarang.Fields("harga_modal"), "###,###,##0")
                mitem.SubItems(4) = Format(rsbarang.Fields("harga_jual"), "###,###,##0")
                mitem.SubItems(5) = rsbarang.Fields("jumlah_akhir")
                mitem.SubItems(6) = rsbarang!nmsuplier
                mitem.SubItems(7) = Format(rsbarang.Fields("tgl_masuk"), "dd-MM-yyyy")
                rsbarang.MoveNext
            Loop
        End If
    End If
End Sub
