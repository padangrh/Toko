VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form_Retur 
   Caption         =   "Penjualan"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15465
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Retur.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   15465
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2640
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   108
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":701E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":7F03
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":8B7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":955C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":A484
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":A84C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":AC36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":B010
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":B57F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   1667
      BandCount       =   4
      _CBWidth        =   18255
      _CBHeight       =   945
      _Version        =   "6.0.8169"
      Caption1        =   "Filter"
      Child1          =   "txt_filter"
      MinHeight1      =   600
      Width1          =   6000
      NewRow1         =   0   'False
      Caption2        =   "Tanggal"
      Child2          =   "DTPicker1"
      MinHeight2      =   600
      Width2          =   3495
      NewRow2         =   0   'False
      Child3          =   "Toolbar1"
      MinHeight3      =   885
      Width3          =   9000
      NewRow3         =   0   'False
      MinHeight4      =   360
      NewRow4         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   885
         Left            =   9720
         TabIndex        =   4
         Top             =   30
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   1561
         ButtonWidth     =   3043
         ButtonHeight    =   1455
         Appearance      =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   600
         Left            =   6900
         TabIndex        =   2
         Top             =   165
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1058
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   48234499
         CurrentDate     =   42191
      End
      Begin VB.TextBox txt_filter 
         Appearance      =   0  'Flat
         Height          =   600
         Left            =   615
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   165
         Width           =   5355
      End
   End
   Begin MSComctlLib.ImageList ImageListOrder 
      Left            =   1800
      Top             =   3120
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
            Picture         =   "Form_Retur.frx":B95D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Retur.frx":BCD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv_Retur 
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageListOrder"
      ForeColor       =   0
      BackColor       =   12648447
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode Barang"
         Object.Width           =   5583
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   17701
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tanggal"
         Object.Width           =   3916
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nama_Kasir"
         Object.Width           =   5001
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Jumlah"
         Object.Width           =   2963
      EndProperty
   End
End
Attribute VB_Name = "Form_Retur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub refreshlist()
    lv_Retur.Sorted = False
    
    Dim rsJual As ADODB.Recordset
    Dim tunai, nontunai As Long
    tunai = 0
    nontunai = 0
    Dim mitem As ListItem
    Dim query_all, query_some As String
'    query_all = "SELECT * from bill where tanggal='" & Format(DTPicker1, "yyyy-mm-dd") & "' and nobukti like '%" & txt_filter & "%'"
'    query_some = "SELECT * from bill where tanggal='" & Format(DTPicker1, "yyyy-mm-dd") & "' and kasir='" & username & "' and nobukti like '%" & txt_filter & "%'"
    query_all = "Select * from tbretur where nama like '%" & txt_filter & "%'"
'    If isSPV Or isMaster Then
'      Set rsJual = con.Execute(query_all)
'    Else
'      Set rsJual = con.Execute(query_some)
'    End If
    Set rsJual = con.Execute(query_all)
  
    lv_Retur.ListItems.Clear
    If rsJual.RecordCount = 0 Then
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
    Else
        Toolbar1.Buttons(2).Enabled = True
        Toolbar1.Buttons(3).Enabled = True
        If Not rsJual.EOF Then
            rsJual.MoveFirst
      
            Do While Not rsJual.EOF
                Set mitem = lv_Retur.ListItems.Add(, , rsJual!kode)
                mitem.SubItems(1) = rsJual!Nama
                mitem.SubItems(2) = rsJual!tgl_retur
                mitem.SubItems(3) = rsJual!userid
                mitem.SubItems(4) = rsJual!jumlah
        
                rsJual.MoveNext
            Loop
        End If
    End If
    rsJual.Close
    
    Set rsJual = Nothing
End Sub

'Private Sub Command1_Click()
'    Dim i As Integer
'    For i = 1 To lv_Retur.ColumnHeaders.count
'        MsgBox lv_Retur.ColumnHeaders(i).Width
'    Next
'End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    Form_Resize
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub Form_Load()
    DTPicker1 = Date
    Dim i As Integer
    For i = 1 To lv_Retur.ColumnHeaders.count
      lv_Retur.ColumnHeaders.item(i).Icon = 0
    Next
    lv_Retur.ColumnHeaders.item(1).Icon = 1
    txt_filter.Text = ""
    
'    Toolbar1.Buttons(4).Visible = isMaster
End Sub
  
Private Sub Form_Resize()
    CoolBar1.Width = Me.ScaleWidth
    lv_Retur.Top = Me.ScaleTop + CoolBar1.Height
    lv_Retur.Left = Me.ScaleLeft
    lv_Retur.Width = Me.ScaleWidth
    lv_Retur.Height = IIf(Me.ScaleHeight - CoolBar1.Height > 0, Me.ScaleHeight - CoolBar1.Height, 0)
    
End Sub

Private Sub lv_Retur_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 ' lv_Retur.sortedby lvwAscending
    lv_Retur.Sorted = True
    Dim i As Byte
    For i = 1 To lv_Retur.ColumnHeaders.count
        lv_Retur.ColumnHeaders.item(i).Icon = 0
    Next
    If lv_Retur.SortKey <> ColumnHeader.index - 1 Then
        lv_Retur.SortOrder = lvwAscending
        ColumnHeader.Icon = 1
        lv_Retur.SortKey = ColumnHeader.index - 1
    Else
        If lv_Retur.SortOrder = lvwAscending Then
            lv_Retur.SortOrder = lvwDescending
            ColumnHeader.Icon = 2
        Else
            lv_Retur.SortOrder = lvwAscending
            ColumnHeader.Icon = 1
        End If
    End If
End Sub

Private Sub txt_filter_change()
    refreshlist
End Sub

Private Sub lv_Retur_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then lv_Retur_DblClick
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
        Call deletePenjualan
      End Select
End Sub

Private Sub deletePenjualan()
    If Me.ActiveControl.Name = "lv_Retur" Then
        If (Not lv_Retur.SelectedItem Is Nothing) Then
            If hapusRetur(lv_Retur.SelectedItem.Text) Then
                lv_Retur.ListItems.Remove (lv_Retur.SelectedItem.index)
            End If
        End If
    Else
        MsgBox "Tidak ada data yang dipilih"
    End If
End Sub

Private Function hapusRetur(inKode As String) As Boolean
    If MsgBox("Hapus retur barang " + inKode + "?", vbYesNo, "Konfirmasi") = vbYes Then
        Dim rsJual As ADODB.Recordset
        con.Execute ("delete from tbretur where kode ='" & inKode & "'")
'        Set rsJual = con.Execute("select * from tbjual where nobukti='" & inKode & "'")
'        If Not rsJual.EOF Then
'            rsJual.MoveFirst
'            Do While Not rsJual.EOF
'                con.Execute ("update tbbarang set jumlah_akhir = jumlah_akhir + " & rsJual!jumlah_jual & " where kode='" & rsJual!kode & "'")
'                rsJual.MoveNext
'            Loop
'            con.Execute ("delete from tbjual where nobukti='" & no_bon & "'")
'        End If
        hapusRetur = True
    Else
        hapusRetur = False
    End If
End Function

Private Sub tambah()
    Form_Tambah_Retur.Show
'    CoolBar1.Bands(3).Caption = "Record : " & lv_Retur.ListItems.count
End Sub

Private Sub dtpicker1_Change()
    Call refreshlist
End Sub

Private Sub lv_Retur_DblClick()
    If Not (lv_Retur.SelectedItem Is Nothing) Then
'        Form_Print.Show
'        Form_Print.Init lv_Retur.SelectedItem.Text, lv_Retur.SelectedItem.SubItems(3), False
        Dim x As String
        Dim y As Integer
        x = InputBox("Masukkan jumlah baru", "Update jumlah retur", lv_Retur.SelectedItem.SubItems(4))
        
'        If Len(x) > 4 Or Val(x) = 0 Then
'            y = 0
'        Else
'            y = Val(x)
'        End If
        
        y = IIf(Len(x) > 4 Or Val(x) = 0, 0, Val(x))
        
        If x = "" Then
        ElseIf y = 0 Then
            MsgBox "Jumlah yg dimasukkan tidak valid", vbOKOnly, "Gagal Disimpan"
        Else
            'con.Execute ("update tbretur set tgl_retur = '" & Format(Date, "yyyy-mm-dd") & "', userid = '" & username & "', jumlah = '" & y & "' where kode = '" & lv_Retur.SelectedItem.Text & "'")
            con.Execute ("update tbretur set jumlah = '" & y & "' where kode = '" & lv_Retur.SelectedItem.Text & "'")
            
            DoEvents
            refreshlist
            MsgBox "Data berhasil diubah"
        End If
    End If
End Sub

Private Sub txt_filter_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub
