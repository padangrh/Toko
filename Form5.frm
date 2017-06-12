VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form_List_barang 
   Caption         =   "Barang"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   ControlBox      =   0   'False
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3600
      Top             =   3240
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
            Picture         =   "Form5.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":701E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":8052
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":8F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":9E5F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2143
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode Barang"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   9701
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Kategori"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Modal"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Jual"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Supplier"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tgl. Masuk"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   2355
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
            Picture         =   "Form5.frx":AADB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":AEA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":B28D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":B667
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":BBD6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   945
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17655
      _ExtentX        =   31141
      _ExtentY        =   1667
      _CBWidth        =   17655
      _CBHeight       =   945
      _Version        =   "6.0.8169"
      Caption1        =   "Filter"
      Child1          =   "txt_filter"
      MinHeight1      =   600
      Width1          =   6000
      NewRow1         =   0   'False
      Child2          =   "Toolbar1"
      MinHeight2      =   885
      Width2          =   9000
      NewRow2         =   0   'False
      MinHeight3      =   360
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   885
         Left            =   6195
         TabIndex        =   3
         Top             =   30
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   1561
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
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txt_filter 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   615
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   165
         Width           =   5355
      End
   End
   Begin MSComctlLib.ImageList ImageListOrder 
      Left            =   2400
      Top             =   2475
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
            Picture         =   "Form5.frx":BFB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":C32F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form_List_barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FilterCat, strsql As String
Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
  Form_Resize
End Sub
Private Sub Form_Resize()
  CoolBar1.Width = Me.ScaleWidth
  LV1.Top = Me.ScaleTop + CoolBar1.Height
  LV1.Left = Me.ScaleLeft
  LV1.Width = Me.ScaleWidth
  LV1.Height = IIf(Me.ScaleHeight - CoolBar1.Height > 0, Me.ScaleHeight - CoolBar1.Height, 0)
End Sub
Public Sub refreshlist()
    Dim mitem As ListItem
    Dim rsbarang As New ADODB.Recordset
    Set rsbarang = con.Execute("SELECT * from v_barang where nama like '%" & txt_filter & "%'")
    
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
  
  CoolBar1.Bands(3).Caption = "Record : " & LV1.ListItems.count
  rsbarang.Close
  Set rsbarang = Nothing
End Sub

Private Sub Form_Load()
  Dim i As Integer
  For i = 1 To LV1.ColumnHeaders.count
    LV1.ColumnHeaders.item(i).Icon = 0
  Next
  LV1.ColumnHeaders.item(1).Icon = 1
  txt_filter.Text = ""
  
 End Sub
Private Sub LV1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
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
Private Sub txt_filter_change()
    refreshlist
End Sub

Private Sub LV1_DblClick()
  If LV1.ListItems.count = 0 Then
    tambah
  Else
    perbaiki
  End If
End Sub

Private Sub LV1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then LV1_DblClick
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.index
  Case 1
    tambah
  Case 2
    perbaiki
  Case 3
    Call hapus
  Case 4
    Call refreshlist
  Case 5
    Unload Me
  End Select
End Sub
Private Sub tambah()
  Form_Entri_Barang.Show (1)
  CoolBar1.Bands(3).Caption = "Record : " & LV1.ListItems.count
End Sub
Private Sub perbaiki()
    Form_Entri_Barang.txt_kode.Text = LV1.SelectedItem.Text
    Form_Entri_Barang.Show (1)
End Sub

Private Sub hapus()
  If LV1.ListItems.count = 0 Then Exit Sub
  If MsgBox("Benar Data akan dihapus?", vbQuestion + vbYesNo, "Hapus") = vbYes Then
    strsql = "delete from tbbarang where kode='" & LV1.SelectedItem.Text & "'"
    con.BeginTrans
    con.Execute (strsql)
    con.CommitTrans
    LV1.ListItems.Remove (LV1.SelectedItem.index)
  End If
  LV1.SetFocus
  CoolBar1.Bands(3).Caption = "Record : " & LV1.ListItems.count
  If LV1.ListItems.count = 0 Then refreshlist
End Sub

Private Sub txt_filter_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub
