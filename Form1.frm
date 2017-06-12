VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form_List_Supplier 
   Caption         =   "Suplier"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15600
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   15600
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
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
            Picture         =   "Form1.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":701E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8052
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9E5F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv_supplier 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2143
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
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode Suplier"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Suplier"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Alamat Suplier"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Telp/HP"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tgl. Bergabung"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   3240
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
            Picture         =   "Form1.frx":AADB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AEA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B28D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B667
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BBD6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   19575
      _ExtentX        =   34528
      _ExtentY        =   1667
      _CBWidth        =   19575
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
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   1561
         ButtonWidth     =   3043
         ButtonHeight    =   1455
         Appearance      =   1
         ImageList       =   "ImageList2"
         HotImageList    =   "ImageList2"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   615
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   165
         Width           =   5355
      End
   End
   Begin MSComctlLib.ImageList ImageListOrder 
      Left            =   1680
      Top             =   3360
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
            Picture         =   "Form1.frx":BFB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C32F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form_List_Supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
  Form_Resize
End Sub
Private Sub Form_Resize()
  CoolBar1.Width = Me.ScaleWidth
  lv_supplier.Top = Me.ScaleTop + CoolBar1.Height
  lv_supplier.Left = Me.ScaleLeft
  lv_supplier.Width = Me.ScaleWidth
  lv_supplier.Height = IIf(Me.ScaleHeight - CoolBar1.Height > 0, Me.ScaleHeight - CoolBar1.Height, 0)
End Sub
Public Sub refreshlist()
    Set rsSupplier = con.Execute("select * from tbsuplier where nmsuplier like '%" & txt_filter & "%'")
    
    lv_supplier.ListItems.Clear
    If rsSupplier.RecordCount = 0 Then
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
    Else
        Toolbar1.Buttons(2).Enabled = True
        Toolbar1.Buttons(3).Enabled = True
        If Not rsSupplier.EOF Then
            rsSupplier.MoveFirst
            Do While Not rsSupplier.EOF
                Dim mitem As ListItem
                Set mitem = lv_supplier.ListItems.Add(, , rsSupplier.Fields("Kdsuplier"))
                mitem.SubItems(1) = rsSupplier.Fields("Nmsuplier")
                mitem.SubItems(2) = rsSupplier.Fields("alamat")
                mitem.SubItems(3) = rsSupplier.Fields("telp")
                mitem.SubItems(4) = Format(rsSupplier.Fields("tgl_gabung"), "dd-MM-yyyy")
                rsSupplier.MoveNext
            Loop
        End If
    End If
    rsSupplier.Close
    CoolBar1.Bands(3).Caption = "Record : " & lv_supplier.ListItems.count
End Sub

Private Sub Form_Load()
  Dim i As Integer
  For i = 1 To lv_supplier.ColumnHeaders.count
    lv_supplier.ColumnHeaders.item(i).Icon = 0
  Next
  lv_supplier.ColumnHeaders.item(1).Icon = 1
  txt_filter.Text = ""
  refreshlist
 End Sub
Private Sub lv_supplier_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Dim i As Byte
  For i = 1 To lv_supplier.ColumnHeaders.count
    lv_supplier.ColumnHeaders.item(i).Icon = 0
  Next
  If lv_supplier.SortKey <> ColumnHeader.index - 1 Then
    lv_supplier.SortOrder = lvwAscending
    ColumnHeader.Icon = 1
    lv_supplier.SortKey = ColumnHeader.index - 1
  Else
    If lv_supplier.SortOrder = lvwAscending Then
      lv_supplier.SortOrder = lvwDescending
      ColumnHeader.Icon = 2
    Else
      lv_supplier.SortOrder = lvwAscending
      ColumnHeader.Icon = 1
    End If
  End If
End Sub

Private Sub lv_supplier_DblClick()
  If lv_supplier.ListItems.count = 0 Then
    tambah
  Else
    perbaiki
  End If
End Sub

Private Sub lv_supplier_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then lv_supplier_DblClick
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
  Form_Entri_Supplier.Show (1)
  CoolBar1.Bands(3).Caption = "Record : " & lv_supplier.ListItems.count
End Sub
Private Sub perbaiki()
  Form_Entri_Supplier.txt_kode.Text = lv_supplier.SelectedItem.Text
  Form_Entri_Supplier.Show (1)
End Sub

Private Sub hapus()
  If lv_supplier.ListItems.count = 0 Then Exit Sub
  If MsgBox("Benar Data akan dihapus?", vbQuestion + vbYesNo, "Hapus") = vbYes Then
    con.BeginTrans
    con.Execute ("delete from tbsuplier where kdsuplier='" & lv_supplier.SelectedItem.Text & "'")
    con.CommitTrans
    lv_supplier.ListItems.Remove (lv_supplier.SelectedItem.index)
  End If
  lv_supplier.SetFocus
  CoolBar1.Bands(3).Caption = "Record : " & lv_supplier.ListItems.count
  If lv_supplier.ListItems.count = 0 Then refreshlist
End Sub


Private Sub txt_filter_change()
    refreshlist
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
