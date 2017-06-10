VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form19 
   Caption         =   "Browse Return Pembelian"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14475
   ControlBox      =   0   'False
   Icon            =   "Form19.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   14475
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2640
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   108
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form19.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form19.frx":701E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form19.frx":7F03
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   945
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   17970
      _ExtentX        =   31697
      _ExtentY        =   1667
      BandCount       =   4
      _CBWidth        =   17970
      _CBHeight       =   945
      _Version        =   "6.0.8169"
      Caption1        =   "Filter:"
      Child1          =   "Text1"
      MinHeight1      =   600
      Width1          =   6000
      NewRow1         =   0   'False
      Caption2        =   "Tanggal"
      Child2          =   "tgl"
      MinHeight2      =   600
      Width2          =   3495
      NewRow2         =   0   'False
      Child3          =   "Toolbar1"
      MinHeight3      =   885
      Width3          =   9000
      NewRow3         =   0   'False
      MinHeight4      =   360
      Width4          =   840
      NewRow4         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   885
         Left            =   9720
         TabIndex        =   4
         Top             =   30
         Width           =   7995
         _ExtentX        =   14102
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
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
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
         Top             =   165
         Width           =   5310
      End
      Begin MSComCtl2.DTPicker tgl 
         Height          =   600
         Left            =   6900
         TabIndex        =   1
         Top             =   165
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
         Format          =   128581633
         CurrentDate     =   39459
      End
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   2040
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
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No. Return Beli"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No. Bukti Beli"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tgl. Return"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Kode Suplier"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nama Suplier"
         Object.Width           =   7408
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Kode Barang"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Nama Barang"
         Object.Width           =   7585
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Jml. Return"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListOrder 
      Left            =   1320
      Top             =   3600
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
            Picture         =   "Form19.frx":8B7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form19.frx":8EFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   3480
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
            Picture         =   "Form19.frx":9276
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form19.frx":963E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form19.frx":9A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form19.frx":9E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form19.frx":A371
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FilterCat, strsql As String
Dim StartStat As Boolean

Public Sub refreshlist()
 
  Dim mitem
  strsql = "select * from v_rbelitampil where tglbukti = '" & Format(tgl, "yyyy-MM-dd") & "'"
  Dim tmpC
  
  If Text1.Text <> "" Then
    Select Case FilterCat
    Case "No_Return_Beli"
      strsql = strsql & " and v_rbelitampil.noreturnbeli like '" & Text1.Text & "%'"
    Case "Nama_Suplier"
      strsql = strsql & " and v_rbelitampil.nama like '" & Text1.Text & "%'"
    Case "Kd_Suplier"
      strsql = strsql & " and v_rbelitampil.kdsuplier like '" & Text1.Text & "%'"
    Case Else
      MsgBox "Filter Tidak Diterima!", vbOKOnly + vbInformation, "Filter"
    End Select
  Else
  strsql = "select * from v_rbelitampil where tglbukti = '" & Format(tgl, "yyyy-MM-dd") & "'"
   End If
 
  Set rsreturnbeli = con.Execute(strsql)
  LV1.ListItems.Clear
  If rsreturnbeli.RecordCount = 0 Then
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = False
  Else
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    If Not rsreturnbeli.EOF Then
    rsreturnbeli.MoveFirst
    
    Do While Not rsreturnbeli.EOF
      Set mitem = LV1.ListItems.Add(, , rsreturnbeli.Fields("noreturnbeli"))
      mitem.SubItems(1) = rsreturnbeli.Fields("nobukti")
      mitem.SubItems(2) = rsreturnbeli.Fields("tglbukti")
      mitem.SubItems(3) = rsreturnbeli.Fields("kdsuplier")
      mitem.SubItems(4) = rsreturnbeli.Fields("nmsuplier")
      mitem.SubItems(5) = rsreturnbeli.Fields("kode")
      mitem.SubItems(6) = rsreturnbeli.Fields("nama")
      mitem.SubItems(7) = rsreturnbeli.Fields("jumlah")
      rsreturnbeli.MoveNext
    Loop
    End If
  End If
  CoolBar1.Bands(3).Caption = "Record : " & LV1.ListItems.count
  rsreturnbeli.Close
  Set rsreturnbeli = Nothing
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
  Form_Resize
End Sub

Private Sub Form_Activate()
  If StartStat = True Then refreshlist
  StartStat = False
  FilterCat = "No_Return_Beli"
  Text1.ToolTipText = "Filter " & FilterCat & " (eg. 'xxx', 'xxx%', '%xxx%')"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyInsert
    tambah
  Case vbKeyEscape
    Unload Me
  Case vbKeyF5
    refreshlist
  Case vbKeyDelete
    If Not Me.ActiveControl Is Text1 Then
      Hapus
    End If
  End Select
End Sub

Private Sub Form_Load()
    tgl = Date
  If con.State = adStateClosed Then connect
  StartStat = True
  Dim i As Integer
  For i = 1 To LV1.ColumnHeaders.count
    LV1.ColumnHeaders.item(i).Icon = 0
  Next
  LV1.ColumnHeaders.item(1).Icon = 1
  Text1.Text = ""

End Sub

Private Sub Form_Resize()
  CoolBar1.Width = Me.ScaleWidth
  LV1.Top = Me.ScaleTop + CoolBar1.Height
  LV1.Left = Me.ScaleLeft
  LV1.Width = Me.ScaleWidth
  LV1.Height = IIf(Me.ScaleHeight - CoolBar1.Height > 0, Me.ScaleHeight - CoolBar1.Height, 0)
End Sub



Private Sub LV1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Dim i As Byte
  For i = 1 To LV1.ColumnHeaders.count
    LV1.ColumnHeaders.item(i).Icon = 0
  Next
  If LV1.SortKey <> ColumnHeader.Index - 1 Then
    LV1.SortOrder = lvwAscending
    ColumnHeader.Icon = 1
    LV1.SortKey = ColumnHeader.Index - 1
  Else
    If LV1.SortOrder = lvwAscending Then
      LV1.SortOrder = lvwDescending
      ColumnHeader.Icon = 2
    Else
      LV1.SortOrder = lvwAscending
      ColumnHeader.Icon = 1
    End If
  End If
  FilterCat = ColumnHeader.Text
  Text1.ToolTipText = "Filter " & FilterCat & " (eg. 'xxx', 'xxx%', '%xxx%')"
End Sub

Private Sub text1_change()
If Trim(Text1) <> "" Then
refreshlist
Else
refreshlist
End If

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then refreshlist
End Sub

Private Sub LV1_DblClick()
  If LV1.ListItems.count = 0 Then
    tambah
  Else
    'perbaiki
  End If
End Sub

Private Sub LV1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then LV1_DblClick
End Sub

Private Sub tgl_Change()
Call refreshlist
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
  Case 1
    tambah
  'Case 2
    'perbaiki
  'Case 2
    Call Hapus
  Case 2
    Call refreshlist
 ' Case 6
    'LapBarang.Show
    'LapBarang.SetFocus
    Case 3
    Unload Me
  End Select
End Sub
Private Sub tambah()
  Form20.Show
  CoolBar1.Bands(3).Caption = "Record : " & LV1.ListItems.count
End Sub
Private Sub Hapus()
  If LV1.ListItems.count = 0 Then Exit Sub
  If MsgBox("Benar Data akan dihapus?", vbQuestion + vbYesNo, "Hapus") = vbYes Then
    strsql = "delete from tbreturnbeli where noreturnbeli='" & LV1.SelectedItem.Text & "'"
    con.BeginTrans
    con.Execute (strsql)
    con.CommitTrans
    LV1.ListItems.Remove (LV1.SelectedItem.Index)
  End If
  LV1.SetFocus
  CoolBar1.Bands(3).Caption = "Record : " & LV1.ListItems.count
  If LV1.ListItems.count = 0 Then refreshlist
End Sub


