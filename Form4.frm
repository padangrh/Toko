VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Entri Data Kategori"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12585
   ControlBox      =   0   'False
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   12585
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3600
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
            Picture         =   "Form4.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":701E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":8052
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":8F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":9E5F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   1215
      Left            =   120
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Kategori"
         Object.Width           =   17639
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
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
            Picture         =   "Form4.frx":AADB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":AEA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":B28D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":B667
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":BBD6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   945
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   1667
      _CBWidth        =   17175
      _CBHeight       =   945
      _Version        =   "6.0.8169"
      Caption1        =   "Filter"
      Child1          =   "Text1"
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
         Width           =   10725
         _ExtentX        =   18918
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
      Begin VB.TextBox Text1 
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
      Left            =   2280
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
            Picture         =   "Form4.frx":BFB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":C32F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FilterCat, strsql, statuskategori As String
Dim rsKategori As New ADODB.Recordset
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
  Dim mitem
  strsql = "SELECT * from tbkategori"
  Dim tmpC
  
  If Text1.Text <> "" Then
    Select Case FilterCat
    Case "Kode Suplier"
      strsql = strsql & " where kode like '" & Text1.Text & "%'"
      Case Else
      MsgBox "Filter Tidak Diterima!", vbOKOnly + vbInformation, "Filter"
    End Select
  Else
    strsql = "SELECT * from tbkategori"
  End If
   Set rsKategori = con.Execute(strsql)
  LV1.ListItems.Clear
  If rsKategori.RecordCount = 0 Then
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = False
  Else
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    If Not rsKategori.EOF Then
    rsKategori.MoveFirst
    
    Do While Not rsKategori.EOF
      Set mitem = LV1.ListItems.Add(, , rsKategori.Fields("kode"))
     rsKategori.MoveNext
    Loop
    End If
  End If
  CoolBar1.Bands(3).Caption = "Record : " & LV1.ListItems.count
  rsKategori.Close
  Set rsKategori = Nothing
End Sub

Private Sub Form_Load()
  Dim i As Integer
  For i = 1 To LV1.ColumnHeaders.count
    LV1.ColumnHeaders.item(i).Icon = 0
  Next
  LV1.ColumnHeaders.item(1).Icon = 1
  Text1.Text = ""
  
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
    perbaiki
  End If
End Sub

Private Sub LV1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then LV1_DblClick
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
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
  statuskategori = "tambah"
  Form8.Show (1)
  CoolBar1.Bands(3).Caption = "Record : " & LV1.ListItems.count
End Sub
Private Sub perbaiki()
statuskategori = "perbaiki"
  Form8.Text1.Text = LV1.SelectedItem.Text
  Form8.Show (1)
End Sub

Private Sub hapus()
  If LV1.ListItems.count = 0 Then Exit Sub
  If MsgBox("Benar Data akan dihapus?", vbQuestion + vbYesNo, "Hapus") = vbYes Then
    strsql = "delete from tbkategori where kode='" & LV1.SelectedItem.Text & "'"
    con.BeginTrans
    con.Execute (strsql)
    con.CommitTrans
    LV1.ListItems.Remove (LV1.SelectedItem.Index)
  End If
  LV1.SetFocus
  CoolBar1.Bands(3).Caption = "Record : " & LV1.ListItems.count
  If LV1.ListItems.count = 0 Then refreshlist
End Sub



