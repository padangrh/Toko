VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form7 
   Caption         =   "Pembatan Penjualan"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11160
   ControlBox      =   0   'False
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   11160
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2400
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   128
      ImageHeight     =   51
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":72BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":81A3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17985
      _ExtentX        =   31724
      _ExtentY        =   1720
      BandCount       =   4
      _CBWidth        =   17985
      _CBHeight       =   975
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
      MinHeight3      =   915
      Width3          =   9000
      NewRow3         =   0   'False
      MinHeight4      =   825
      Width4          =   9000
      NewRow4         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   915
         Left            =   9720
         TabIndex        =   4
         Top             =   30
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   1614
         ButtonWidth     =   3572
         ButtonHeight    =   1508
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
      Begin MSComCtl2.DTPicker tgl 
         Height          =   600
         Left            =   6900
         TabIndex        =   2
         Top             =   180
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1058
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   132972545
         CurrentDate     =   39459
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
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   5310
      End
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No_Bukti"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tgl_Bukti"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama_Kasir"
         Object.Width           =   15875
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jumlah Faktur"
         Object.Width           =   5909
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListOrder 
      Left            =   840
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
            Picture         =   "Form7.frx":8E1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":919A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":9516
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":98A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":A581
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":A935
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":ACBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":B065
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox cr 
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strsql As String
Dim FilterCat As String
Dim rsJual As New ADODB.Recordset

Public Sub refreshlist()
  Dim mitem
  strsql = "SELECT * from v_jualtampil where tglbukti='" & Format(tgl, "yyyy-mm-dd") & "' "
  Dim tmpC
  
  If Text1.Text <> "" Then
    Select Case FilterCat
    Case "No_Bukti"
      strsql = strsql & " and v_jualtampil.nobukti like '" & Text1.Text & "%'"
    Case "Nama_Pel"
      strsql = strsql & " and v_jualtampil.nama_pel like '" & Text1.Text & "%'"
    Case "Nama_Kasir"
      strsql = strsql & " and v_jualtampil.nm_kasir like '" & Text1.Text & "%'"
    Case Else
      MsgBox "Filter Tidak Diterima!", vbOKOnly + vbInformation, "Filter"
    End Select
  Else
  
  strsql = "SELECT * from v_jualtampil where tglbukti='" & Format(tgl, "yyyy-mm-dd") & "' "
  End If
 
  Set rsJual = con.Execute(strsql)
  LV1.ListItems.Clear
  If rsJual.RecordCount = 0 Then
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = False
  Else
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    If Not rsJual.EOF Then
    rsJual.MoveFirst
    
    Do While Not rsJual.EOF
      Set mitem = LV1.ListItems.Add(, , rsJual.Fields("NoBukti"))
      mitem.SubItems(1) = rsJual.Fields("tglbukti")
      mitem.SubItems(2) = rsJual.Fields("nm_kasir")
      mitem.SubItems(3) = rsJual.Fields("Jumlah")
      rsJual.MoveNext
    Loop
    End If
  End If
  CoolBar1.Bands(3).Caption = "Record : " & LV1.ListItems.count
  rsJual.Close
  Set rsJual = Nothing
End Sub
Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
  Form_Resize
End Sub


Private Sub Form_Load()
  tgl = Date
  If con.State = adStateClosed Then connect
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
    'tambah
  Else
    perbaiki
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
    perbaiki
  Case 2
      Call refreshlist
  Case 3
    Unload Me
  End Select
End Sub


Private Sub perbaiki()
'SAJU = "udapte"
  If LV1.ListItems.count <> 0 Then
  Form_List_Supplier0.Text1.Text = LV1.SelectedItem.Text
  Form_List_Supplier0.Show
  End If
End Sub



