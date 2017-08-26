VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_InputRetur 
   BackColor       =   &H00808080&
   Caption         =   "Retur"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   15210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_Batal 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   2
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmd_Masukkan 
      Caption         =   "Masukkan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   5760
      Width           =   2175
   End
   Begin MSComctlLib.ListView lv_Retur 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12632319
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
         Object.Width           =   4101
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   10848
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tanggal"
         Object.Width           =   3651
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nama_Kasir"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Jumlah"
         Object.Width           =   2408
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   10215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier :"
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
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Retur Barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form_InputRetur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Batal_Click()
    Unload Me
End Sub

Private Sub cmd_Masukkan_Click()
    If lv_Retur.ListItems.count > 0 Then
        Dim i As Integer
        Dim lvBeliID As Integer
        For i = 1 To lv_Retur.ListItems.count
            lvBeliID = 0
            lvBeliID = cekList(lv_Retur.ListItems(i).Text)
            If lvBeliID > 0 Then
                'ada
                Form_Pembelian.lv_beli.ListItems(lvBeliID).SubItems(4) = Val(Form_Pembelian.lv_beli.ListItems(lvBeliID).SubItems(4)) + Val(lv_Retur.ListItems(i).SubItems(4))
                Form_Pembelian.lv_beli.ListItems(lvBeliID).SubItems(5) = priceToNum(Form_Pembelian.lv_beli.ListItems(lvBeliID).SubItems(2)) * (Val(Form_Pembelian.lv_beli.ListItems(lvBeliID).SubItems(3)) - Val(Form_Pembelian.lv_beli.ListItems(lvBeliID).SubItems(4)))
                Form_Pembelian.lv_beli.ListItems(lvBeliID).SubItems(5) = Format(Form_Pembelian.lv_beli.ListItems(lvBeliID).SubItems(5), "###,###,##0")
            Else
                'tak ada
                Dim litem As ListItem
                Dim tempHarga As Long
                tempHarga = getHarga(lv_Retur.ListItems(i).Text)
                Set litem = Form_Pembelian.lv_beli.ListItems.Add(, , lv_Retur.ListItems(i).Text)
                litem.SubItems(1) = lv_Retur.ListItems(i).SubItems(1)
                litem.SubItems(2) = Format(tempHarga, "###,###,##0")
                litem.SubItems(3) = 0
                litem.SubItems(4) = Val(lv_Retur.ListItems(i).SubItems(4))
                litem.SubItems(5) = Format((0 - Val(lv_Retur.ListItems(i).SubItems(4))) * tempHarga, "###,###,##0")
            End If
        Next
        Call Form_Pembelian.hitungTxtTotal
        Unload Me
    Else
        cmd_Batal_Click
    End If
End Sub

'Private Sub Command1_Click()
'    Dim i As Integer
'    For i = 1 To lv_Retur.ColumnHeaders.count
'        MsgBox lv_Retur.ColumnHeaders(i).Width
'    Next
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        cmd_Masukkan_Click
    End If
    If KeyCode = 46 Then
        If Shift = 1 Then
            lv_Retur.ListItems.Clear
        Else
            lv_Retur.ListItems.Remove (lv_Retur.SelectedItem.index)
        End If
    End If
End Sub

Private Sub Form_unload(cancel As Integer)
    Form_Pembelian.Enabled = True
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
            lv_Retur.SelectedItem.SubItems(4) = y
                        
            DoEvents
'            refreshlist
'            MsgBox "Data berhasil diubah"
        End If
    End If
End Sub

Private Sub lv_Retur_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then lv_Retur_DblClick
End Sub

Private Function cekList(inKode As String) As Integer
    cekList = 0
    If Form_Pembelian.lv_beli.ListItems.count > 0 Then
        Dim i As Integer
        For i = 1 To Form_Pembelian.lv_beli.ListItems.count
            If inKode = Form_Pembelian.lv_beli.ListItems(i).Text Then
                cekList = i
            End If
        Next
    End If
End Function

Private Function getHarga(inKode As String) As Long
    Dim rsHarga As ADODB.Recordset
    Set rsHarga = con.Execute("select harga_modal from tbbarang where kode = '" & inKode & "'")
    If Not rsHarga.EOF Then
        getHarga = rsHarga!harga_modal
    Else
        getHarga = 0
    End If
    Set rsHarga = Nothing
End Function
