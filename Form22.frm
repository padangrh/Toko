VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_Warning 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form22"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14265
   LinkTopic       =   "Form22"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   14265
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LV1 
      Height          =   7695
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   13573
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
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tanggal Masuk"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ketahanan"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sisa Waktu"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Jumlah"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListOrder 
      Left            =   0
      Top             =   0
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
            Picture         =   "Form22.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form22.frx":037B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Barang Yang Hampir Expired"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   13455
   End
End
Attribute VB_Name = "Form_Warning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 If con.State = adStateClosed Then connect
  
 strsql = "SELECT * from v_warning where sisa_waktu < ketahanan * 0.3 and ketahanan > 0"
 Set rswarning = con.Execute(strsql)
 LV1.ListItems.Clear
  If rswarning.RecordCount = 0 Then
        Label1.Caption = "TIDAK ADA BARANG HAMPIR EXPIRED"
  Else
    If Not rswarning.EOF Then
    
    rswarning.MoveFirst
    Do While Not rswarning.EOF
      Set mitem = LV1.ListItems.Add(, , rswarning.Fields("kode"))
      mitem.SubItems(1) = rswarning.Fields("nama")
      mitem.SubItems(2) = rswarning.Fields("tgl_masuk")
      mitem.SubItems(3) = CStr(rswarning.Fields("ketahanan")) + " hari"
      mitem.SubItems(4) = CStr(rswarning.Fields("sisa_waktu")) + " hari"
      mitem.SubItems(5) = rswarning.Fields("jumlah")
      rswarning.MoveNext
    Loop
    End If
  End If
  LV1.ColumnHeaders(5).Icon = 1
End Sub

Private Sub LV1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim strQuery As String
    strQuery = "SELECT * from v_warning where sisa_waktu < ketahanan * 0.3 and ketahanan > 0"
    
    Select Case ColumnHeader.index
        Case 1
            strQuery = strQuery & " order by kode"
        Case 2
            strQuery = strQuery & " order by nama"
        Case 3
            strQuery = strQuery & " order by tgl_masuk"
        Case 4
            strQuery = strQuery & " order by ketahanan"
        Case 5
            strQuery = strQuery & " order by sisa_waktu"
        Case 6
            strQuery = strQuery & " order by jumlah"
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
    
    Set rswarning = con.Execute(strQuery)
    LV1.ListItems.Clear
    If rswarning.RecordCount = 0 Then
        Label1.Caption = "TIDAK ADA BARANG HAMPIR EXPIRED"
    Else
        If Not rswarning.EOF Then
            rswarning.MoveFirst
            Do While Not rswarning.EOF
                Set mitem = LV1.ListItems.Add(, , rswarning.Fields("kode"))
                mitem.SubItems(1) = rswarning.Fields("nama")
                mitem.SubItems(2) = rswarning.Fields("tgl_masuk")
                mitem.SubItems(3) = CStr(rswarning.Fields("ketahanan")) + " hari"
                mitem.SubItems(4) = CStr(rswarning.Fields("sisa_waktu")) + " hari"
                mitem.SubItems(5) = rswarning.Fields("jumlah")
                rswarning.MoveNext
            Loop
        End If
    End If
End Sub
