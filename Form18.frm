VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form18 
   BackColor       =   &H0000C000&
   Caption         =   "Entri Return Jual"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11190
   ControlBox      =   0   'False
   Icon            =   "Form18.frx":0000
   LinkTopic       =   "Form18"
   ScaleHeight     =   8520
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Top             =   8100
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "&Tutup"
      Height          =   405
      Left            =   6360
      TabIndex        =   21
      Top             =   8100
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   7065
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10905
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   0
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   7080
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Return"
         Height          =   465
         Left            =   9720
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3600
         Width           =   1035
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5400
         TabIndex        =   9
         Top             =   210
         Width           =   2505
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2040
         TabIndex        =   8
         Top             =   210
         Width           =   1785
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Text            =   "Text6"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Text            =   "Text7"
         Top             =   2160
         Width           =   1575
      End
      Begin MSMask.MaskEdBox text9 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text8 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   2640
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##.00"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   127860739
         CurrentDate     =   41300
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form18.frx":628A
         Height          =   2760
         Left            =   120
         TabIndex        =   10
         Top             =   4200
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4868
         _Version        =   393216
         BackColor       =   12648447
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   20
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "KODE"
            Caption         =   "KODE BARANG"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "NAMA"
            Caption         =   "NAMA BARANG"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "JUMLAH"
            Caption         =   "JUMLAH"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "HARGA"
            Caption         =   "HARGA JUAL"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "DASAR"
            Caption         =   "HARGA DASAR"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "STOK"
            Caption         =   "STOK AKHIR"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "nomor"
            Caption         =   "NOMOR"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Return"
         Height          =   240
         Left            =   4275
         TabIndex        =   20
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STOK Akhir"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   19
         Top             =   2280
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Jual/Return"
         Height          =   195
         Left            =   600
         TabIndex        =   18
         Top             =   2280
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         Height          =   195
         Left            =   600
         TabIndex        =   16
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   240
         Left            =   180
         TabIndex        =   15
         Top             =   705
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti Penjualan"
         Height          =   240
         Left            =   165
         TabIndex        =   14
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jual"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   13
         Top             =   2760
         Width           =   765
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Modal"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   12
         Top             =   3240
         Width           =   915
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3600
      Top             =   8040
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   6270
      Top             =   7995
      Width           =   1305
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   1515
      Top             =   7980
      Width           =   1335
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsbarang, rsrjual As New ADODB.Recordset
Private Sub Form_Activate()
Text1.SetFocus
End Sub
Private Sub Command2_Click()
kosong
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Command4_Click()
Dim saldo, saldoAkhir, hrgbeli, hrgbelibaru
con.Execute ("insert into tbreturnjual values(" & Trim(Text15) & ",'" & Trim(Text1.Text) & "','" & Format(DTPicker1, "yyyy-MM-dd") & "','" & Text4.Text & "','" & Val(Text6.Text) & "'," & Val(text8) & "," & Val(text9) & "," & Val(Text2) & ")")
con.Execute ("insert into stock values('" & Format(DTPicker1, "yyyy-MM-dd") & "','" & Trim(Text1.Text) & "','" & Trim(Text4.Text) & "','" & Val(Text6.Text) & "',0,'Return Penjualan','" & username & "'," & Val(Text2) & ")")
con.Execute ("update tbjual set tanggal_r='" & Format(DTPicker1, "yyyy-MM-dd") & "',jumlah='" & Val(Text6.Text) & "',harga=" & Val(text8) & ",dasar=" & Val(text9) & " where nobukti='" & Trim(Text1) & "' and kode='" & Text4.Text & "' and nomor='" & Text2 & "'")
kosong
sql = "select max(tbreturnjual.noreturnjual) AS nob  From tbreturnjual"
Set Rec = con.Execute(sql)
    If IsNull(Rec!nob) = True Then
       Text15.Text = 1
    Else
       Text15.Text = Rec!nob + 1
    End If
Adodc1.Refresh
Form17.refreshlist
End Sub
Private Sub Form_Load()
kosong1
DTPicker1 = Date
If con.State = adStateClosed Then
connect
End If
sql = "select max(tbreturnjual.noreturnjual) AS nob  From tbreturnjual"
Set Rec = con.Execute(sql)
    If IsNull(Rec!nob) = True Then
       Text15.Text = 1
    Else
       Text15.Text = Rec!nob + 1
    End If
Adodc1.ConnectionString = "dsn=data"
sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah` AS `jumlah`,`j`.`harga` AS `harga`,`j`.`dasar` AS `dasar`,`o`.`jumlah_akhir` AS `stok`,`j`.`nomor` AS `nomor` From (`tbbarang` `o` join `tbjual` `j`) Where (`j`.`kode` = `o`.`Kode`) and (`j`.`nobukti` = '" & Trim(Text1.Text) & "')"
Adodc1.RecordSource = sql
End Sub
Private Sub text1_change()
sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah_jual` AS `jumlah`,`j`.`harga_jual` AS `harga`,`j`.`harga_dasar` AS `dasar`,`o`.`jumlah_akhir` AS `stok`,`j`.`nomor` AS `nomor` From (`tbbarang` `o` join `tbjual` `j`) Where (`j`.`kode` = `o`.`Kode`) and (`j`.`nobukti` = '" & Trim(Text1.Text) & "')"
Set rsbarang = con.Execute(sql)
If Not rsbarang.EOF Then
    rsbarang.MoveFirst
    Do While Not rsbarang.EOF
    rsbarang.MoveNext
    Loop
End If
Adodc1.RecordSource = sql
Adodc1.Refresh
End Sub
Sub kosong()
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
text8.Text = ""
text9.Text = ""
End Sub
Private Sub dataGrid1_DblClick()
Text2.Text = Val(DataGrid1.Columns(6))
Text4.Text = DataGrid1.Columns(0)
Text5.Text = DataGrid1.Columns(1)
Text6.Text = DataGrid1.Columns(2)
text8.Text = Val(DataGrid1.Columns(3))
text9.Text = Val(DataGrid1.Columns(4))
Text7.Text = Val(DataGrid1.Columns(5))
End Sub

Sub kosong1()
Text1.Text = ""
Text2.Text = ""
'Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
text8.Text = ""
text9.Text = ""
End Sub



