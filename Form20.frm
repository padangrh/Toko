VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form20 
   BackColor       =   &H0000C000&
   Caption         =   "Entri Return Beli"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12930
   ControlBox      =   0   'False
   Icon            =   "Form20.frx":0000
   LinkTopic       =   "Form20"
   ScaleHeight     =   7890
   ScaleWidth      =   12930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   6420
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   12465
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2040
         TabIndex        =   15
         Top             =   210
         Width           =   1785
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   2040
         TabIndex        =   14
         Text            =   " "
         Top             =   1080
         Width           =   2205
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Return"
         Height          =   465
         Left            =   11160
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3720
         Width           =   1155
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5400
         TabIndex        =   11
         Top             =   210
         Width           =   2505
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Text            =   "Text6"
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Text            =   "Text7"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Text            =   "Text10"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Text            =   "Text11"
         Top             =   3000
         Width           =   1335
      End
      Begin MSMask.MaskEdBox text12 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   3480
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
         TabIndex        =   10
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   129105923
         CurrentDate     =   41296
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form20.frx":628A
         Height          =   1920
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3387
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
         ColumnCount     =   5
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
            DataField       =   "JMLBELI"
            Caption         =   "JML. BELI"
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
            Caption         =   "HARGA"
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
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti Pembelian"
         Height          =   240
         Left            =   480
         TabIndex        =   25
         Top             =   300
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Suplier"
         Height          =   195
         Left            =   480
         TabIndex        =   24
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Suplier"
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   1605
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   240
         Left            =   480
         TabIndex        =   22
         Top             =   705
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Return"
         Height          =   240
         Left            =   4440
         TabIndex        =   21
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   19
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Akhir"
         Height          =   195
         Index           =   6
         Left            =   3480
         TabIndex        =   18
         Top             =   3120
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jml. Beli/Return"
         Height          =   195
         Index           =   7
         Left            =   480
         TabIndex        =   17
         Top             =   3120
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
         Height          =   195
         Index           =   8
         Left            =   480
         TabIndex        =   16
         Top             =   3600
         Width           =   435
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Simpan"
      Height          =   345
      Left            =   360
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   7080
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "&Tutup"
      Height          =   405
      Left            =   7560
      TabIndex        =   0
      Top             =   7080
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   10440
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Left            =   7830
      Top             =   6960
      Width           =   1305
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   420
      Left            =   480
      Top             =   6960
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   4080
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   26
      Top             =   7200
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsrbeli As New ADODB.Recordset

Private Sub Command2_Click()
kosongkan2
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Form_Activate()
Text1.SetFocus
End Sub
Private Sub Command4_Click()
Dim saldo, saldoAkhir, hrgbeli, hrgbelibaru
con.Execute ("insert into tbreturnbeli values(" & Trim(Text15) & ",'" & Text1.Text & "','" & Format(DTPicker1, "yyyy-MM-dd") & "','" & Text2.Text & "','" & Text6.Text & "','" & Val(Text10.Text) & "'," & Val(text12.Text) & ")")
con.Execute ("insert into stock values('" & Format(DTPicker1, "yyyy-MM-dd") & "','" & Text1.Text & "','" & Trim(Text6.Text) & "',0,'" & Val(Text10.Text) & "','Return Pembelian','" & Text2.Text & "',0)")
con.Execute ("update tbbeli set tanggal_r='" & Format(DTPicker1, "yyyy-MM-dd") & "', jumlah_r='" & Val(Text10) & "',harga_r=" & Val(text12) & " where nobukti='" & Trim(Text1) & "' and kdsuplier='" & Trim(Text2) & "' and kode='" & Text6.Text & "'")
kosongkan
sql = "select max(tbreturnbeli.noreturnbeli) AS nob  From tbreturnbeli"
Set Rec = con.Execute(sql)
    If IsNull(Rec!nob) = True Then
       Text15.Text = 1
    Else
       Text15.Text = Rec!nob + 1
    End If
Adodc1.Refresh
Form19.refreshlist
End Sub

Private Sub Form_Load()
kosongkan2
tgl = Date
If con.State = adStateClosed Then
connect
End If
sql = "select max(tbreturnbeli.noreturnbeli) AS nob  From tbreturnbeli"
Set Rec = con.Execute(sql)
    If IsNull(Rec!nob) = True Then
       Text15.Text = 1
    Else
       Text15.Text = Rec!nob + 1
    End If
Adodc1.ConnectionString = "dsn=data"
sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`o`.`kategori` AS `kategori`,`j`.`jumlah` AS `jmlbeli`,`j`.`harga` AS `harga`,`o`.`jumlah_akhir` AS `stok` From (`tbbarang` `o` join `tbbeli` `j`) Where (`j`.`kode` = `o`.`Kode`) and (`j`.`nobukti` = " & Val(Text1.Text) & ")"
Adodc1.RecordSource = sql
DTPicker1 = Date
End Sub
Private Sub text1_change()
sql1 = "select * from tbbeli,tbsuplier where tbsuplier.kdsuplier=tbbeli.kdsuplier and nobukti='" & Text1 & "'"
Set rsrbeli = con.Execute(sql1)
If Not rsrbeli.EOF Then
   Text2.Text = rsrbeli!kdsuplier
   Text3.Text = rsrbeli!nmsuplier

sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`o`.`kategori` AS `kategori`,`j`.`jumlah` AS `jmlbeli`,`j`.`harga` AS `harga`,`o`.`jumlah_akhir` AS `stok` From (`tbbarang` `o` join `tbbeli` `j`) Where (`j`.`kode` = `o`.`Kode`) and (`j`.`nobukti` = " & Val(Text1.Text) & ")"
Adodc1.RecordSource = sql
If Not rsrbeli.EOF Then
    rsrbeli.MoveFirst
    Do While Not rsrbeli.EOF
        rsrbeli.MoveNext
    Loop
End If
Adodc1.RecordSource = sql
Adodc1.Refresh
End If
End Sub

Sub kosongkan()
Text6.Text = ""
Text7.Text = ""
Text10.Text = 0
Text11.Text = 0
text12.Text = 0
End Sub
Sub kosongkan2()
Text2.Text = ""
Text3.Text = ""
Text6.Text = ""
Text7.Text = ""
Text10.Text = 0
Text11.Text = 0
text12.Text = 0
End Sub

Private Sub dataGrid1_DblClick()
Text6.Text = DataGrid1.Columns(0)
Text7.Text = DataGrid1.Columns(1)
Text10.Text = Val(DataGrid1.Columns(2))
text12.Text = Val(DataGrid1.Columns(3))
Text11.Text = Val(DataGrid1.Columns(4))
End Sub



