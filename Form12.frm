VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form12 
   BackColor       =   &H0000C000&
   Caption         =   "Entri Penjualan"
   ClientHeight    =   10935
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   19080
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form12.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form12"
   ScaleHeight     =   10935
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox admin_approval 
      Height          =   240
      Left            =   1200
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   13920
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   13080
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
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
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   600
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   6360
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   11760
      Top             =   10440
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   11160
      Top             =   10440
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   17160
      TabIndex        =   21
      Text            =   " "
      Top             =   9960
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
      Caption         =   "HAPUS DATA DALAM TABEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   0
      TabIndex        =   18
      Top             =   2400
      Width           =   20715
      Begin MSComctlLib.ListView ListView2 
         Height          =   4335
         Left            =   4440
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7646
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kode"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "HargaJual"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sakhir"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Harga"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form12.frx":628A
         Height          =   6855
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   30000
         _ExtentX        =   52917
         _ExtentY        =   12091
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   8454016
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   2
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
            DataField       =   "JUMLAH_JUAL"
            Caption         =   "JML. JUAL"
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
            DataField       =   "HARGA_JUAL"
            Caption         =   "HRG. JUAL"
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
            DataField       =   "BAYAR"
            Caption         =   "BAYAR"
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
            DataField       =   "KASIR"
            Caption         =   "KASIR"
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
               ColumnWidth     =   5999.812
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   20730
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   34
         Text            =   "Text7"
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   240
         Width           =   7335
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13200
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9960
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Text            =   "Text6"
         Top             =   2040
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1095
         Left            =   2520
         TabIndex        =   7
         Top             =   2280
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   1931
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "KODE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "NAMA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Harga Jual"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Stok"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Harga"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSMask.MaskEdBox mb3 
         Height          =   495
         Left            =   14640
         TabIndex        =   8
         Top             =   1920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MB2 
         Height          =   495
         Left            =   4920
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MB1 
         Height          =   495
         Left            =   14880
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "harga jual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   17
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Harga Dasar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12120
         TabIndex        =   16
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Stok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H0080FF80&
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
      Left            =   19080
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text15"
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14880
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin MSMask.MaskEdBox mb4 
         Height          =   855
         Left            =   1200
         TabIndex        =   1
         Top             =   120
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1508
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Rp."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1920
      TabIndex        =   22
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   128253955
      CurrentDate     =   41271
   End
   Begin VB.PictureBox WebBrowser1 
      Height          =   375
      Left            =   17760
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   23
      Top             =   10440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   15360
      Top             =   10560
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
      OLEDBString     =   "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbbarang"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   14520
      Top             =   10560
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
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
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Caption         =   "F5-->Cetak Ulang Bill"
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   33
      Top             =   9960
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Caption         =   "F4-->Tutup"
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   32
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "JAM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   31
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "KASIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   16080
      TabIndex        =   30
      Top             =   9960
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Caption         =   "F3-->Delete "
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   29
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "F1 --> Cetak Bill"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   9960
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "NO. FAKTUR"
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
      Index           =   0
      Left            =   7680
      TabIndex        =   27
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "TANGGAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C000&
      Caption         =   "Label7"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   25
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Terima Kasih, Kasir Hari Ini :"
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
      Index           =   16
      Left            =   9960
      TabIndex        =   24
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim data1 As Long
Dim nama(1), alamat(100) As String

Dim rsBarang As New ADODB.Recordset
Dim rsbeli As New ADODB.Recordset
Public total_belanja As Long

Private Enum ImageSizingTypes
[sizeNone] = 0
[sizeCheckBox]
[sizeIcon]
End Enum

Private Enum LedgerColours
vbledgerWhite = &HF9FEFF
vbLedgerGreen = &HD0FFCC
vbLedgerYellow = &HE1FAFF
vbLedgerRed = &HE1E1FF
vbLedgerGrey = &HE0E0E0
vbLedgerBeige = &HD9F2F7
vbLedgerSoftWhite = &HF7F7F7
vbledgerPureWhite = &HFFFFFF
End Enum

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long


Private Sub Command1_Click()
kosong
End Sub

Private Sub Form_Load()
If con.State = adStateClosed Then
connect
End If
'sql = "select max(tbno.nomor) AS nob  From tbno"
'Set Rec = con.Execute(sql)
'    If IsNull(Rec!nob) = True Then
'       Text1.Text = 1
'    Else
'       Text1.Text = Rec!nob + 1
'    End If
'sql = "insert into tbno values('" & Val(Text1) & "')"
'con.Execute (sql)
kosong
DTPicker1 = Date
Adodc1.ConnectionString = "DSN=data"
sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah_jual` AS `jumlah_jual`,`j`.`harga_jual` AS `harga_jual`,`j`.`bayar` AS `bayar`,`j`.`nm_kasir` AS `kasir`,`j`.`nomor` AS `nomor` From (`tbbarang` `o` join `tbjual` `j`) Where (`j`.`kode` = `o`.`Kode`) and(`j`.`nobukti` = '" & Trim(text8 + Text1.Text) & "') order by `j`.`nomor` desc"
Adodc1.RecordSource = sql
Adodc1.Refresh
Text15 = username
Text16 = username

On Error GoTo ErrorFound
MSComm1.CommPort = 3
MSComm1.Settings = "9600,N,8,1"
MSComm1.PortOpen = True

ErrorFound:
    'nothing happens
On Error GoTo 0

'Membuka file untuk membaca
namafile = App.Path & "\faktur.txt"
Open namafile For Input As #1
'Membaca semua data file
'sampai data terakhir (End Of File)
While Not EOF(1)
'membaca data
    Input #1, data
'Menampilkan data di listbox
    kalimat = data
    text8 = Left(kalimat, 1)
    data1 = Mid(kalimat, 2, 20)
    Text1 = data1 + 1
    'List1.AddItem kalimat
Wend
'Menutup file
Close #1
End Sub
Private Sub Form_Activate()
Text7.SetFocus

If MSComm1.PortOpen Then
    MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
    MSComm1.Output = "Selamat Datang      Kasir: " + username
End If
End Sub
Private Sub kosong()
'Text2.Text = ""
Text7.Text = ""
Text3.Text = ""
Text4.Text = 1
Text5.Text = ""
MB1 = 0
mb3 = 0
mb4 = 0
Text15.Text = ""
End Sub
Private Sub kosongkan()
'Text2.Text = ""
Text7.Text = ""
Text3.Text = ""
Text4.Text = 1
Text5.Text = ""
MB1 = 0
MB2 = 0
mb3 = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sql = "delete from tbjual where nobukti='" & Trim(text8 + Text1.Text) & "'"
    Set rsObat = con.Execute(sql)
    yuyu = "delete from stock where  nobukti='" & Trim(text8 + Text1.Text) & "'"
    Set rsObat = con.Execute(yuyu)
    yuyu1 = "delete from bill where  nofaktur='" & Trim(text8 + Text1.Text) & "'"
    Set rsObat = con.Execute(yuyu1)
    Form9.refreshlist
    
     If MSComm1.PortOpen = True Then
      Do While MSComm1.OutBufferCount > 0
          DoEvents
       Loop
       MSComm1.PortOpen = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
 If KeyCode = 115 Then
     If MsgBox("Yakin Form Ini ditutup?", vbQuestion + vbYesNo, "Tutup Form") = vbYes Then
     If Adodc1.Recordset.RecordCount = 0 Then
        Unload Me
     Else
        sql = "delete from tbjual where nobukti='" & Trim(text8 + Text1.Text) & "'"
        Set rsObat = con.Execute(sql)
        yuyu = "delete from stock where  nobukti='" & Trim(text8 + Text1.Text) & "'"
        Set rsObat = con.Execute(yuyu)
        yuyu1 = "delete from bill where  nofaktur='" & Trim(text8 + Text1.Text) & "'"
        Set rsObat = con.Execute(yuyu1)
        Form9.refreshlist
        Unload Me
     End If
     Else
        Exit Sub
     
     End If
     End If
        If KeyCode = 113 Then
          'Call simpansaja
       End If
       If KeyCode = 114 Then
          Call Hapus
       End If
   
      If KeyCode = 112 Then
          Call cetakbill
       End If
       
       If KeyCode = 116 Then
           Form2.Show
       End If
End Sub
Private Sub cetakbill()
Form13.Show
Form13.mb10 = Form12.mb4
If MSComm1.PortOpen Then
    MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
    Form12.MSComm1.Output = "Total Belanja:      " + mb4.Text
End If
End Sub
Private Sub simpansaja()
If Trim(Text1) <> "" Then
               Set rsbantu1 = con.Execute("select * from tbbantu1 where nobukti='" & Trim(text8 + Text1.Text) & "'")
               If rsbantu1.EOF Then
                  con.Execute ("insert into tbbantu1 values('" & Trim(text8 + Text1.Text) & "','" & Format(DTPicker1, "yyyy-MM-dd") & "'," & Val(mb4) & "," & Val(text30) & "," & Val(text31) & ")")
               Else
                  sql1 = "Update tbbantu1 set tglbukti='" & Format(DTPicker1, "yyyy-MM-dd") & "',jml_bayar=" & Val(mb4) & ",jml_uang=" & Val(mb4) & ",kembali=" & Val(text31) & " where nobukti='" & Trim(text8 + Text1.Text) & "'"
                  con.Execute (sql1)
               End If
               'kosong
            Else
               MsgBox "Kode Tidak Boleh Kosong", vbYesNo + vbQuestion, "Confirm"
            End If
kosong


'Buka File untuk menyimpan data
namafile = App.Path & "\faktur.txt"
Open namafile For Output As #1
'Menyimpan semua data
'For I = 1 To Ndata
    Print #1, Trim(text8 + Text1)
'Next I
'Menutup file
Close #1

'sql = "select max(tbno.nomor) AS nob  From tbno"
'Set Rec = con.Execute(sql)
'    If IsNull(Rec!nob) = True Then
'       Text1.Text = 1
'    Else
'       Text1.Text = Rec!nob + 1
'    End If
'sql = "insert into tbno values('" & Val(Text1) & "')"
'con.Execute (sql)

Text1 = Text1 + 1

Text7.SetFocus
End Sub

Private Sub Timer1_Timer()
If Text16.ForeColor = vbRed Then
Text16.ForeColor = vbBlue
Else
Text16.ForeColor = vbRed
End If
End Sub
Private Sub Timer2_Timer()
Label7.Caption = Time
End Sub
Private Sub Hapus()
Dim sakhir1
If Adodc1.Recordset.RecordCount = 0 Then
Exit Sub
Else
    If MsgBox("Benar Data akan dihapus?", vbQuestion + vbYesNo, "Hapus") = vbYes Then
        sql = "delete from tbjual where nobukti='" & Trim(text8 + Text1.Text) & "' and nomor='" & DataGrid1.Columns(6) & "' "
        Set rsObat = con.Execute(sql)
        mb4 = Val(mb4) - Val(DataGrid1.Columns(4))
        yuyu = "delete from stock where  nobukti='" & Trim(text8 + Text1.Text) & "' and nomor='" & DataGrid1.Columns(6) & "'"
        Set rsObat = con.Execute(yuyu)
        yuyu1 = "delete from bill where  nofaktur='" & Trim(text8 + Text1.Text) & "' and nomor='" & DataGrid1.Columns(6) & "'"
        Set rsObat = con.Execute(yuyu1)
        Adodc1.Refresh
        Form9.refreshlist
 End If
End If
End Sub
Private Sub text1_change()
mb4 = 0
admin_approval.Value = 0
Adodc1.ConnectionString = "DSN=data"
sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah_jual` AS `jumlah_jual`,`j`.`harga_jual` AS `harga_jual`,`j`.`bayar` AS `bayar`,`j`.`nm_kasir` AS `kasir`,`j`.`nomor` AS `nomor` From (`tbbarang` `o` join `tbjual` `j`) Where (`j`.`kode` = `o`.`Kode`) and(`j`.`nobukti` = '" & Trim(text8 + Text1.Text) & "') order by `j`.`nomor` desc"
Adodc1.RecordSource = sql
Adodc1.Refresh
End Sub

Private Sub text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim strcari2 As String
Dim rstemp2 As ADODB.Recordset
Text7.Text = Replace(Text7.Text, " ", "")
Set rstemp2 = con.Execute("select * from tbbarang  where kode='" & Text7.Text & "' ")
With rstemp2
If (.EOF And .BOF) Then
    Text3.Text = ""
Else
  Text3.Text = !nama
End If
Text4.SelStart = 1
Text4.SetFocus
End With
End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sql1 = "select max(tbjual.nomor) AS nomor From tbjual"
Set Rec = con.Execute(sql1)
    If IsNull(Rec!nomor) = True Then
       Text6.Text = 1
    Else
       Text6.Text = Rec!nomor + 1
    End If
    
    Text4 = Int(Val(Text4))
Dim strcari3 As String
Dim rstemp3 As ADODB.Recordset
Set rstemp3 = con.Execute("select * from tbbarang  where kode='" & Trim(Text7.Text) & "' ")
With rstemp3
    If (.EOF And .BOF) Then
        Text3.Text = ""
        MB2 = ""
        Text5.Text = ""
        mb3 = ""
    Else
      Text3.Text = !nama
      MB2 = !harga_jual
      Text5.Text = !jumlah_akhir
      
      If (Text3 = "Diskon" And username <> "admin") Then
        MsgBox ("Kamu tidak bisa memberikan diskon!!")
        Exit Sub
      End If
      
      If (Len(Text4) > 3 And Text3 <> "Diskon") Then
        MsgBox ("Jumlah tidak valid!")
        Exit Sub
      End If
        
      If (Val(Text4.Text) < 1 And admin_approval.Value = 0) Then
        Form23.Show (1)
        If admin_approval.Value = 0 Then
            Text4.Text = "1"
            Exit Sub
        End If
      End If
    mb3 = !harga_modal
    MB1 = Val(Text4) * Val(MB2)
    mb4 = Val(mb4) + Val(MB1)
    
    If MSComm1.PortOpen Then
        MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
        Dim baris1, baris2 As String
        baris1 = Text4 + " " + Text3
        If Len(baris1) < 20 Then
           Do While (Len(baris1) < 20)
            baris1 = baris1 + " "
           Loop
        Else
            baris1 = Left$(baris1, 20)
        End If
        
        MSComm1.Output = baris1
          
        Dim spaces As Integer
        spaces = 20 - (Len(Format(MB1, "###,###,###")) + Len(Format(mb4, "###,###,###")) + 2)
        Do While (Len(baris2) < spaces)
            baris2 = baris2 + " "
        Loop
        baris2 = Format(MB1, "###,###,###") + baris2 + "(" + Format(mb4, "###,###,###") + ")"
        MSComm1.Output = baris2
    End If
    
    'Dim rsCheck As ADODB.Recordset
    'Set rsCheck = con.Execute("select * from tbjual  where nobukti='" & Trim(Text8 + Text1.Text) & "' and kode='" & Trim(Text7.Text) & "'  ")
    
     'If (rsCheck.EOF And rsCheck.BOF) Then
        sql3 = "insert into tbjual values('" & Trim(text8 + Text1.Text) & "','" & Format(DTPicker1, "yyyy-MM-dd") & "','" & Text7 & "'," & Val(mb3) & ",'" & Val(Text4) & "'," & Val(MB2) & ",'" & Trim(username) & "'," & Val(MB1) & ",'" & Format(DTPicker1, "yyyy-MM-dd") & "',0,0,0,'" & Val(Text6) & "',1)"
        yuyu3 = "insert into stock values('" & Format(DTPicker1, "yyyy-MM-dd") & "','" & Trim(text8 + Text1.Text) & "','" & Trim(Text7.Text) & "',0," & Val(Text4.Text) & ",'Penjualan','" & Trim(username) & "','" & Val(Text6) & "')"
        yuyu4 = "insert into bill values('" & Trim(text8 + Text1.Text) & "','" & Format(DTPicker1, "yyyy-MM-dd") & "','" & Label7 & "','" & Text3.Text & "','" & Val(Text4) & "'," & Val(MB1) & ",'" & Val(Text6) & "'," & Val(MB2) & ")"
    
    'Else
      'Dim nomor As Long
      'nomor = rsCheck.Fields("nomor")
      'sql3 = "update tbjual set jumlah_jual = jumlah_jual + '" & Val(Text4.Text) & "', bayar = bayar + '" & Val(MB1.Text) & "' where nomor='" & nomor & "' "
      'yuyu3 = "update stock set keluar = keluar + '" & Val(Text4.Text) & "' where nomor='" & nomor & "' "
      'yuyu4 = "update bill set jumlah_beli = jumlah_beli + '" & Val(Text4.Text) & "', bayar = bayar + '" & Val(MB1.Text) & "' where nomor='" & nomor & "' "
    
    'End If
    
    con.Execute (sql3)
    con.Execute (yuyu3)
    con.Execute (yuyu4)
    'con.Execute ("update stock set masuk=0,keluar=" & Val(Text4.Text) & " where tglbukti='" & Format(DTPicker1, "yyyy-MM-dd") & "' and nobukti='" & Trim(Text1.Text) & "' and nomor='" & DataGrid1.Columns(6) & "'")
    kosongkan
    Adodc1.Refresh
    Form9.refreshlist
    Text7.SetFocus
    Set rsCheck = Nothing
    End If
  End With
  End If
End Sub

Private Sub text7_LostFocus()
Dim strcari5 As String
Dim rstemp5 As ADODB.Recordset
Set rstemp5 = con.Execute("select * from tbbarang where kode='" & Text7.Text & "' ", , 1)
strcari5 = "kode_barang='" & Text7.Text & "' "
With rstemp5
If (.EOF And .BOF) Then
   Text3.Text = ""
   MB2 = ""
   Text5.Text = ""
   mb3 = ""
Else

  Text3.Text = !nama
  MB2 = !harga_jual
  Text5.Text = !jumlah_akhir
  mb3 = !harga_modal

      End If
  End With
End Sub


Private Sub ListView2_KeyPress(KeyAscii As Integer)
Text7 = ListView2.SelectedItem.SubItems(1)
Text3 = ListView2.SelectedItem.SubItems(2)
MB2 = ListView2.SelectedItem.SubItems(3)
Text5 = ListView2.SelectedItem.SubItems(4)
mb3 = ListView2.SelectedItem.SubItems(5)
'text13 = ListView2.SelectedItem.SubItems(6)
'text14 = ListView2.SelectedItem.SubItems(7)
Text4.SetFocus
ListView2.Visible = False
End Sub


Private Sub Text3_keypress(KeyAscii As Integer)

 With ListView2
    .Visible = False
    .Checkboxes = False
    .FullRowSelect = True
    Set .SmallIcons = Nothing

    'Call ListView2_KeyPress
    Call SetListViewLedger(ListView2, _
    vbLedgerYellow, _
    vbLedgerGrey, _
    sizeNone)
    .Refresh
    .Visible = True '/* Restore visibility
    End With

If KeyAscii = 13 Then
      cari3
End If
End Sub
Sub cari3()
On Error Resume Next
Dim strcari3 As String
strcari3 = "nama like '%" & Text3.Text & "%'"
Adodc2.Refresh
With Adodc2.Recordset
     .MoveFirst
     .Find strcari3
If .EOF Then
    MsgBox "Data Barang Belum Ada, Cek Lagi Data Barang Barangnya, Jika Belum Ada Entri", vbOKOnly + vbInformation, "Cek Kode Barang"
   Text3.SetFocus
Else
   ListView2.Visible = True
   RefreshData3
            End If
   End With
End Sub
Sub RefreshData3()
    Dim rsbarang3 As New ADODB.Recordset
    Dim LI As ListItem
    Dim no As Integer

    With ListView2.ListItems
        .Clear

        rsbarang3.Open "select * from tbbarang where nama like  '%" & Trim(Text3.Text) & "%'  Order by nama ASC", con, adOpenKeyset
        Do Until rsbarang3.EOF
        no = no + 1
            Set LI = .Add(, , no)
            LI.SubItems(1) = rsbarang3!kode
            LI.SubItems(2) = rsbarang3!nama
            LI.SubItems(3) = rsbarang3!harga_jual
            LI.SubItems(4) = rsbarang3!jumlah_awal
            LI.SubItems(5) = rsbarang3!jumlah_akhir
            rsbarang3.MoveNext
        Loop
         
    
           rsbarang3.Close
           Text4.SetFocus
    End With
    
    Call tutup3
End Sub
Private Sub tutup3()
ListView2.SetFocus
End Sub


Private Sub SetListViewLedger(lv As ListView, _
Bar1Color As LedgerColours, _
Bar2Color As LedgerColours, _
nSizingType As ImageSizingTypes)

Dim iBarHeight As Long
Dim lBarWidth As Long
Dim diff As Long
Dim twipsy As Long

iBarHeight = 0
lBarWidth = 0
diff = 0

On Local Error GoTo SetListViewColor_Error

twipsy = Screen.TwipsPerPixelY

If lv.View = lvwReport Then


With lv
.Picture = Nothing
.Refresh
.Visible = 1
.PictureAlignment = lvwTile
lBarWidth = .Width
End With ' lv


With Picture1
.AutoRedraw = False
.Picture = Nothing
.BackColor = vbWhite
.Height = 1
.AutoRedraw = True
.BorderStyle = vbBSNone
.ScaleMode = vbTwips
.Top = Form12.Top - 10000
.Width = Screen.Width
.Visible = False
.Font = lv.Font

With .Font
.Bold = lv.Font.Bold
.Charset = lv.Font.Charset
.Italic = lv.Font.Italic
.Name = lv.Font.Name
.Strikethrough = lv.Font.Strikethrough
.Underline = lv.Font.Underline
.Weight = lv.Font.Weight
.Size = lv.Font.Size
End With

iBarHeight = .TextHeight("W")

Select Case nSizingType
Case sizeNone:

iBarHeight = iBarHeight + twipsy

Case sizeCheckBox:
If (iBarHeight \ twipsy) > 18 Then
iBarHeight = iBarHeight + twipsy
Else
diff = 18 - (iBarHeight \ twipsy)
iBarHeight = iBarHeight + (diff * twipsy) + (twipsy * 1)
End If

Case sizeIcon:
diff = ImageList1.ImageHeight - (iBarHeight \ twipsy)
iBarHeight = iBarHeight + (diff * twipsy) + (twipsy * 1)

End Select

.Height = iBarHeight * 2
.Width = lBarWidth

Picture1.Line (0, 0)-(lBarWidth, iBarHeight), Bar1Color, BF
Picture1.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), Bar2Color, BF

.AutoSize = True
.Refresh

End With 'Picture1


lv.Refresh
lv.Picture = Picture1.Image

Else

lv.Picture = Nothing

End If 'lv.View = lvwReport

SetListViewColor_Exit:
On Local Error GoTo 0
Exit Sub

SetListViewColor_Error:

With lv
.Picture = Nothing
.Refresh
End With

Resume SetListViewColor_Exit
End Sub

