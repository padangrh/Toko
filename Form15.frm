VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form15 
   BackColor       =   &H0000C000&
   Caption         =   "Entri dan Update Pembelian"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19410
   ControlBox      =   0   'False
   Icon            =   "Form15.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Entri dan Update Pembelian"
   ScaleHeight     =   9465
   ScaleWidth      =   19410
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2400
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   6840
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   45
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   7005
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   18975
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9360
         TabIndex        =   44
         Text            =   "Text5"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9360
         TabIndex        =   42
         Text            =   "Text4"
         Top             =   1080
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2415
         Left            =   1800
         TabIndex        =   39
         Top             =   2400
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kode Barang"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama Barang"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Kode Suplier"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Harga Modal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Kategori"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10680
         TabIndex        =   40
         Text            =   "Text3"
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
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
         Height          =   405
         Left            =   9360
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
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
         Height          =   360
         Left            =   1800
         TabIndex        =   20
         Top             =   120
         Width           =   2640
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Form15.frx":628A
         Left            =   1800
         List            =   "Form15.frx":628C
         TabIndex        =   18
         Text            =   "T | Lunas"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   3930
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   18735
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "Form15.frx":628E
            Height          =   3465
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   18495
            _ExtentX        =   32623
            _ExtentY        =   6112
            _Version        =   393216
            BackColor       =   12648447
            ForeColor       =   16711680
            HeadLines       =   1
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "KODE"
               Caption         =   "KODE BARANG"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
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
                  LCID            =   1057
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
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "RETURN"
               Caption         =   "RETUR"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "#,##.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "HARGA"
               Caption         =   "HARGA"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   """Rp""#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "BAYAR"
               Caption         =   "BAYAR"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   """Rp""#.##0,00"
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
                  Alignment       =   1
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   4004.788
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   2160
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1995.024
               EndProperty
            EndProperty
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form15.frx":62A3
         Left            =   5400
         List            =   "Form15.frx":62A5
         TabIndex        =   14
         Text            =   "0"
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "ADD KEDALAM TABEL"
         Height          =   495
         Left            =   15000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000080FF&
         Caption         =   "HAPUS DALAM TABEL"
         Height          =   495
         Left            =   17280
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Text            =   "Text6"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
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
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Text            =   "Text7"
         Top             =   2520
         Width           =   5055
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   8
         Text            =   "Text12"
         Top             =   2040
         Width           =   1575
      End
      Begin MSMask.MaskEdBox text21 
         Height          =   375
         Left            =   9360
         TabIndex        =   6
         Top             =   2520
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text13 
         Height          =   375
         Left            =   9360
         TabIndex        =   7
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   65280
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
         Format          =   "#,##.00"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   108593155
         CurrentDate     =   41289
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5400
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   108593155
         CurrentDate     =   41289
      End
      Begin MSComCtl2.DTPicker dtpicker1 
         Height          =   405
         Left            =   1800
         TabIndex        =   19
         Top             =   600
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   714
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   108527619
         CurrentDate     =   39459
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Beli Bersih"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   7680
         TabIndex        =   43
         Top             =   2160
         Width           =   1560
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sisa Stok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   7680
         TabIndex        =   41
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Suplier "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7680
         TabIndex        =   37
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Beli"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type Beli"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl.Pembayaran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Jatuh Tempo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lama Kredit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3600
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   1440
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Beli"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   7680
         TabIndex        =   23
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Beli"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   7680
         TabIndex        =   22
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bayar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   7680
         TabIndex        =   21
         Top             =   2640
         Width           =   630
      End
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   8520
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSMask.MaskEdBox text19 
      Height          =   375
      Left            =   14880
      TabIndex        =   0
      Top             =   8280
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##.00"
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   3000
      Top             =   9120
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
      RecordSource    =   "tbbantu11"
      Caption         =   "Adodc5"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   3000
      Top             =   9600
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
      RecordSource    =   "tbbeli"
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   840
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   840
      Top             =   9600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "tbsuplier"
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
   Begin MSMask.MaskEdBox text17 
      Height          =   375
      Left            =   14880
      TabIndex        =   2
      Top             =   8760
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text16 
      Height          =   375
      Left            =   14880
      TabIndex        =   3
      Top             =   7800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text15 
      Height          =   375
      Left            =   14880
      TabIndex        =   4
      Top             =   7200
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##.00"
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   840
      Top             =   9120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":62A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":70F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":754B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":7865
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":64B6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":65449
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":6589B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":C2BA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":C347F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":120789
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":121F1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":16A7B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":16C4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":16CD99
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":16DA73
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":16E74D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form15.frx":16F427
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "TOTAL [Rp.]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   11760
      TabIndex        =   36
      Top             =   7320
      Width           =   1320
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "BAYAR (Jika Ada DP)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   11760
      TabIndex        =   35
      Top             =   7800
      Width           =   2280
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "ISI JIKA LUNAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   11760
      TabIndex        =   34
      Top             =   8760
      Width           =   1620
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080FF80&
      Caption         =   "F2 -->SIMPAN      F3 -->DELETE     F4 --> KELUAR      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   7320
      Width           =   6375
   End
   Begin VB.Label Label17 
      BackColor       =   &H0000C000&
      Caption         =   "HUTANG SETELAH KURANG DP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   32
      Top             =   8280
      Width           =   3015
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsbarang, rsbeli As New ADODB.Recordset
Dim tot, bayar, jumlah1 As Double
Dim is_empty As Boolean
Dim nama_suplier As String
Dim kode_suplier As Integer

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


Private Sub Command2_Click()
If Adodc1.Recordset.RecordCount = 0 Then
Exit Sub
Else
If MsgBox("Benar Data akan dihapus?", vbQuestion + vbYesNo, "Hapus") = vbYes Then
Text15 = Val(Text15) - Val(DataGrid1.Columns(5))
sql = "delete from tbbeli where nobukti='" & Text1 & "' and kode='" & DataGrid1.Columns(0) & "'"
Set rsbarang = con.Execute(sql)
con.Execute (sql)
yuyu = "delete from stock where  nobukti='" & Trim(Text1.Text) & "'"
Set rsbarang = con.Execute(yuyu)
con.Execute (yuyu)
Adodc1.Refresh
Form14.refreshlist
End If
End If
End Sub

Private Sub Form_Load()
is_empty = True
nama_suplier = ""
kode_suplier = 0
Combo1.AddItem "L |Lunas"
Combo1.AddItem "B |Bon"
'Combo1.AddItem "L |Lunas"
Combo1.ListIndex = 0

Combo2.AddItem "0"
Combo2.AddItem "7"
Combo2.AddItem "8"
Combo2.AddItem "9"
Combo2.AddItem "10"
Combo2.AddItem "11"
Combo2.AddItem "12"
Combo2.AddItem "13"
Combo2.AddItem "14"
Combo2.AddItem "15"
Combo2.AddItem "16"
Combo2.AddItem "17"
Combo2.AddItem "18"
Combo2.AddItem "19"
Combo2.AddItem "20"
Combo2.AddItem "21"
Combo2.AddItem "22"
Combo2.AddItem "23"
Combo2.AddItem "24"
Combo2.AddItem "25"
Combo2.AddItem "26"
Combo2.AddItem "27"
Combo2.AddItem "28"
Combo2.AddItem "29"
Combo2.AddItem "30"
Combo2.AddItem "31"
Combo2.AddItem "32"
Combo2.AddItem "33"
Combo2.AddItem "34"
Combo2.AddItem "35"
Combo2.AddItem "36"
Combo2.AddItem "37"
Combo2.AddItem "38"
Combo2.AddItem "39"
Combo2.AddItem "40"
Combo2.AddItem "41"
Combo2.AddItem "42"
Combo2.AddItem "43"
Combo2.AddItem "44"
Combo2.AddItem "45"
Combo2.AddItem "46"
Combo2.AddItem "47"
Combo2.AddItem "48"
Combo2.AddItem "49"
Combo2.AddItem "50"
Combo2.AddItem "51"
Combo2.AddItem "52"
Combo2.AddItem "53"
Combo2.AddItem "54"
Combo2.AddItem "55"
Combo2.AddItem "56"
Combo2.AddItem "57"
Combo2.AddItem "58"
Combo2.AddItem "59"
Combo2.AddItem "60"
Combo2.ListIndex = 0

If con.State = adStateClosed Then
connect
End If

namafile = App.Path & "\fakturbeli.txt"
    Open namafile For Input As #1
    'Membaca semua data file
    'sampai data terakhir (End Of File)
    While Not EOF(1)
    'membaca data
        Input #1, data
    'Menampilkan data di listbox
        kalimat = data
        Dim huruf As String
        Dim angka As Long
        huruf = Left(kalimat, 1)
        angka = Val(Mid(kalimat, 2, 20))
        Text1 = huruf + CStr(angka + 1)
        'List1.AddItem kalimat
    Wend
    'Menutup file
    Close #1

Adodc1.ConnectionString = "dsn=data"
sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah` AS `jumlah`,`j`.`return` AS `return`,`j`.`harga` AS `harga`,`j`.`bayar` AS `bayar` From (`tbbarang` `o` join `tbbeli` `j`) Where (`j`.`kode` = `o`.`Kode`) and (`j`.`nobukti` = '" & Trim(Text1) & "')"
Adodc1.RecordSource = sql
Adodc1.Refresh
DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
kosong
End Sub

Private Sub text1_change()
'text15 = 1000
    sql1 = "select * from tbsuplier, tbbeli,tbbantu11 where tbsuplier.kdsuplier=tbbeli.kdsuplier and tbbeli.nobukti='" & Trim(Text1) & "' and tbbantu11.nobukti='" & Trim(Text1) & "'"
    Set rsbeli = con.Execute(sql1)
    If Not rsbeli.EOF Then
        DTPicker1 = rsbeli!tglbukti
        Combo1.Text = rsbeli!tk
        Combo2.Text = rsbeli!lama
        DTPicker2 = rsbeli!jatuh
        DTPicker3 = rsbeli!tglbayar
        Text2.Text = rsbeli!kdsuplier
        Text16.Text = rsbeli!dp
        text19.Text = rsbeli!sisa
        text17.Text = rsbeli!lunas
        Else
        DTPicker1 = Date
        'Combo1.Text = ""
        Combo2.Text = ""
        DTPicker2 = Date
        DTPicker3 = Date
        Text2.Text = ""
        Text16.Text = 0
        text19.Text = 0
        text17.Text = 0
        
sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah` AS `jumlah`,`j`.`return` AS `return`,`j`.`harga` AS `harga`,`j`.`bayar` AS `bayar` From (`tbbarang` `o` join `tbbeli` `j`) Where (`j`.`kode` = `o`.`Kode`) and (`j`.`nobukti` = '" & Trim(Text1) & "')"
Set rsbarang = con.Execute(sql)
End If

sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah` AS `jumlah`,`j`.`return` AS `return`,`j`.`harga` AS `harga`,`j`.`bayar` AS `bayar` From (`tbbarang` `o` join `tbbeli` `j`) Where (`j`.`kode` = `o`.`Kode`) and (`j`.`nobukti` = '" & Trim(Text1) & "')"
Set rsbarang = con.Execute(sql)
If Not rsbarang.EOF Then
    rsbarang.MoveFirst
    Do While Not rsbarang.EOF
       Text15 = Val(Text15) + rsbarang!bayar
       rsbarang.MoveNext
    Loop
End If

Adodc1.ConnectionString = "DSN=data"
sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah` AS `jumlah`,`j`.`return` AS `return`,`j`.`harga` AS `harga`,`j`.`bayar` AS `bayar` From (`tbbarang` `o` join `tbbeli` `j`) Where (`j`.`kode` = `o`.`Kode`) and (`j`.`nobukti` = '" & Trim(Text1) & "')"
Adodc1.RecordSource = sql
Adodc1.Refresh
End Sub
Private Sub kosong()
'Combo1 = ""
Combo2 = 0
Combo3 = 0
Text3 = ""
Text4 = 0
Text5 = ""
Text2 = ""
Text6 = ""
Text7 = ""
text8 = 0
text9 = ""
Text10 = ""
Text11 = ""
text12 = ""
Text15 = 0
Text16 = 0
text17 = 0
Text18 = 0
text19 = 0
text20 = 0
text21 = 0
End Sub
Private Sub kosong2()
Text4 = 0
Text5 = ""
Text3 = ""
Text6 = ""
Text7 = ""
text9 = ""
Text10 = ""
Text11 = ""
text12 = ""
text13 = 0
text14 = 0
text8 = 0
text20 = 0
text21 = 0
Combo2 = 0
End Sub
'Private Sub ListView1_KeyPress(KeyAscii As Integer)
'Text2 = ListView1.SelectedItem.SubItems(1)
'Text6.SetFocus
'ListView1.Visible = False
'End Sub


Private Sub Form_Activate()
     'Membuka file untuk membaca
    
    
    Text6.SetFocus
End Sub
'Private Sub text2_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'      cari1
'End If
'End Sub
'Sub cari1()
'On Error Resume Next
'Dim strcari As String
'strcari = "kdsuplier like '%" & Text2.Text & "%'"
'Adodc2.Refresh
'With Adodc2.Recordset
'     .MoveFirst
'     .Find strcari
'If .EOF Then
'   MsgBox "Suplier Baru....Silakan Enter!!", vbYesNo + vbQuestion, "Confirm"
'   Text3.SetFocus
'Else
'   ListView1.Visible = True
'   RefreshData
'            End If
'   End With
'End Sub

'Sub RefreshData()
'    Dim rsSuplier As New ADODB.Recordset
'    Dim LI As ListItem
'    Dim no As Integer
'
'    With ListView1.ListItems
'        .Clear
'
'        rsSuplier.Open "select * from tbsuplier where kdsuplier like  '%" & Trim(Text2.Text) & "%'  Order by kdsuplier ASC", con, adOpenKeyset
'        Do Until rsSuplier.EOF
'        no = no + 1
'            Set LI = .Add(, , no)
'            LI.SubItems(1) = rsSuplier!kdsuplier
'            LI.SubItems(2) = rsSuplier!nmsuplier
'            rsSuplier.MoveNext
'        Loop
'
'           rsSuplier.Close
'           Text6.SetFocus
'    End With
'    Call cari
'End Sub
'Private Sub cari()
'ListView1.SetFocus
'End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
Text6 = ListView2.SelectedItem.SubItems(1)
Text7 = ListView2.SelectedItem.SubItems(2)
Text2 = ListView2.SelectedItem.SubItems(3)
text13 = ListView2.SelectedItem.SubItems(4)
Text3 = ListView2.SelectedItem.SubItems(5)
'Text4 = ListView2.SelectedItem.SubItems(6)
Text4.SetFocus
ListView2.Visible = False
End Sub



Private Sub Text6_keypress(KeyAscii As Integer)

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
      cari2
End If
End Sub
Sub cari2()
On Error Resume Next
Dim strcari As String
strcari = "kode like '%" & Text6.Text & "%'"
Adodc3.Refresh
With Adodc3.Recordset
     .MoveFirst
     .Find strcari
If .EOF Then
   MsgBox "Barang Belum tersedia, silakan Keluar dari Sub Menu ini dan masuk ke Input Barang Baru", vbYesNo + vbQuestion, "Confirm"
   Text7.SetFocus
Else
   ListView2.Visible = True
   RefreshData1
            End If
   End With
End Sub

Sub RefreshData1()
    Dim rsbarang As New ADODB.Recordset
    Dim LI As ListItem
    Dim no As Integer
    
    With ListView2.ListItems
        .Clear
        
        rsbarang.Open "select * from tbbarang where kode like  '%" & Trim(Text6.Text) & "%'  Order by kode ASC", con, adOpenKeyset
        Do Until rsbarang.EOF
        no = no + 1
            Set LI = .Add(, , no)
            LI.SubItems(1) = rsbarang!kode
            LI.SubItems(2) = rsbarang!nama
            LI.SubItems(3) = rsbarang!kdsuplier
            LI.SubItems(4) = rsbarang!harga_modal
            LI.SubItems(5) = rsbarang!kategori
            'LI.SubItems(6) = rsbarang!jumlah_akhir
            
           
            rsbarang.MoveNext
        Loop
           
           rsbarang.Close
           Text4.SetFocus
    End With
    Call cari11
End Sub
Private Sub cari11()
ListView2.SetFocus
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text6.SetFocus
End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5.SetFocus
End If
End Sub
Private Sub text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    jumlah1 = Val(Text5) - Val(Text4)
    text12 = jumlah1
    bayar = Val(text12) * Val(text13)
    text21 = bayar
    Command1.SetFocus
End If
End Sub

'Private Sub text12_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'
'    bayar = Val(Text12) * Val(text13)
'    text21 = bayar
'    Command1.SetFocus
'End If
'End Sub

Private Sub Command1_Click()
    Dim saldo, saldoAkhir, hrgbeli, hrgbelibaru
    If Text1.Text = "" Or Text2.Text = "" Or Text6.Text = "" Or text12.Text = "" Or Combo1.Text = "" Then
        MsgBox "Data Tidak Lengkap.....!", vbYesNo + vbQuestion, "Confirm"
        Exit Sub
    End If
    
    Dim rsSuplier As ADODB.Recordset
    Set rsSuplier = con.Execute("select * from tbsuplier where kdsuplier = '" & Val(Text2) & "'")
    If Not (rsSuplier.EOF Or rsSuplier.BOF) Then
        If nama_suplier = "" Then
            nama_suplier = rsSuplier!nmsuplier
            kode_suplier = rsSuplier!kdsuplier
        End If
    End If
    
    sql = "insert into tbbeli values('" & Text1.Text & "','" & Format(DTPicker1, "yyyy-MM-dd") & "','" & kode_suplier & "','" & Text6.Text & "'," & Val(Text5.Text) & ",'" & Val(text13) & "'," & Val(text21) & ",'" & Left(Combo1.Text, 1) & "','" & Format(DTPicker2, "yyyy-MM-dd") & "','" & Val(Combo2) & "','" & Format(DTPicker1, "yyyy-MM-dd") & "',0,0,'" & Val(Text4) & "', '" & Text7 & "',0,0)"
    beli = ("update tbbarang set jumlah_akhir= jumlah_akhir+" & Val(text12) & " where kode='" & Text6.Text & "'")
    con.Execute (sql)
    con.Execute (beli)
    'saldoAkhir = saldo + Val(Text12)
    is_empty = False
    sql = "select * from tbbarang where kode='" & Text6.Text & "'"
    Set Rec = con.Execute(sql)
    Text15 = Val(Text15) + Val(text21)
    sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah` AS `jumlah`,`j`.`return` AS `return`,`j`.`harga` AS `harga`,`j`.`bayar` AS `bayar` From (`tbbarang` `o` join `tbbeli` `j`) Where (`j`.`kode` = `o`.`Kode`) and (`j`.`nobukti` = '" & Trim(Text1) & "')"
    Adodc1.RecordSource = sql
    Adodc1.Refresh
    Form14.refreshlist
    Text6.SetFocus
    
    kosong2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 115 Then
     If MsgBox("Yakin Form Ini ditutup?", vbQuestion + vbYesNo, "Tutup Form") = vbYes Then
            con.Execute ("delete from tbbeli where nobukti='" & Text1 & "'")
            con.Execute ("delete from tbbantu11 where nobukti='" & Text1 & "'")
            Unload Me
        Else
            Exit Sub
        End If
         End If
          If KeyCode = 114 Then
          Call Hapus
      End If
        If KeyCode = 113 Then
          Call simpan
      End If
     
         
End Sub
Private Sub simpan()

If Trim(Text1) <> "" Then
        
       Set rsbeli = con.Execute("select * from tbbeli where nobukti='" & Trim(Text1) & "'")
       If Not rsbeli.EOF Then
        rsbeli.MoveFirst
        Printer.Font = "courier new"
        Printer.CurrentX = 0
        Printer.CurrentY = 0
        Printer.FontSize = 17
        Printer.FontBold = True
        Printer.Print Tab(2); " BON PEMBELIAN";
        Printer.Print Tab(2); "CHRISTINE HAKIM";
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.Print Tab(2); "                                                            "; 'baris kosong
        Printer.Print Tab(3); "Jl. Adinegoro No. 11A Padang";
        Printer.Print Tab(2); "                                                       ";
        Printer.Print Tab(2); "No. FAKTUR : "; Text1;
        Printer.Print Tab(2); "Supplier: "; CStr(kode_suplier); "/"; nama_suplier;
        Printer.Print Tab(2); Format(Date, "dd/mm/yyyy"); "  "; Time;
        Printer.Print Tab(2); "                                                       ";
        
        Dim total As Long
        total = 0
        Do While Not rsbeli.EOF
           con.Execute ("update tbbarang set tgl_masuk='" & Format(rsbeli.Fields("tglbukti"), "yyyy-mm-dd") & "' where kode='" & rsbeli("kode") & "'")
           Printer.Print Tab(2); rsbeli.Fields("nama_barang");
           Dim jumlah_awal As Integer
           Printer.Print Tab(2); rsbeli.Fields("jumlah"); "-("; rsbeli.Fields("return"); ")x"; Format(rsbeli.Fields("harga"), "###,###,###"); Tab(22); Format(rsbeli.Fields("bayar"), "###,###,###");
           total = total + rsbeli.Fields("bayar")
           rsbeli.MoveNext
        Loop
        
        Printer.Print Tab(2); "                                                       ";
        Printer.FontBold = True
        Printer.Print Tab(2); "Total        "; Tab(20); Format(total, "###,###,###");
        Printer.FontBold = False
        Printer.Print Tab(2); "                                                       ";
        Printer.Print Tab(2); "Diterima oleh"; Tab(18); "Dibayar oleh";
        Printer.Print Tab(2); "                                                       ";
        Printer.Print Tab(2); "                                                       ";
        Printer.Print Tab(2); "                                                       ";
        Printer.Print Tab(2); "(___________)"; Tab(18); "(___________)";
        Printer.Print Tab(2); "                                                       ";
        Printer.Print Tab(2); "        *Terima Kasih*";
        Printer.EndDoc
        End If
        
       Set rsbantu11 = con.Execute("select * from tbbantu11 where nobukti='" & Trim(Text1) & "'")
       If rsbantu11.EOF Then
          con.Execute ("insert into tbbantu11 values('" & Text1.Text & "','" & Format(DTPicker1, "yyyy-MM-dd") & "','" & Left(Combo1.Text, 1) & "','" & Text2 & "'," & Val(Text15) & ",'" & Format(DTPicker2, "yyyy-MM-dd") & "'," & Val(text17) & ",'" & Format(DTPicker3, "yyyy-MM-dd") & "'," & Val(Text16) & "," & Val(text19) & ")")
       Else
          sql1 = "Update tbbantu11 set tglbukti='" & Format(DTPicker1, "yyyy-MM-dd") & "',tk='" & Left(Combo1.Text, 1) & "',kdsuplier='" & Text2 & "', jumlah=" & Val(Text15) & ",jatuh='" & Format(DTPicker2, "yyyy-MM-dd") & "',lunas=" & Val(text17) & ",tglbayar='" & Format(DTPicker3, "yyyy-MM-dd") & "',dp=" & Val(Text16) & ",sisa=" & Val(text19) & " where nobukti='" & Text1.Text & "'"
          con.Execute (sql1)
       End If
       kosong
    Else
       MsgBox "Kode Tidak Boleh Kosong", vbYesNo + vbQuestion, "Confirm"
    End If
        
    namafile = App.Path & "\fakturbeli.txt"
    Open namafile For Output As #1
    'Menyimpan semua data
    'For I = 1 To Ndata
        Print #1, Trim(Text1.Text)
    'Next I
    'Menutup file
    Close #1
    
    Unload Me
'Text1.SetFocus
End Sub
Private Sub Hapus()
 Dim X As String
    X = MsgBox("Ingin Hapus data ini?[Y/N]", vbYesNo + vbQuestion)
    If X = vbYes Then
    con.Execute ("delete from tbbantu11 where nobukti='" & Trim(Text1) & "'")
    con.Execute ("delete from tbbeli where nobukti='" & Trim(Text1) & "'")
    End If
    kosong
   
End Sub

Private Sub text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      If Val(Text16) >= 0 And Val(Text16) < Val(Text15) Then
       text19 = Val(Text15) - Val(Text16)
       text17.SetFocus
       Else
       MsgBox "DP Besar dari Harga Cek Lagi", vbYesNo + vbQuestion, "Confirm"
       End If
End If
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
.Top = Form15.Top - 10000
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







