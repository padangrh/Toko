VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form10 
   BackColor       =   &H0000C000&
   Caption         =   "Entri Pembatalan  Penjualan"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16380
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form10"
   ScaleHeight     =   7170
   ScaleWidth      =   16380
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   14280
      TabIndex        =   17
      Top             =   0
      Width           =   7695
      Begin MSMask.MaskEdBox mb4 
         Height          =   855
         Left            =   1200
         TabIndex        =   18
         Top             =   120
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
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
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H000080FF&
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
      Left            =   17400
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text15"
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   600
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   21975
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   9840
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Height          =   495
         Left            =   13200
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   4560
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   240
         Width           =   8055
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   873
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
         TabIndex        =   9
         Top             =   1560
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
         Top             =   1560
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MB1 
         Height          =   495
         Left            =   14880
         TabIndex        =   11
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Caption         =   "Stok"
         Height          =   495
         Left            =   8040
         TabIndex        =   14
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Harga Dasar"
         Height          =   375
         Left            =   12000
         TabIndex        =   13
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "harga jual"
         Height          =   495
         Left            =   3720
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   21135
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "form10.frx":0000
         Height          =   7095
         Left            =   240
         TabIndex        =   32
         Top             =   120
         Width           =   20055
         _ExtentX        =   35375
         _ExtentY        =   12515
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         BackColor       =   8454016
         ForeColor       =   0
         HeadLines       =   2
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
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
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000C000&
      Caption         =   "HAPUS DATA DALAM TABEL"
      Height          =   375
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   15120
      TabIndex        =   0
      Text            =   " "
      Top             =   9960
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   11760
      Top             =   10800
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   12360
      Top             =   10680
   End
   Begin VB.PictureBox Cr 
      Height          =   480
      Left            =   16080
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   33
      Top             =   10680
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2400
      TabIndex        =   20
      Top             =   600
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
      Format          =   108724227
      CurrentDate     =   41271
   End
   Begin VB.PictureBox WebBrowser1 
      Height          =   375
      Left            =   15360
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   14160
      Top             =   10680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=minimarket"
      OLEDBString     =   "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=minimarket"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbbarang"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12960
      Top             =   10680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Terima Kasih, Kasir Hari Ini :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   16560
      TabIndex        =   31
      Top             =   10800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C000&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   30
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "TANGGAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   29
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "NO. FAKTUR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   28
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C000&
      Caption         =   "F1 --> Cetak Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   9960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C000&
      Caption         =   "F3-->Delete "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   26
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "KASIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   13920
      TabIndex        =   25
      Top             =   9960
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "JAM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   24
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C000&
      Caption         =   "F4-->Tutup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   23
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C000&
      Caption         =   "F2 -->Tekan Tombol  F2 Jika Tidak Cetak Bill "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   22
      Top             =   9960
      Visible         =   0   'False
      Width           =   4935
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsbarang As New ADODB.Recordset
Dim rsbeli As New ADODB.Recordset
Dim rsJual As String
Private Sub Command1_Click()
kosong
End Sub

Private Sub Form_Load()
If con.State = adStateClosed Then
connect
End If

kosong
DTPicker1 = Date
Adodc1.ConnectionString = "DSN=data"
sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah_jual` AS `jumlah_jual`,`j`.`harga_jual` AS `harga_jual`,`j`.`bayar` AS `bayar`,`j`.`nm_kasir` AS `kasir`,`j`.`nomor` AS `nomor` From (`tbbarang` `o` join `tbjual` `j`) Where (`j`.`kode` = `o`.`Kode`) and(`j`.`nobukti` = '" & Trim(Text1.Text) & "')"
Adodc1.RecordSource = sql
Adodc1.Refresh
Text16 = username
End Sub

Private Sub kosong()
Text2.Text = ""
Text3.Text = ""
Text4.Text = 1
Text5.Text = ""
MB1 = 0
mb3 = 0
mb4 = 0
Text15.Text = ""
End Sub
Private Sub kosongkan()
Text2.Text = ""
Text3.Text = ""
Text4.Text = 1
Text5.Text = ""
MB1 = 0
MB2 = 0
mb3 = 0
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 115 Then
     If MsgBox("Yakin Form Ini ditutup?", vbQuestion + vbYesNo, "Tutup Form") = vbYes Then
            Unload Me
        Else
            Exit Sub
        End If
     End If
        If KeyCode = 113 Then
         ' Call simpansaja
       End If
       If KeyCode = 114 Then
          Call Hapus
       End If
   
      If KeyCode = 112 Then
          'Call cetakbill
       End If
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
        sql = "delete from tbjual where nobukti='" & Trim(Text1) & "' and nomor='" & DataGrid1.Columns(6) & "' "
        Set rsObat = con.Execute(sql)
        mb4 = Val(mb4) - Val(DataGrid1.Columns(4))
        yuyu = "delete from stock where  nobukti='" & Trim(Text1.Text) & "' and nomor='" & DataGrid1.Columns(6) & "'"
        Set rsObat = con.Execute(yuyu)
        yuyu1 = "delete from bill where  nofaktur='" & Trim(Text1.Text) & "' and nomor='" & DataGrid1.Columns(6) & "'"
        Set rsObat = con.Execute(yuyu1)
        Adodc1.Refresh
        Form7.refreshlist
 End If
End If
End Sub

Private Sub text1_change()
sql1 = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah_jual` AS `jumlah_jual`,`j`.`harga_jual` AS `harga_jual`,`j`.`bayar` AS `bayar`,`j`.`nm_kasir` AS `kasir` From (`tbbarang` `o` join `tbjual` `j`) Where (`j`.`kode` = `o`.`Kode`) and(`j`.`nobukti` = '" & Trim(Text1.Text) & "')"
Set rsbarang = con.Execute(sql1)
If Not rsbarang.EOF Then
    rsbarang.MoveFirst
    Do While Not rsbarang.EOF
    mb4 = Val(mb4) + rsbarang!bayar
    rsbarang.MoveNext
    Loop
End If


Adodc1.ConnectionString = "DSN=data"
sql = "select `j`.`kode` AS `kode`,`o`.`Nama` AS `Nama`,`j`.`jumlah_jual` AS `jumlah_jual`,`j`.`harga_jual` AS `harga_jual`,`j`.`bayar` AS `bayar`,`j`.`nm_kasir` AS `kasir`,`j`.`nomor` AS `nomor` From (`tbbarang` `o` join `tbjual` `j`) Where (`j`.`kode` = `o`.`Kode`) and(`j`.`nobukti` = '" & Trim(Text1.Text) & "')"
Adodc1.RecordSource = sql
Adodc1.Refresh
End Sub


