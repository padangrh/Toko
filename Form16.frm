VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form16 
   BackColor       =   &H0000C000&
   Caption         =   "Update Return Jual"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8820
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   ScaleHeight     =   5700
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   570
      TabIndex        =   3
      Top             =   360
      Width           =   7665
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4680
         TabIndex        =   24
         Text            =   "Text3"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4680
         TabIndex        =   10
         Text            =   "Text11"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Text            =   "Text8"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Text            =   "Text7"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Text            =   "Text6"
         Top             =   1920
         Width           =   4695
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Text            =   "Text5"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5415
         TabIndex        =   5
         Text            =   "Auto"
         Top             =   210
         Width           =   1785
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   1785
      End
      Begin MSMask.MaskEdBox text10 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   3360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text9 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   2880
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##.00"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker tgl 
         Height          =   405
         Left            =   2025
         TabIndex        =   13
         Top             =   645
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   108855299
         CurrentDate     =   39459
      End
      Begin VB.Line Line1 
         X1              =   3240
         X2              =   3120
         Y1              =   2400
         Y2              =   2760
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stok Akhir"
         Height          =   195
         Index           =   3
         Left            =   3360
         TabIndex        =   22
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modal"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   21
         Top             =   3480
         Width           =   435
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jual"
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   19
         Top             =   3000
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Jual/ Return"
         Height          =   195
         Left            =   525
         TabIndex        =   18
         Top             =   2400
         Width           =   1425
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         Height          =   195
         Left            =   525
         TabIndex        =   17
         Top             =   1965
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         Height          =   195
         Left            =   540
         TabIndex        =   16
         Top             =   1545
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
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Simpan"
      Height          =   345
      Left            =   720
      TabIndex        =   2
      Top             =   5040
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   5040
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "&Tutup"
      Height          =   405
      Left            =   6720
      TabIndex        =   0
      Top             =   5040
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4080
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      BackColor       =   &H0000FF00&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   6600
      Top             =   4935
      Width           =   1305
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   420
      Left            =   840
      Top             =   4935
      Width           =   1305
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   2520
      Top             =   4920
      Width           =   1215
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
      Left            =   7020
      TabIndex        =   23
      Top             =   3720
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql, sql1, sql2, yuyu As String
Dim rsbarang As New ADODB.Recordset
Dim Rec As New ADODB.Recordset
Dim hut, hak, haw, TEXT999 As Double
Private Sub Command1_Click()
Dim saldo, saldoAkhir, hrgbeli, hrgbelibaru
If Text7 > 0 Then
    'sql = "update tbreturnjual set jumlah= " & Val(Text7) & ", harga=" & Val(text9) & ",dasar=" & Val(text10) & " where kode='" & Text5.Text & "' and noreturnjual='" & Text1 & "' and nomor='" & Text2 & "'"
   ' con.Execute (sql)
    'sql1 = "update stock set masuk=" & Text7 & ",keluar=0 where tglbukti='" & Format(tgl, "yyyy-MM-dd") & "' and nobukti='" & Trim(Text2.Text) & "' and kode='" & Trim(Text5.Text) & "' and nomor='" & Text2 & "' "
    'con.Execute (sql1)
    'con.Execute ("update tbjual set tanggal_r='" & Format(tgl, "yyyy-MM-dd") & "',jumlah='" & Val(Text7.Text) & "',harga=" & Val(text9) & ",dasar=" & Val(text10) & " where nobukti='" & Trim(Text2) & "' and kode='" & Text5.Text & "' and nomor='" & Text2 & "'")
    
Else
     con.Execute ("delete  from stock where tglbukti='" & Format(tgl, "yyyy-MM-dd") & "' and nobukti='" & Trim(Text2.Text) & "' and kode='" & Trim(Text5.Text) & "'")
     con.Execute ("delete  from tbreturnjual where kode='" & Text5.Text & "' and noreturnjual='" & Text1 & "'")
     sql2 = "update tbjual set tanggal_r='" & Format(tgl, "yyyy-MM-dd") & "',jumlah=0,harga=0,dasar=0 where nobukti='" & Trim(Text2) & "' and kode='" & Text5.Text & "' and nomor='" & Text2 & "'"
     con.Execute (sql2)
End If
kosong
End Sub
Private Sub Command2_Click()
kosong
End Sub
Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Form_Load()
If con.State = adStateClosed Then
connect
End If
tgl = Date
End Sub
Private Sub text1_change()
strsql = " SELECT rb.`noreturnjual`, rb.`nobukti`,rb.`tglbukti`,rb.`KODE`, o.`Nama` as namao, rb.`jumlah`,o.`jumlah_akhir`,rb.`harga`,rb.`dasar`,rb.`nomor` From  tbreturnjual rb, tbbarang o Where   rb.`KODE`=o.`Kode` and noreturnjual='" & Text1 & "'"
Set rsbarang = con.Execute(strsql)
Text1 = rsbarang!noreturnjual
DTPicker1 = rsbarang!tglbukti
Text2 = rsbarang!nobukti
Text3 = rsbarang!nomor
Text5 = rsbarang!kode
Text6 = rsbarang!namao
Text7 = rsbarang!jumlah
text8 = rsbarang!jumlah_akhir
text9 = rsbarang!harga
Text10 = rsbarang!dasar
Text11 = Val(Text7)
End Sub
Private Sub kosong()
Text2 = ""
Text5 = ""
Text6 = ""
Text7 = 0
text8 = 0
text9 = 0
Text10 = 0
Text11 = 0
End Sub


