VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H000080FF&
   Caption         =   "Cetak Bill"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6180
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5610
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "NON-TUNAI"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   44367875
      CurrentDate     =   42224
   End
   Begin VB.TextBox Text1 
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
      Height          =   615
      Left            =   2640
      TabIndex        =   9
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak Bill"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   4800
      Width           =   1575
   End
   Begin MSMask.MaskEdBox mb12 
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   3840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text30 
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mb10 
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   0
      BackColor       =   65280
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "TANGGAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "ISI NO. FAKTUR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "KEMBALIAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "JUMLAH UANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "TOTAL BELANJA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Val(text30) > 0 Then
    Check1.Value = 0
    Exit Sub
End If
sql = "Update tbjual set cash=0 where nobukti='" & Text1 & "'"
Set rsObat = con.Execute(sql)
Command1_Click
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If con.State = adStateClosed Then
connect
End If
DTPicker1 = Date
End Sub
Private Sub Form_Activate()
Text1.SetFocus
End Sub
Private Sub text1_change()
    sql1 = "select * from bill where nofaktur='" & Trim(Text1) & "'"
    Set rsbill = con.Execute(sql1)
    If Not rsbill.EOF Then
    
        Do While Not rsbill.EOF
        Dim total_bayar As Double
        total_bayar = total_bayar + rsbill!bayar
        rsbill.MoveNext
        Loop
        
'        Dim total_bayar As Double
'        total_bayar = total_bayar + rsbill!bayar
         mb10.Text = total_bayar
    Else
        mb10 = 0
      
End If
End Sub

Private Sub Command1_Click()
   Dim total As Currency
   Dim tot_bel As Integer
   Dim rscetak As New Recordset
   rscetak.Open "select * from bill where nofaktur='" & Trim(Text1.Text) & "' ", con, adOpenStatic, adLockOptimistic
         Printer.Font = "times new roman"
        Printer.CurrentX = 0
        Printer.CurrentY = 0
        Printer.FontSize = 18
        Printer.FontBold = True
        Printer.Print Tab(3); " KRIPIK BALADO";
        Printer.Print Tab(2); "CHRISTINE HAKIM";
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.Print Tab(3); "                                                            "; 'baris kosong
        Printer.Print Tab(3); "Jl. Adinegoro No. 11A Padang";
        Printer.Print Tab(3); "                                                             ";
        Printer.Print Tab(3); "No. FAKTUR :"; rscetak.Fields("nofaktur")
        Printer.Print Tab(3); Format(rscetak.Fields("tanggal"), "dd-MM-yyyy"); "  "; Format(rscetak.Fields("jam"), "hh:mm:ss AM/PM");
        'Printer.Print Tab(1); "Jam        : "; Form12.Label7;
        Printer.Print Tab(3); "                                                             ";
        Do While Not rscetak.EOF
             no = no + 1
                Printer.Print Tab(3); rscetak.Fields("nama")
                Printer.Print Tab(3); rscetak.Fields("jumlah_beli"); "x"; Format(rscetak.Fields("harga_jual"), "###,###,###"); Tab(35); Format(rscetak.Fields("bayar"), "###,###,###")
                'Printer.Print Tab(1); ""; Format(rscetak.Fields("bayar"), "###,###,###")
        total = total + rscetak("bayar")
        tot_bel = tot_bel + rscetak("jumlah_beli")
        rscetak.MoveNext
        Loop
        Printer.Print Tab(30); "--------------------------";
        Printer.FontSize = 14
        Dim txt As String
        txt = "Total: " + Format(total, "###,###,###")
        Printer.Print Tab(12); txt
        Printer.CurrentX = 0
        Printer.FontSize = 10
        Printer.Print Tab(3); "                                                             ";
        If (Check1.Value = 0) Then
        Printer.Print Tab(3); "Jumlah Uang  "; Tab(25); Format(text30, "###,###,###");
        Printer.Print Tab(3); "Kembalian    "; Tab(25); Format(mb12, "###,###,###");
        Else
        Printer.Print Tab(3); "-NON TUNAI-";
        End If
        Printer.Print Tab(3); "                                                             ";
        Printer.FontSize = 10
        Printer.Print Tab(3); "Total Items   :"; " "; Format(tot_bel, "###,###,###")
        Printer.Print Tab(3); "Nama Kasir    :"; " "; Form12.Text16.Text;
        Printer.Print Tab(3); "                                                             ";
        Printer.Print Tab(3); "Customer Service: (0751)483518";
        Printer.Print Tab(3); "HP Pemesanan: 0811 668 5000";
        Printer.Print Tab(3); "Website: www.christinehakimideapark.com";
        Printer.Print Tab(3); "                                                             ";
        Printer.FontSize = 8
        Printer.Print Tab(3); "*Barang yang sudah dibeli tidak dapat dikembalikan*";
        Printer.Print Tab(25); "*Terima Kasih*";
      
Close #1
Printer.EndDoc

If Trim(Form12.Text1) <> "" Then
               Set rsbantu1 = con.Execute("select * from tbbantu1 where nobukti='" & Trim(Text1) & "'")
               If rsbantu1.EOF Then
                  con.Execute ("insert into tbbantu1 values('" & Trim(Text1) & "','" & Format(DTPicker1, "yyyy-MM-dd") & "'," & Val(mb10) & "," & Val(text30) & "," & Val(mb12) & ")")
               Else
                  sql1 = "Update tbbantu1 set tglbukti='" & Format(DTPicker1, "yyyy-MM-dd") & "',jml_bayar=" & Val(mb10) & ",jml_uang=" & Val(mb10) & ",kembali=" & Val(mb12) & " where nobukti='" & Trim(Text1.Text) & "'"
                  con.Execute (sql1)
               End If
               'kosong
            Else
               MsgBox "Kode Tidak Boleh Kosong", vbYesNo + vbQuestion, "Confirm"
            End If

akhir
End Sub

Private Sub akhir()
Unload Me
End Sub

Private Sub text30_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If text30 = "" Or Val(text30) < (mb10) Then
            MsgBox "Jumlah Pembayaran Kurang (enter)", vbOKOnly + vbInformation, "Cek Jumlah Pembayaran"
            text30.SetFocus
        Else
            mb12 = Val(text30) - Val(mb10)
            sql = "Update tbjual set cash=1 where nobukti='" & Text1 & "'"
            Set rsObat = con.Execute(sql)
            Command1.SetFocus
        End If
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    text30.SetFocus
    End If
End Sub
