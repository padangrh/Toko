VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Entri_Supplier 
   BackColor       =   &H0000C000&
   Caption         =   "Entri dan Update Data Suplier"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8295
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7980
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_nama_rek 
      Appearance      =   0  'Flat
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
      Left            =   2640
      TabIndex        =   6
      Top             =   4680
      Width           =   5175
   End
   Begin VB.TextBox txt_no_rek 
      Appearance      =   0  'Flat
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
      Left            =   2640
      TabIndex        =   7
      Top             =   5400
      Width           =   5175
   End
   Begin VB.TextBox txt_bank 
      Appearance      =   0  'Flat
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
      Left            =   2640
      TabIndex        =   8
      Top             =   6120
      Width           =   5175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   94830595
      CurrentDate     =   42145
   End
   Begin VB.CommandButton btn_cancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   10
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton btn_save 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txt_telp 
      Appearance      =   0  'Flat
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
      Left            =   2640
      TabIndex        =   4
      Top             =   3120
      Width           =   5175
   End
   Begin VB.TextBox txt_alamat 
      Appearance      =   0  'Flat
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
      Left            =   2640
      TabIndex        =   3
      Top             =   2400
      Width           =   5175
   End
   Begin VB.TextBox txt_nama 
      Appearance      =   0  'Flat
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1680
      Width           =   5175
   End
   Begin VB.TextBox txt_kode 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
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
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C000&
      Caption         =   "Nama Rekening"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C000&
      Caption         =   "No. Rekening"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C000&
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C000&
      Caption         =   "Tgl. Bergabung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C000&
      Caption         =   "Telp Suplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "Alamat Suplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "Nama Suplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Kode Suplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTRI DAN  UPDATE DATA SUPLIER"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "Form_Entri_Supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Save_Click()
    If Trim(txt_kode) = "" Then
        MsgBox "Kode Tidak Boleh Kosong"
        Exit Sub
    ElseIf Trim(txt_nama) = "" Then
        MsgBox "Nama Tidak Boleh Kosong"
        Exit Sub
    End If
    
    Set rsSupplier = con.Execute("select * from tbsuplier")
    
    If getSupplier(txt_kode) Then
        con.Execute ("Update tbsuplier set nmsuplier='" & txt_nama & "', alamat='" & txt_alamat & "',telp='" & txt_telp & "',tgl_gabung='" & Format(DTPicker1, "yyyy-MM-dd") & "', nama_rek = '" & txt_nama_rek & "', no_rek = '" & txt_no_rek & "', bank = '" & txt_bank & "' where kdsuplier='" & txt_kode.Text & "'")
    Else
        'editV2
        con.Execute ("Insert into tbsuplier (kdsuplier, nmsuplier, alamat, telp, tgl_gabung, nama_rek, no_rek, bank) values('" & txt_kode & "','" & txt_nama & "','" & txt_alamat & "','" & txt_telp & "','" & Format(DTPicker1, "yyyy-MM-dd") & "', '" & txt_nama_rek & "', '" & txt_no_rek & "', '" & txt_bank & "')")
    End If
    Unload Me
    Form_List_Supplier.refreshlist
End Sub

Sub kosongkan()
    txt_nama = ""
    txt_alamat = ""
    txt_telp = ""
    txt_nama_rek = ""
    txt_no_rek = ""
    txt_bank = ""
End Sub

Private Sub btn_cancel_Click()
    Unload Me
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_nama_rek.SetFocus
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub Form_Activate()
    If txt_kode = "" Then
        txt_kode.SetFocus
    Else
        txt_nama.SetFocus
    End If
End Sub
Private Sub Form_Load()
    kosongkan
    DTPicker1 = Date
End Sub
Private Sub txt_kode_change()
    Set rsSupplier = con.Execute("select * from tbsuplier where kdsuplier='" & Trim(txt_kode) & "'")
    If Not rsSupplier.EOF Then
        txt_nama = rsSupplier!nmsuplier
        txt_alamat.Text = rsSupplier!alamat
        txt_telp.Text = rsSupplier!telp
        DTPicker1 = rsSupplier!tgl_gabung
        txt_nama_rek = rsSupplier!nama_rek
        txt_no_rek = rsSupplier!no_rek
        txt_bank = rsSupplier!bank
    Else
        kosongkan
    End If
End Sub

Private Sub txt_kode_KeyPress(key As Integer)
    If key = 13 Then
        txt_nama.SetFocus
    End If
End Sub
Private Sub txt_nama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_alamat.SetFocus
        If txt_nama_rek = "" Then
            txt_nama_rek = txt_nama
        End If
    End If
    KeyAscii = validateKey(KeyAscii, 3)
End Sub

Private Sub txt_alamat_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_telp.SetFocus
    KeyAscii = validateKey(KeyAscii, 3)
End Sub

Private Sub txt_telp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DTPicker1.SetFocus
    KeyAscii = validateKey(KeyAscii, 3)
End Sub

Private Sub txt_nama_rek_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_no_rek.SetFocus
    KeyAscii = validateKey(KeyAscii, 3)
End Sub

Private Sub txt_no_rek_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_bank.SetFocus
    KeyAscii = validateKey(KeyAscii, 3)
End Sub

Private Sub txt_bank_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then btn_save.SetFocus
    KeyAscii = validateKey(KeyAscii, 3)
End Sub


