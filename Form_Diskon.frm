VERSION 5.00
Begin VB.Form Form_Diskon 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Diskon"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8895
   FillColor       =   &H0000FF00&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2310
      Left            =   5880
      ScaleHeight     =   2250
      ScaleWidth      =   1785
      TabIndex        =   12
      Top             =   0
      Width           =   1845
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4920
      Top             =   840
   End
   Begin VB.CommandButton btn_cancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton btn_ok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txt_diskon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   9
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txt_customer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   7
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox cb_status 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form_Diskon.frx":0000
      Left            =   2040
      List            =   "Form_Diskon.frx":0013
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txt_password 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txt_spv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lbl_Result 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4920
      TabIndex        =   14
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label lbl_Finger 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Diskon"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Customer"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Status"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Supervisor"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form_Diskon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents myDevices4 As FPDevices
Attribute myDevices4.VB_VarHelpID = -1
Dim WithEvents dev4 As FPDevice
Attribute dev4.VB_VarHelpID = -1
Dim verTemplate As FPTemplate
Dim regTemplate As FPTemplate
Dim Rec As New ADODB.Recordset
Dim c As Integer
Dim hhkLowLevelKybd As Long
Dim Mulai As Boolean

Private Sub btn_cancel_Click()
    Unload Me
End Sub

Private Sub btn_ok_Click()

    If cek_Status = False Then
        MsgBox "Status tidak valid"
        Exit Sub
    End If
    
    If txt_diskon.Text = "" Then txt_diskon.Text = 0

    If priceToNum(Form_Print.txt_total) < priceToNum(txt_diskon.Text) Then
        MsgBox "Diskon yg diberikan lebih besar dari harga barang"
    Else
        Dim rsUser As ADODB.Recordset
        Set rsUser = con.Execute("select * from tblogin where userid = '" & txt_spv & "'")
        If rsUser.EOF Or rsUser.BOF Then
            MsgBox "Supervisor tidak terdaftar"
            Exit Sub
        End If
        
        If rsUser!posisi = "Karyawan" Then
            MsgBox "Hanya supervisor yang bisa memberi diskon"
            Exit Sub
        End If
        
        If rsUser!pass <> txt_password Then
            MsgBox "Password salah"
            Exit Sub
        End If
        
        Form_Print.txt_diskon = Format(txt_diskon, "###,###,##0")
        
        Form_Print.diskon_query
        
    '    con.Execute ("insert into tbdiskon values('" & Form_Print.txt_bon & "', '" & txt_spv & "', '" & cb_status.Text & "', '" & txt_customer & "', " & priceToNum(txt_diskon) & ")")
    '
    '    Printer.Font = "Times new roman"
    '    Printer.FontSize = 12
    '    Printer.Print Tab(4); Format(Now, "dd-MM-yyyy  hh:mm:ss");
    '    Printer.Print Tab(4); "No Faktur"; Tab(18); ": "; Form_Print.txt_bon
    '    Printer.Print Tab(4); "Supervisor"; Tab(18); ": "; txt_spv
    '    Printer.Print Tab(4); "Status"; Tab(18); ": "; cb_status.Text
    '    Printer.Print Tab(4); "Customer"; Tab(18); ": "; txt_customer
    '    Printer.Print Tab(4); "Diskon"; Tab(18); ": Rp."; txt_diskon
    '    Printer.EndDoc
        'If priceToNum(Form_Print.txt_uang.Text) > 0 Then
        Form_Print.txt_kembali = Format(priceToNum(Form_Print.txt_uang) - (Val(Form_Print.txt_total) - Val(Form_Print.txt_diskon)), "###,###,##0")
        rsUser.Close
        Unload Me
        
        Form_Print.txt_uang.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Mulai = False
    Dim X As Variant
    lbl_Finger.Caption = "FingerPrint Sedang DiAktifkan ...."
'    On Error GoTo Keluar
    
    Set myDevices4 = New FPDevices
    If myDevices4.count <> 0 Then
        For Each X In myDevices4
            Set dev4 = X
            dev4.SubScribe Dp_StdPriority, Me.hWnd
        Next
        
        lbl_Finger.Caption = "Letakan Jari Anda pada FingerPrint"
    Else
        lbl_Finger.Caption = "FingerPrint Belum Terpasang !!!"
    End If
    Set X = Nothing
    Mulai = True
    Exit Sub
'Keluar:
'    MsgBox Err.Description, vbInformation + vbSystemModal, "Informasi"
End Sub

Private Sub Form_unload(cancel As Integer)
    Form_Print.Enabled = True
    If Not (dev4 Is Nothing) Then
        dev4.UnSubScribe
    End If
    Set dev4 = Nothing
    Set myDevices4 = Nothing
    Set Rec2 = Nothing
    Set Rec = Nothing
    DoEvents
End Sub

Private Sub txt_diskon_LostFocus()
    txt_diskon = Format(txt_diskon, "###,###,##0")
End Sub

Private Sub txt_spv_keypress(key As Integer)
    If key = 13 Then
        txt_password.SetFocus
    End If
End Sub

Private Sub txt_password_keypress(key As Integer)
    If key = 13 Then
        cb_status.SetFocus
    End If
End Sub

Private Sub cb_status_keypress(key As Integer)
    If key = 13 Then
        txt_customer.SetFocus
    End If
End Sub

Private Sub txt_customer_keypress(key As Integer)
    If key = 13 Then
        txt_diskon.SetFocus
    End If
End Sub

Private Sub txt_diskon_keypress(key As Integer)
'    If key = 13 Then
'        btn_ok.SetFocus
'    End If
    Select Case key
        Case 48 To 57, 8 '  0-9 & backspace
        'Let these key codes pass through
        Case 13
            btn_ok.SetFocus
        Case Else
        'All others get trapped
        key = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Function cek_Status() As Boolean
    cek_Status = False
    Dim i As Integer
    Do While i < cb_status.ListCount
        If Trim(UCase(cb_status.Text)) = Trim(UCase(cb_status.List(i))) Then
            cek_Status = True
        End If
        i = i + 1
    Loop
End Function

Private Sub dev4_FingerLeaving()
    lbl_Finger.Caption = "Letakan Jari Anda pada FingerPrint"
End Sub

Private Sub dev4_FingerTouching()
    lbl_Finger.Caption = "Sidik Jari di-Process"
End Sub

Private Sub dev4_SampleAcquired(ByVal pRawSample As Object)
    If Mulai = False Then Exit Sub
    
    On Error Resume Next
    
    Dim Sample As FPSample
    Dim smpPro As FPRawSamplePro
    
    Set smpPro = New FPRawSamplePro
    smpPro.Convert pRawSample, Sample
    
    Sample.PictureOrientation = Or_Portrait
    Sample.PictureWidth = Picture1.Width / Screen.TwipsPerPixelX
    Sample.PictureHeight = Picture1.Height / Screen.TwipsPerPixelY
    Picture1.Picture = Sample.Picture
    DoEvents
    
    Dim ftrex As FPFtrEx
    Dim qt As AISampleQuality
    
    Set ftrex = New FPFtrEx
    ftrex.Process Sample, Tt_Verification, verTemplate, qt
    
    lbl_Finger.Caption = "Proses Selesai !!!"
    If qt = Sq_Good Then
        Cek
    Else
        c = 0
        Timer1.Enabled = True
        lbl_Result.Caption = "Hasil Scan Tidak Bagus, Letakan Jari Anda Dengan Benar"
    End If
    'Picture1.Picture = LoadPicture("")
    Text1.SetFocus
End Sub



Private Sub myDevices4_DeviceConnected(ByVal serNum As String)
    If myDevices4.count <> 0 Then
        Set dev4 = Nothing
'        For Each x In myDevices4
'            Set dev = x
'
'        Next
        dev4.SubScribe Dp_StdPriority, Me.hWnd
        lbl_Finger.Caption = "Letakan Jari Anda pada FingerPrint"
    End If
End Sub

Private Sub myDevices4_DeviceDisconnected(ByVal serNum As String)

    lbl_Finger.Caption = "FingerPrint Belum Terpasang !!!"
    On Error Resume Next
'        For Each x In myDevices4
'            Set dev = x
'
'        Next
        dev4.UnSubScribe
    Set dev4 = Nothing
End Sub

Private Sub Timer1_Timer()
    lbl_Result.Visible = True
    c = c + 1
    If c = 10 Then
        Timer1.Enabled = False
        lbl_Result.Visible = False
    End If

End Sub

Private Sub Cek()
    
    Dim verify As FPVerify
    Dim result As Boolean
    Dim score As Variant
    Dim threshold As Variant
    Dim learn As Boolean
    Dim sec As AISecureModeMask
    Dim blob As String
    Dim blobarray() As Byte
    Set verify = New FPVerify
    Dim Kjk As Byte
'    Dim temp_String As String
    Dim Nama As String
    Dim kode As String
    Dim LogOke As Boolean
    
    lbl_Result.ForeColor = vbBlack
'    temp_String = "Start : " & Format(Now, "h:mm:ss")
    Rec.Open "select * from tblogin", con, adOpenForwardOnly, adLockReadOnly
    Do Until Rec.EOF
        blob = Rec.Fields("fingerprint")
        'MsgBox blob
        hextoarray blob, blobarray
'        blobarray = Base64Decode(blob)
        'stringToarray blob, blobarray
        
        Set regTemplate = New FPTemplate
        regTemplate.Import blobarray
        verify.compare regTemplate, verTemplate, result, score, threshold, learn, sec
        
        If result = True Then
'            Kode = Rec.Fields("kode")
'            Nama = Rec.Fields("nama")
            LogOke = True
            txt_spv.Text = Rec.Fields("userid")
            txt_password.Text = Rec.Fields("pass")
            Exit Do
        End If
        Set regTemplate = Nothing
        Rec.MoveNext
    Loop
    Rec.Close
    
    If LogOke = True Then
        lbl_Result.ForeColor = vbBlue
        lbl_Result.Caption = Nama & " Berhasil Login !!!"
    Else
        lbl_Result.Caption = "Login Gagal, Ulangi !!!"
        lbl_Result.ForeColor = vbRed
    End If
    c = 0
'    MsgBox temp_String & vbNewLine & "End   : " & Format(Now, "h:mm:ss")
    Timer1.Enabled = True
'    Lbl_Kode.Caption = ""
'    Label7.Caption = ""
End Sub
