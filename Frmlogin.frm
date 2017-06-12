VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H0000C000&
   Caption         =   "Security"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "Frmlogin.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3270
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5400
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2310
      Left            =   6480
      ScaleHeight     =   2250
      ScaleWidth      =   1785
      TabIndex        =   7
      Top             =   120
      Width           =   1845
   End
   Begin VB.CommandButton Commandbatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Commandlogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtpass 
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
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtuser 
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
      Height          =   525
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lbl_Finger 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   2520
      Width           =   3615
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
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "Password"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "User Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Login"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents myDevices As FPDevices
Attribute myDevices.VB_VarHelpID = -1
Dim WithEvents dev3 As FPDevice
Attribute dev3.VB_VarHelpID = -1
Dim verTemplate As FPTemplate
Dim regTemplate As FPTemplate
Dim Rec As New ADODB.Recordset
Dim c As Integer
Dim hhkLowLevelKybd As Long
Dim Mulai As Boolean
Dim cmdbatal As Boolean

Private Sub CommandBatal_Click()
    cmdbatal = True
'    con.Close
    Unload Me
End Sub

Private Sub CommandLogin_Click()
'On Error GoTo salah:
    Dim Rec As ADODB.Recordset
    Set Rec = con.Execute("select * from tblogin where UserID='" & Trim(txtuser.Text) & "'")
    If Not Rec.EOF Then
        If UCase(Rec.Fields("UserID")) = UCase(Trim(txtuser)) And Rec.Fields("pass") = Trim(txtpass) Then
            username = Rec!userid
            status = Rec!posisi
            If status = "Master" Then
                FrmMain.tbhs.Visible = True
            End If
            FrmMain.p.Enabled = CBool(Rec.Fields("hak1"))
            FrmMain.Toolbar1.Buttons(1).Enabled = CBool(Rec.Fields("hak1"))
            FrmMain.l.Enabled = CBool(Rec.Fields("hak2"))
            FrmMain.Toolbar1.Buttons(2).Enabled = CBool(Rec.Fields("hak2"))
            FrmMain.b.Enabled = CBool(Rec.Fields("hak3"))
            FrmMain.Toolbar1.Buttons(3).Enabled = CBool(Rec.Fields("hak3"))
            FrmMain.a.Enabled = CBool(Rec.Fields("hak4"))
            FrmMain.Toolbar1.Buttons(4).Enabled = CBool(Rec.Fields("hak4"))
            'Call Form_unload(0)
            
            Unload Me
            DoEvents
            FrmMain.Show

            FrmMain.Toolbar1.Enabled = True
            Exit Sub
        Else
            lbl_Result.ForeColor = vbRed
            lbl_Result.Caption = "Login Gagal, Ulangi !!!"
            Timer1.Enabled = False
            Timer1.Enabled = True
'            MsgBox "Nama user atau password anda tidak cocok!"
            txtuser.SetFocus
        End If
    Else
'      MsgBox "Nama user atau password anda tidak cocok!"
        lbl_Result.ForeColor = vbRed
        lbl_Result.Caption = "Login Gagal, Ulangi !!!"
        Timer1.Enabled = False
        Timer1.Enabled = True
        txtuser.SetFocus
    End If

    
    
'salah:
'MsgBox "Periksa komputer server hidup atau tidak, kabel internet tercolok di komputer atau tidak, coba restart modem"
End Sub

Private Sub Form_Activate()
    txtuser.SetFocus
End Sub

Private Sub Form_Load()
    
    Mulai = False
    cmdbatal = False
    lbl_Finger.Caption = "FingerPrint Sedang DiAktifkan ...."
'    On Error GoTo Keluar
    Dim x As Variant
    Set myDevices = New FPDevices
    If myDevices.count <> 0 Then
        For Each x In myDevices
            Set dev3 = x
            dev3.SubScribe Dp_StdPriority, Me.hWnd
        Next
        
        lbl_Finger.Caption = "Letakan Jari Anda pada FingerPrint"
    Else
        lbl_Finger.Caption = "FingerPrint Belum Terpasang !!!"
    End If
    Set x = Nothing
    Mulai = True
    DoEvents
    Exit Sub
'Keluar:
'    MsgBox Err.Description, vbInformation + vbSystemModal, "Informasi"
End Sub

Private Sub txtpass_keyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        CommandLogin_Click
    End If
End Sub

Private Sub Form_unload(cancel As Integer)
'con.Close
    Mulai = False
    If Not (dev3 Is Nothing) Then
        dev3.UnSubScribe
    End If
'    Set dev3 = Nothing
'    Set myDevices = Nothing
'    Set Rec = Nothing
    If cmdbatal = True Then End
    DoEvents
'    Dim ctrl As Control
'    For Each ctrl In Me.Controls
'        If TypeOf ctrl Is CommandButton Then
'            ctrl.Enabled = False
'        End If
'    Next
'    Set frmlogin = Nothing
    

End Sub

Private Sub dev3_FingerLeaving()
    lbl_Finger.Caption = "Letakan Jari Anda pada FingerPrint"
End Sub

Private Sub dev3_FingerTouching()
    lbl_Finger.Caption = "Sidik Jari di-Process"
End Sub

Private Sub dev3_SampleAcquired(ByVal pRawSample As Object)
    If Mulai = False Then Exit Sub
    If dev3 Is Nothing Then Exit Sub
'    On Error Resume Next
    
'    MsgBox Me.Name
    
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
'    Text1.SetFocus
End Sub



Private Sub myDevices_DeviceConnected(ByVal serNum As String)
    If myDevices.count <> 0 Then
        Set dev3 = Nothing
'        For Each x In myDevices
'            Set dev = x
'
'        Next
'        dev3.SubScribe Dp_StdPriority, Me.hWnd
        lbl_Finger.Caption = "Letakan Jari Anda pada FingerPrint"
    End If
End Sub

Private Sub myDevices_DeviceDisconnected(ByVal serNum As String)

    lbl_Finger.Caption = "FingerPrint Belum Terpasang !!!"
    On Error Resume Next
'        For Each x In myDevices
'            Set dev = x
'
'        Next
    dev3.UnSubScribe
    Set dev3 = Nothing
End Sub

Private Sub Timer1_Timer()
    lbl_Result.Visible = False
    Timer1.Enabled = False

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
    Dim temp_String As String
    Dim Nama As String
    Dim kode As String
    Dim LogOke As Boolean
    
    lbl_Result.ForeColor = vbBlack
    temp_String = "Start : " & Format(Now, "h:mm:ss")
    Rec.Open "select * from tblogin where fingerprint <> ''", con, adOpenForwardOnly, adLockReadOnly
    Do Until Rec.EOF
        blob = Rec.Fields("fingerprint")
        'MsgBox blob
'        ReDim blobarray(0 To Len(blob) / 2) As Byte
        hextoarray blob, blobarray
'        blobarray = Base64Decode(blob)
        'stringToarray blob, blobarray

        Set regTemplate = New FPTemplate
        regTemplate.Import blobarray

'        verify.compare regTemplate, verTemplate, result, score, threshold, learn, sec
        verify.compare regTemplate, verTemplate, result, score, threshold, learn, sec
        If result = True Then
'            Kode = Rec.Fields("kode")
'            Nama = Rec.Fields("nama")
            txtuser.Text = Rec.Fields("userid")
            txtpass.Text = Rec.Fields("pass")

            LogOke = True
            Exit Do
        End If
        Set regTemplate = Nothing
        Rec.MoveNext
        Erase blobarray
    Loop
    Set verTemplate = Nothing
    Set regTemplate = Nothing
    Set score = Nothing
'    Set result = Nothing
    Set threshold = Nothing
'    Set learn = Nothing
    Rec.Close
    
    If LogOke = True Then
        lbl_Result.ForeColor = vbBlue
        lbl_Result.Caption = Nama & " Berhasil Login !!!"
        Call CommandLogin_Click
    Else
        lbl_Result.ForeColor = vbRed
        lbl_Result.Caption = "Login Gagal, Ulangi !!!"
        Timer1.Enabled = False
        Timer1.Enabled = True
    End If
'    c = 0
'    MsgBox temp_String & vbNewLine & "End   : " & Format(Now, "h:mm:ss")

'    Lbl_Kode.Caption = ""
'    Label7.Caption = ""
    
End Sub


