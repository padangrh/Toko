VERSION 5.00
Begin VB.Form Form_RegisterFP 
   Caption         =   "User"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_ConfirmPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   2565
   End
   Begin VB.TextBox txt_NewPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2565
   End
   Begin VB.CommandButton cmd_Batal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txt_Username 
      BackColor       =   &H8000000E&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      MaxLength       =   12
      TabIndex        =   0
      Top             =   240
      Width           =   2565
   End
   Begin VB.TextBox txt_Password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2565
   End
   Begin VB.CommandButton cmd_Simpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scan Sidik Jari Anda"
      Height          =   2535
      Left            =   840
      TabIndex        =   10
      Top             =   2415
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text10 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1665
         ScaleWidth      =   1425
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Sta 
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Tahap 4"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Tahap 3"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Tahap 2"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Tahap 1"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Konfir 
         Height          =   615
         Left            =   1920
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Ketik sekali lagi"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1335
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Password baru"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   975
      Width           =   2175
   End
   Begin VB.Label lbl_Fingerprint 
      Alignment       =   2  'Center
      Caption         =   "Sidik jari telah tersimpan"
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   255
      Width           =   2175
   End
   Begin VB.Label Label16 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   615
      Width           =   2175
   End
End
Attribute VB_Name = "Form_RegisterFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser As ADODB.Recordset
'Dim Blank_SJ As String
'Dim Rec As ADODB.Recordset
'Dim WithEvents dev2 As FPDevice
'Dim Template(4) As FPTemplate
'Dim verTemplate As FPTemplate
'Dim Narray As Integer
'Dim ScanOke As Boolean
'Dim WithEvents myDevices2 As FPDevices

Private Sub cmd_Batal_Click()
    Unload Me
End Sub

Private Sub cmd_Simpan_Click()
    Dim flagPassword As Boolean
    flagPassword = False
    
    If txt_password.Text <> rsUser!pass Then
        MsgBox "Password salah"
        Exit Sub
    End If
    
    If txt_NewPassword <> "" Then
        If txt_NewPassword = txt_ConfirmPassword Then
            flagPassword = True
        Else
            MsgBox "Password baru dan konfirmasi tidak sama"
            txt_NewPassword = ""
            txt_ConfirmPassword = ""
            Exit Sub
        End If
    End If
    
'    If ScanOke = False And flagPassword = False Then
'        MsgBox "Lengkapi 4 Tahap Scan Sidik Jari !!!", vbInformation, "Informasi"
'        Exit Sub
'    ElseIf ScanOke = True Then
'        Dim register As FPRegister
'        Set register = New FPRegister
'        register.NewRegistration Rt_Verify
'
'        Dim bDone As Boolean
'
'        For a = 0 To 3
'            register.Add Template(a), bDone
'        Next a
'
'        If bDone = False Then
'            MsgBox "Sidik Jari pada Tiap Tahap Tidak Sama !!!" & Chr(13) & "Scan Ulang Dari Awal ....", vbInformation, "Informasi"
'            Siap2x
'            Exit Sub
'        End If
'
'        Dim regTemplate As FPTemplate
'        Set regTemplate = register.RegistrationTemplate
'
'        Dim blob As Variant
'        Dim blobarray() As Byte
'
'        regTemplate.Export blob
'        blobarray = blob
'        Dim temp_String As String
'        'temp_String = arrayTostring(blobarray)
'        'temp_String = Base64Encode(blobarray)
'        temp_String = arraytohex(blobarray)
'
'    End If
'
'
'    If flagPassword = True Then
'        con.Execute ("update tblogin set pass = '" & txt_NewPassword.Text & "' where userid = '" & username & "'")
'    End If
'
'    con.Execute ("update tblogin set fingerprint = '" & temp_String & "' where userid = '" & username & "'")
'
'    MsgBox "Data Karyawan Sudah Disimpan !!!", vbInformation, "Information"
'    Siap2x
'    Unload Me
    
    Exit Sub

www:
    If Err.Number = -2147467259 Then
        MsgBox "Nama Karyawan Telah Ada !!!", vbInformation, "Peringatan"
        Text2.SetFocus
    Else
        MsgBox Err.Description, vbInformation, Err.Number
    End If

End Sub

Private Sub Form_Load()
'    frmlogin.Enabled = False
'    Dim X As Variant
'
'    Set myDevices2 = New FPDevices
'    If myDevices2.count <> 0 Then
'        For Each X In myDevices2
'            Set dev2 = X
'            dev2.SubScribe Dp_StdPriority, Me.hWnd
'        Next
'    End If
'
'    Set X = Nothing
'    Sta.Caption = "Letakan Jari Anda Pada FingerPrint"

'    ScanOke = False
    
    txt_Username.Text = username
    Set rsUser = con.Execute("Select * from tblogin where userid = '" & txt_Username.Text & "'")
    If rsUser.EOF Then
        MsgBox "User tidak ditemukan"
        Set rsUser = Nothing
        Unload Me
'    ElseIf rsUser!fingerprint = "" Then
'        lbl_Fingerprint.Visible = False
    End If
End Sub

Private Sub Form_unload(cancel As Integer)
'    frmlogin.Enabled = True
'    Dim X As Variant
'    Set rsUser = Nothing
'    Blank_SJ = ""
'    Set Template(4) = Nothing
'    Set verTemplate = Nothing
'    Narray = 0
'    ScanOke = False
'    If Not (dev2 Is Nothing) Then
'        dev2.UnSubScribe
'    End If
'
'    Set myDevices2 = Nothing
'    Set dev2 = Nothing
'    Set Rec = Nothing
'    DoEvents

End Sub

'Private Sub dev2_FingerLeaving()
'    Sta.Caption = "Letakan Jari Anda Pada FingerPrint"
'End Sub
'
'Private Sub dev2_FingerTouching()
'    Sta.Caption = "Sidik Jari di-Process"
'End Sub
'
'Private Sub dev2_SampleAcquired(ByVal pRawSample As Object)
'
'    If ScanOke = False Then
'        Dim Sample As FPSample
'        Dim smpPro As FPRawSamplePro
'
'        Set smpPro = New FPRawSamplePro
'        smpPro.Convert pRawSample, Sample
'
'        Sample.PictureOrientation = Or_Portrait
'        Sample.PictureWidth = Picture1.Width / Screen.TwipsPerPixelX
'        Sample.PictureHeight = Picture1.Height / Screen.TwipsPerPixelY
'        Picture1.Picture = Sample.Picture
'
'        Dim ftrex As FPFtrEx
'        Dim qt As AISampleQuality
'
'        Set ftrex = New FPFtrEx
'        ftrex.Process Sample, Tt_PreRegistration, Template(Narray), qt
'
'        Text10(Narray).Text = Kualitas(qt)
'        If qt = 0 Then
'            Narray = Narray + 1
'            Konfir.Caption = ""
'            If Narray = 4 Then
'                ftrex.Process Sample, Tt_Verification, verTemplate, qt
'                If qt = 0 Then
'                    ScanOke = True
'                    Konfir.Caption = "Semua Tahap Sudah Lengkap !!!"
'                    Sta.Visible = False
'                Else
'                    Narray = Narray - 1
'                    Text10(Narray).Text = Kualitas(qt)
'                    Konfir.Caption = "Scan Tahap " & (Narray + 1) & " di-ulangi, Letakan Jari Anda dg Benar !!!"
'                End If
'            End If
'        Else
'            Konfir.Caption = "Scan Tahap " & (Narray + 1) & " di-ulangi, Letakan Jari Anda dg Benar !!!"
'        End If
'        Sta.Caption = "Process Selesai"
'    End If
'End Sub
'
'Private Sub Siap2x()
'    Narray = 0
'    ScanOke = False
'    Sta.Visible = True
'    Konfir.Caption = ""
'    Picture1.Picture = LoadPicture("")
'    For a = 0 To 3
'        Text10(a).Text = ""
'    Next a
'End Sub

Private Sub txt_ConfirmPassword_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case 13
            cmd_Simpan.SetFocus
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub txt_NewPassword_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub txt_password_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub
