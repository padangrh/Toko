VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_User 
   Caption         =   "User"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7035
   LinkTopic       =   "Form2"
   ScaleHeight     =   5955
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
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
      Left            =   5520
      TabIndex        =   17
      Top             =   5160
      Width           =   1335
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
      ItemData        =   "Form_User.frx":0000
      Left            =   4080
      List            =   "Form_User.frx":000D
      TabIndex        =   16
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CheckBox Check4 
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
      Left            =   6000
      TabIndex        =   15
      Top             =   4560
      Width           =   360
   End
   Begin VB.CheckBox Check3 
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
      Left            =   6000
      TabIndex        =   14
      Top             =   3840
      Width           =   360
   End
   Begin VB.CheckBox Check2 
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
      Left            =   3840
      TabIndex        =   13
      Top             =   4560
      Width           =   360
   End
   Begin VB.CheckBox Check1 
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
      Left            =   3840
      TabIndex        =   12
      Top             =   3840
      Value           =   1  'Checked
      Width           =   360
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
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txt_id 
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
      Left            =   4080
      TabIndex        =   10
      Top             =   1800
      Width           =   2775
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   915
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   1614
      ButtonWidth     =   3043
      ButtonHeight    =   1455
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   108
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_User.frx":002F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_User.frx":0DC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_User.frx":1DF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_User.frx":2D1F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox list_user 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      TabIndex        =   9
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      TabIndex        =   8
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Laporan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   7
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Penjualan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label4 
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
      Height          =   1215
      Left            =   2400
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      Height          =   1215
      Left            =   2400
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "User Manager"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
   End
End
Attribute VB_Name = "Form_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser As ADODB.Recordset

Private Sub cb_status_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub Command1_Click()
    reset
End Sub

Private Sub Form_Load()
    reload
End Sub

Private Sub reload()
    list_user.Clear
    Set rsUser = con.Execute("select * from tblogin")
    If rsUser.EOF Then Exit Sub
    
    rsUser.MoveFirst
    Do While Not rsUser.EOF
        list_user.AddItem (rsUser!userid)
        rsUser.MoveNext
    Loop
End Sub

Private Function getUser(userid As String) As Boolean
    Dim found As Boolean
    found = False
    rsUser.MoveFirst
    Do While Not rsUser.EOF
        If rsUser!userid = userid Then
            found = True
            Exit Do
        End If
        rsUser.MoveNext
    Loop
    
    getUser = found
End Function

Private Sub list_user_Click()
    If getUser(list_user.Text) Then
        txt_id = rsUser!userid
        txt_password = rsUser!pass
        cb_status.Text = rsUser!posisi
        Check1.Value = rsUser!hak1
        Check2.Value = rsUser!hak2
        Check3.Value = rsUser!hak3
        Check4.Value = rsUser!hak4
    Else
        reset
    End If
End Sub

Private Sub reset()
    txt_id = ""
    txt_password = ""
    cb_status.ListIndex = 0
    Check1.Value = 1
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1: tambah
        Case 2: ubah
        Case 3: hapus
        Case 4: keluar
    End Select
End Sub

Private Sub tambah()
    If txt_id = "" Or txt_password = "" Then
        MsgBox "Data tidak lengkap"
        Exit Sub
    End If
    
    If getUser(txt_id) Then
        MsgBox "User ID telah terpakai"
        Exit Sub
    End If
    
    con.Execute ("insert into tblogin values('" & txt_id & "', '" & txt_password & "', '" & cb_status.Text & "', " & Check1.Value & ", " & Check2.Value & ", " & Check3.Value & ", " & Check4.Value & ",'')")
    reload
    reset
End Sub

Private Sub ubah()
    If txt_id = "" Or txt_password = "" Then
        MsgBox "Data tidak lengkap"
        Exit Sub
    End If
    
    If Not getUser(txt_id) Then
        MsgBox "User ID tidak ditemukan"
        Exit Sub
    End If
    
    con.Execute ("update tblogin set pass = '" & txt_password & "', posisi = '" & cb_status.Text & "', hak1 = " & Check1.Value & ", hak2 = " & Check2.Value & ", hak3 = " & Check3.Value & ", hak4 = " & Check4.Value & " where userid = '" & txt_id & "'")
    reload
    reset
End Sub

Private Sub hapus()
    con.Execute ("delete from tblogin where userid = '" & txt_id & "'")
    reload
    reset
End Sub

Private Sub keluar()
    Unload Me
End Sub

Private Sub txt_id_KeyPress(KeyAscii As Integer)
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
