VERSION 5.00
Begin VB.Form Form23 
   BackColor       =   &H000080FF&
   Caption         =   "Form23"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6420
   LinkTopic       =   "Form23"
   ScaleHeight     =   1845
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Batal"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   1320
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
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Password:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
   txtpass.SetFocus
End Sub

Private Sub txtpass_keyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Dim rsLogin As ADODB.Recordset
        Set rsLogin = con.Execute("select * from tblogin where userid = 'admin'")
        If (rsLogin.EOF And rsLogin.BOF) Then
            MsgBox ("Ada yang salah, hubungi Richard sekarang!!!!")
        ElseIf (rsLogin.Fields("pass") = txtpass.Text) Then
            Form_List_Supplier2.admin_approval.Value = 1
            MsgBox ("Transaksi ini telah disetujui oleh admin")
        Else
            MsgBox ("Password Salah!")
        End If
        
        Unload Me
    End If
End Sub
