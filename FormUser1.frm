VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormUser1 
   BackColor       =   &H0000C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "FormUser1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7860
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   5895
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   3180
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   4560
         Width           =   5415
      End
      Begin VB.TextBox txt_pass 
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
         ForeColor       =   &H00C00000&
         Height          =   555
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txt_id 
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Status (Admin atau Karyawan Biasa)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "Posisi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
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
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
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
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   7860
      ItemData        =   "FormUser1.frx":628A
      Left            =   120
      List            =   "FormUser1.frx":628C
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUser1.frx":628E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUser1.frx":6790
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUser1.frx":6D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUser1.frx":70E1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb2 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1058
      ButtonWidth     =   2196
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Baru"
            Key             =   "baru"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Simpan"
            Key             =   "simpan"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hapus"
            Key             =   "hapus"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tutup"
            Key             =   "keluar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   390
         Left            =   5400
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FormUser1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub keluar()
  Unload Me
End Sub

Private Sub simpan()
  Dim i
  If txt_id.Text = "" Then
    txt_id.SetFocus
    Exit Sub
  End If
  Adodc1.Refresh
  If Adodc1.Recordset.RecordCount <> 0 Then Adodc1.Recordset.MoveFirst
  Adodc1.Recordset.Find "UserID='" & txt_id.Text & "'"
  If Not Adodc1.Recordset.EOF Then
    Adodc1.Recordset.Update
  Else
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields(0) = txt_id.Text
  End If
    Adodc1.Recordset.Fields(1) = txt_pass.Text
    Adodc1.Recordset.Fields(2) = Text3.Text
    Adodc1.Recordset.Fields(7) = Text6.Text
    Adodc1.Recordset.Fields(8) = Text4.Text
    Adodc1.Recordset.Fields(9) = DTPicker1
    Adodc1.Recordset.Fields(10) = Text5.Text
    Adodc1.Recordset.Fields(11) = mb1
    For i = 3 To 6
      Adodc1.Recordset.Fields(i) = IIf(List2.Selected(i - 3) = True, "1", "0")
    Next
    Adodc1.Recordset.Update
    Call isiListUser
End Sub

Private Sub hapus()
  Dim msg As Byte
  If List1.Text = "Admin" Then
    msg = MsgBox("Admin tidak bisa di hapus !", vbOKOnly + vbInformation, "Hapus User")
    Exit Sub
  End If
  msg = MsgBox("Benar User Akan dihapus ?", vbYesNo + vbQuestion, "Hapus User")
  If msg = vbYes Then
    Adodc1.Recordset.Delete
    Adodc1.Refresh
    Call isiListUser
  End If
End Sub

Private Sub baru()
  txt_id.Text = ""
  txt_pass.Text = ""
  Text3.Text = ""
  Text6 = ""
  mb1 = 0
  Text4 = ""
  Text5 = ""
  DTPicker1 = Date
  txt_id.SetFocus
  For i = 3 To 6
    List2.Selected(i - 3) = False
  Next
End Sub

Private Sub Form_Activate()
  Call isiListUser
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
'   SendKeys ("{tab}")
ElseIf KeyCode = vbKeyEscape Then
   Unload Me
Else
   'KeyCode = 0
End If
End Sub

Private Sub Form_Load()
  txt_id.Text = ""
  txt_pass.Text = ""
  Text3.Text = ""
  Text6 = ""
  mb1 = 0
  Text4 = ""
  Text5 = ""
  DTPicker1 = Date
  Adodc1.ConnectionString = "DSN=data"
  Adodc1.RecordSource = "select * from tblogin"
  Adodc1.Refresh
  If con.State = adStateClosed Then
    connect
  End If
  List2.Clear
  List2.AddItem "Penjualan"   '1
  List2.AddItem "Laporan"  '2
  List2.AddItem "Barang"   '3
  List2.AddItem "Admin" '4
 
End Sub

Private Sub isiListUser()
  List1.Clear
  If Not Adodc1.Recordset.EOF Then
  Adodc1.Recordset.MoveFirst
  Do While Not Adodc1.Recordset.EOF
    List1.AddItem Adodc1.Recordset.Fields(0)
    Adodc1.Recordset.MoveNext
  Loop
  List1.ListIndex = 0
  End If
End Sub

Private Sub Form_unload(cancel As Integer)
Unload Me
End Sub

Private Sub List1_click()
 Dim i
  Adodc1.Refresh
  Adodc1.Recordset.Find "UserID='" & List1.Text & "'"
  txt_id.Text = Adodc1.Recordset.Fields(0)
  txt_pass.Text = Adodc1.Recordset.Fields(1)
  Text3.Text = Adodc1.Recordset.Fields(2)
  Text6 = Adodc1.Recordset.Fields(7)
  Text4.Text = Adodc1.Recordset.Fields(8)
  DTPicker1 = Adodc1.Recordset.Fields(9)
  Text5.Text = Adodc1.Recordset.Fields(10)
  mb1 = Adodc1.Recordset.Fields(11)
  For i = 3 To 6
    List2.Selected(i - 3) = Adodc1.Recordset.Fields(i).Value
  Next
  List2.ListIndex = 0
End Sub

Private Sub Tlb2_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
  Case "baru"
    baru
  Case "simpan"
    simpan
  Case "hapus"
    hapus
  Case "keluar"
    keluar
  End Select
End Sub
