VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form BeckupOK 
   BorderStyle     =   0  'None
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   1110
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000080&
      Height          =   435
      Left            =   4695
      ScaleHeight     =   375
      ScaleWidth      =   855
      TabIndex        =   11
      Top             =   3105
      Width           =   915
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   15
         MouseIcon       =   "Beckup.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   15
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000080&
      Height          =   435
      Left            =   3750
      ScaleHeight     =   375
      ScaleWidth      =   855
      TabIndex        =   9
      Top             =   3105
      Width           =   915
      Begin VB.CommandButton Command3 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   15
         MouseIcon       =   "Beckup.frx":030A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   15
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   435
      Left            =   2805
      ScaleHeight     =   375
      ScaleWidth      =   855
      TabIndex        =   7
      Top             =   3105
      Width           =   915
      Begin VB.CommandButton Command2 
         Caption         =   "Proses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   15
         MouseIcon       =   "Beckup.frx":0614
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   15
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00000080&
      Height          =   435
      Left            =   4875
      ScaleHeight     =   375
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   2475
      Width           =   525
      Begin VB.CommandButton Command1 
         Height          =   345
         Left            =   15
         MouseIcon       =   "Beckup.frx":091E
         MousePointer    =   99  'Custom
         Picture         =   "Beckup.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   15
         Width           =   435
      End
   End
   Begin VB.PictureBox ComDialog1 
      Height          =   480
      Left            =   1980
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   13
      Top             =   4065
      Width           =   1200
   End
   Begin VB.Image Image5 
      Height          =   3840
      Left            =   30
      Picture         =   "Beckup.frx":0C91
      Stretch         =   -1  'True
      Top             =   45
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   540
      Shape           =   4  'Rounded Rectangle
      Top             =   2085
      Width           =   5055
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   570
      Left            =   1335
      TabIndex        =   4
      Top             =   2400
      Width           =   3465
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TUJUAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   510
      TabIndex        =   3
      Top             =   2115
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   570
      Left            =   1365
      TabIndex        =   2
      Top             =   1365
      Width           =   3465
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00800000&
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   540
      Shape           =   4  'Rounded Rectangle
      Top             =   1050
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "BACKUP DATABASE"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   600
      Left            =   1185
      TabIndex        =   0
      Top             =   435
      Width           =   4470
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOKASI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   510
      TabIndex        =   1
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   165
      Picture         =   "Beckup.frx":0E01
      Stretch         =   -1  'True
      Top             =   975
      Width           =   1260
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   210
      Picture         =   "Beckup.frx":30A4
      Stretch         =   -1  'True
      Top             =   60
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   105
      Index           =   8
      Left            =   1005
      Picture         =   "Beckup.frx":3BB1
      Top             =   3705
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   105
      Index           =   7
      Left            =   165
      Picture         =   "Beckup.frx":3CCD
      Top             =   3705
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   105
      Index           =   6
      Left            =   270
      Picture         =   "Beckup.frx":3DE9
      Top             =   3525
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   105
      Index           =   5
      Left            =   1125
      Picture         =   "Beckup.frx":3F05
      Top             =   3525
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   105
      Index           =   4
      Left            =   5280
      Picture         =   "Beckup.frx":4021
      Top             =   90
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   105
      Index           =   1
      Left            =   4425
      Picture         =   "Beckup.frx":413D
      Top             =   90
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   105
      Index           =   2
      Left            =   4320
      Picture         =   "Beckup.frx":4259
      Top             =   270
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   105
      Index           =   3
      Left            =   5160
      Picture         =   "Beckup.frx":4375
      Top             =   270
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3885
      Left            =   15
      Top             =   15
      Width           =   6135
   End
End
Attribute VB_Name = "BeckupOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As New FileSystemObject
Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 3
Label3.Caption = "C:\ProgramData\MySQL\MySQL Server 6.0\data\data"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{Tab}"
KeyAscii = 0
End If
End Sub
Private Sub Form_Deactivate()
Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyEscape
    Unload Me
  End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set backup = Nothing
End Sub
Private Sub Command1_Click()
With ComDialog
    .InitDir = ""
    .FileName = "data"
    .Filter = "Folder (data)"

    .DialogTitle = "Simpan Dengan Nama"
    .ShowSave
'
End With
Label5 = ComDialog.FileName
End Sub
Private Sub Command2_Click()
If Label5.Caption = "" Then
MsgBox "Pilih dulu tujuannya", vbCritical, "Backup Databae"
Exit Sub
End If
Dim Jawab As Integer
Dim DirAwal, DirAkhir
Jawab = MsgBox("Anda Yakin Akan Melakukan Proses Backup ?", vbYesNo + vbQuestion, "Confirm")
If Jawab = vbYes Then
    DirAwal = Trim(Label3.Caption)
    DirAkhir = Trim(Label5.Caption)
    
    On Error GoTo Perbaikan
    a.CopyFolder DirAwal, DirAkhir
        MsgBox "BackUp Data Sukses...!!!!" _
        , vbInformation, "BackUp Data"
    On Error GoTo 0
    Exit Sub
Perbaikan:
    MsgBox "Ada Kesalahan[" & Err.Description & " ] " & Chr(13) & _
           "Backup Tidak Dilanjutkan", vbOKOnly + vbExclamation, "Error"
Else
Command2.SetFocus
End If
End Sub
Private Sub Command3_Click()
MsgBox "Program Backup Ini !!" & vbCrLf & _
      "Akan Menggandakan Folder Database Utama" & vbCrLf & _
      "Dengan Nama Folder Database yang sama," & vbCrLf & _
      "Pada Directory D,E.. yang anda pilih", vbOKOnly + vbInformation, "Backup"
End Sub
Private Sub Command4_Click()
Set backup = Nothing
Unload Me
End Sub


