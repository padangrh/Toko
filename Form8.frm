VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H0000C000&
   Caption         =   "Entri dan Update Kategori"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8205
   ControlBox      =   0   'False
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   2565
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Nama  Kategori"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "ENTRI DAN  UPDATE DATA KATEGORI"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsKategori, Rec As New ADODB.Recordset
Dim statuskategori, sql, sql1 As String
Private Sub Command1_Click()
Dim a As New ADODB.Recordset
If Trim(Text1) <> "" Then
   Set rsKategori = con.Execute("select * from tbkategori where kode='" & Trim(Text1) & "'")
   If rsKategori.EOF Then
      con.Execute ("Insert into tbkategori values('" & Text1 & "')")
   Else
      sql = "Update tbkategori set kode='" & Text1 & "' where kode='" & Text1.Text & "'"
      con.Execute (sql)
   End If
   kosongkan
Else
MsgBox "Nama Tidak Boleh Kosong"
End If
Form4.refreshlist
Text1.SetFocus
End Sub
Sub kosongkan()
Text1 = ""
End Sub
Private Sub Command2_Click()
kosongkan
End Sub
Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Form_Activate()
Text1.SetFocus
End Sub
Private Sub Form_Load()
If con.State = adStateClosed Then
connect
End If
kosongkan
End Sub
'Private Sub Text1_keypress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'If statuskategori = "perbaiki" Then
'    Text1.Enabled = False
'    Else
'     Text1.Enabled = True
'End If
'   Set rsKategori = con.Execute("Select * from tbkategori where kode='" & Trim(Text1) & "'")
'   If Not rsKategori.EOF Then
'      With Form7
'          .Text2 = rsKategori!nama
'
'
'      End With
'   Else
'      Text2 = ""
'
'      Text2.SetFocus
'    End If
'End If
'End Sub
Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub




