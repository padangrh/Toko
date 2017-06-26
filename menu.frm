VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Main Menu"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15060
   Icon            =   "menu.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "menu.frx":628A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":9F68D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":9F6DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":9F7076
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":9F7364
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":9F76CB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Penjualan"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Laporan"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Barang"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Admin"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Logout"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.TextBox Text2 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   7080
         TabIndex        =   2
         Text            =   "0"
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   1
         Text            =   "0"
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Menu p 
      Caption         =   "Penjualan"
      Begin VB.Menu pjl 
         Caption         =   "Trans. Penjualan"
      End
      Begin VB.Menu lgf 
         Caption         =   "Logoff"
      End
   End
   Begin VB.Menu l 
      Caption         =   "Laporan"
      Begin VB.Menu lpr 
         Caption         =   "Laporan"
      End
   End
   Begin VB.Menu b 
      Caption         =   "Barang"
      Begin VB.Menu sp 
         Caption         =   "Entri Suplier"
      End
      Begin VB.Menu ebr 
         Caption         =   "Entri Barang / Stock"
      End
      Begin VB.Menu pb 
         Caption         =   "Entri Pembelian"
      End
      Begin VB.Menu warning 
         Caption         =   "Warning"
         Index           =   5
      End
   End
   Begin VB.Menu a 
      Caption         =   "Admin"
      Begin VB.Menu tu 
         Caption         =   "User Manager"
      End
   End
   Begin VB.Menu pa 
      Caption         =   "Pengaturan Akun"
   End
   Begin VB.Menu lg 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim active_form As Form
Dim fLogout As Boolean

Private Sub changeForm(new_form As Form)
    If active_form Is new_form Then
        Exit Sub
    End If
    
    new_form.Show
    
    If Not active_form Is Nothing Then
        Unload active_form
    End If
    Set active_form = new_form
End Sub

Private Sub edb_Click()
    Call changeForm(Form_List_barang)
End Sub

Private Sub eds_Click()
    Call changeForm(Form_List_Supplier)
End Sub

Private Sub k_Click()
    Unload Me
End Sub

Private Sub bd_Click()
    Call changeForm(Form_List_Supplier)
End Sub

Private Sub ebr_Click()
    Call changeForm(Form_List_barang)
End Sub

Private Sub lg_Click()
    Unload Me
End Sub

Private Sub lgf_Click()
    fLogout = True
    logoff
End Sub

Public Sub logoff()
    If Not active_form Is Nothing Then
        Unload active_form
        Set active_form = Nothing
    End If
    Unload Me
    username = ""
    status = ""
    frmlogin.Show
End Sub

Private Sub lpr_Click()
    Form_Laporan.Show (1)
End Sub

Private Sub pa_Click()
    Form_RegisterFP.Show (1)
End Sub

Private Sub pb_Click()
    Call changeForm(Form_List_beli)
End Sub

Private Sub pjl_Click()
    Call changeForm(Form_List_Jual)
End Sub

Private Sub rbl_Click()
    Call changeForm(Form_List_Supplier)
End Sub

Private Sub rjl_Click()
    Call changeForm(Form_List_Supplier)
End Sub

Private Sub sp_Click()
    Call changeForm(Form_List_Supplier)
End Sub

Private Sub td_Click()
    If MsgBox("Jalankan Trim Data?", vbYesNo, "Trim Data") = vbYes Then
        Dim rsHistory As ADODB.Recordset
        Dim rsTB As ADODB.Recordset
        Dim countDeleted As Integer
        countDeleted = 0
        Set rsHistory = con.Execute("Select * from tbbarang_history")
        Do While Not rsHistory.EOF
            Set rsTB = con.Execute("Select * from tbbarang where kode = '" & rsHistory!kode & "'")
            If rsTB.EOF Then
                con.Execute ("Delete from tbbarang_history where kode = '" & rsHistory!kode & "'")
                countDeleted = countDeleted + 1
            End If
            rsHistory.MoveNext
        Loop
        Set rsHistory = Nothing
        Set rsTB = Nothing
        MsgBox (countDeleted & " data dihapus")
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
    Case 1: PopupMenu p
    Case 2: PopupMenu l
    Case 3: PopupMenu b
    Case 4: PopupMenu a
    Case 5: Unload Me
    End Select
End Sub


Private Sub MDIForm_Activate()
    If username = "" Then
       frmlogin.Show 1
    End If
End Sub

Private Sub MDIForm_Load()
'    Set active_form = Nothing
'
'    If con.State = adStateClosed Then
'        connect
'    End If
    Set active_form = Nothing
    Dim temp_stringX, File_StringX As String
    Dim fso As FileSystemObject
    fLogout = False
    Set fso = New FileSystemObject
    If fso.FileExists(App.Path & "\Settings.json") Then
        'load JSON file
        temp_stringX = ReadTextFile(App.Path & "\Settings.json")
        'Decode file
        File_StringX = Base64DecodeString(temp_stringX)
        'Generate variables
        Set Setting_Object = JSON.parse(File_StringX)
        
        If con.State = adStateClosed Then
            connect
        End If
     Else
        MsgBox "Settings file is missing."
        Unload Me
    End If
    
End Sub

Private Sub MDIForm_Unload(cancel As Integer)
    If Setting_Object("Absen") Then
        Dim Rec As ADODB.Recordset
        Set Rec = con.Execute("select * from tbabsen where userid = '" & username & "' and tanggal = '" & Format(Now, "yyyy-MM-dd") & "'")
        If Not Rec.EOF Then
            con.Execute ("update tbabsen set jam_keluar = '" & Format(Now, "HH:mm:ss") & "' where userid = '" & username & "' and tanggal = '" & Format(Now, "yyyy-MM-dd") & "'")
        End If
        Set Rec = Nothing
    End If
    Dim Form As VB.Form
    For Each Form In VB.Forms
        Unload Form
    Next
    If fLogout = False Then
        con.Close
    End If
End Sub

Private Sub tu_Click()
    Call changeForm(Form_User)
End Sub

Private Sub uh_Click()
    If MsgBox("Jalankan update 1?", vbYesNo, "Update 1") = vbYes Then
        Dim rsHistory As ADODB.Recordset
        Dim tempHarga As Long
        Dim tempKode As String
        tempHarga = 0
        tempKode = ""
        Set rsHistory = con.Execute("select * from v_update1")
        Do While Not rsHistory.EOF
            If rsHistory!kode = tempKode And rsHistory!harga_jual <> tempHarga And tempKode <> "" Then
                'editV2
                con.Execute ("insert into tbbarang_history (kode, nama, tanggal, harga_modal, harga_jual) values ('" & tempKode & "','" & rsHistory!nama_barang & "','" & Format(rsHistory!tglbukti, "yyyy-MM-dd") & "','0','" & rsHistory!harga_jual & "')")
            End If
            tempKode = rsHistory!kode
            tempHarga = rsHistory!harga_jual
            rsHistory.MoveNext
        Loop
        Set rsHistory = Nothing
    '    rsHistory.Close
        MsgBox "Update 1 berhasil"
    End If
End Sub

Private Sub uh2_Click()
    If MsgBox("Jalankan update 2?", vbYesNo, "Update 2") = vbYes Then
        Dim rsHistory As ADODB.Recordset
        Dim rsCekTb As ADODB.Recordset
        Dim tempHarga As Long
        Dim tempKode As String
        Dim tempTanggal As Date
        tempHarga = 0
        tempKode = ""
        tempTanggal = 0
        Set rsHistory = con.Execute("select * from v_update2")
        Do While Not rsHistory.EOF
            If rsHistory!kode = tempKode And rsHistory!harga <> tempHarga And tempKode <> "" Then
                Set rsCekTb = con.Execute("Select * from tbbarang_history where tanggal = '" & Format(rsHistory!tglbukti, "yyyy-MM-dd") & "' and kode = '" & tempKode & "'")
                If rsCekTb.EOF Then
                    'editV2
                    con.Execute ("insert into tbbarang_history (kode, nama, tanggal, harga_modal, harga_jual) values ('" & tempKode & "','" & rsHistory!nama_barang & "','" & Format(rsHistory!tglbukti, "yyyy-MM-dd") & "','" & rsHistory!harga & "','0')")
                Else
                    If rsCekTb!harga_modal = 0 Then
                        con.Execute ("update tbbarang_history set harga_modal =  '" & rsHistory!harga & "' where tanggal = '" & Format(rsHistory!tglbukti, "yyyy-MM-dd") & "' and kode = '" & tempKode & "'")
                    Else
                        'editV2
                        con.Execute ("insert into tbbarang_history (kode, nama, tanggal, harga_modal, harga_jual) values ('" & tempKode & "','" & rsHistory!nama_barang & "','" & Format(rsHistory!tglbukti, "yyyy-MM-dd") & "','" & rsHistory!harga & "','0')")
                    End If
                End If
            End If
            tempKode = rsHistory!kode
            tempHarga = rsHistory!harga
            rsHistory.MoveNext
        Loop
        Set rsCekTb = Nothing
        Set rsHistory = Nothing
    '    rsHistory.Close
        MsgBox "Update 2 berhasil"
    End If
End Sub

Private Sub uh3_Click()
    If MsgBox("Jalankan update 3?", vbYesNo, "Update 3") = vbYes Then
        Dim rsHistory As ADODB.Recordset
        Dim rsTB As ADODB.Recordset
        Set rsHistory = con.Execute("select * from tbbarang_history")
        Do While Not rsHistory.EOF
'            Set rsTB = con.Execute("select * from tbbeli where kode = '" & rsHistory!kode & "' and tglbukti <= '" & rsHistory!tanggal & "' order by tglbukti desc limit 1")
'            If Not rsTB.EOF Then
'                If rsHistory!harga_modal = 0 Then
'                    con.Execute ("update tbbarang_history set harga_modal = '" & rsTB!harga_modal & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
'                ElseIf rsHistory!harga_jual = 0 Then
'                    con.Execute ("update tbbarang_history set harga_jual = '" & rsTB!harga_jual & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
'                End If
'            End If
            If rsHistory!harga_modal = 0 Then
                Set rsTB = con.Execute("select * from tbbeli where kode = '" & rsHistory!kode & "' and tglbukti <= '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' order by tglbukti desc limit 1")
                If Not rsTB.EOF Then
                    con.Execute ("update tbbarang_history set harga_modal = '" & rsTB!harga & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                End If
            ElseIf rsHistory!harga_jual = 0 Then
                Set rsTB = con.Execute("select * from tbjual where kode = '" & rsHistory!kode & "' and tglbukti <= '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' order by tglbukti desc limit 1")
                If Not rsTB.EOF Then
                    con.Execute ("update tbbarang_history set harga_jual = '" & rsTB!harga_jual & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                End If
            End If
            rsHistory.MoveNext
        Loop
        Set rsTB = Nothing
        Set rsHistory = Nothing
        MsgBox "Update 3 berhasil"
    End If
End Sub

Private Sub uh4_Click()
    If MsgBox("Jalankan update 4?", vbYesNo, "Update 4") = vbYes Then
        Dim rsHistory As ADODB.Recordset
        Dim rsTB As ADODB.Recordset
        Set rsHistory = con.Execute("select * from tbbarang_history")
        Do While Not rsHistory.EOF
'            Set rsTB = con.Execute("select * from tbbeli where kode = '" & rsHistory!kode & "' and tglbukti <= '" & rsHistory!tanggal & "' order by tglbukti desc limit 1")
'            If Not rsTB.EOF Then
'                If rsHistory!harga_modal = 0 Then
'                    con.Execute ("update tbbarang_history set harga_modal = '" & rsTB!harga_modal & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
'                ElseIf rsHistory!harga_jual = 0 Then
'                    con.Execute ("update tbbarang_history set harga_jual = '" & rsTB!harga_jual & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
'                End If
'            End If
            If rsHistory!harga_modal = 0 Then
                Set rsTB = con.Execute("select * from tbbeli where kode = '" & rsHistory!kode & "' and tglbukti >= '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' order by tglbukti asc limit 1")
                If Not rsTB.EOF Then
                    con.Execute ("update tbbarang_history set harga_modal = '" & rsTB!harga & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                Else
                    Set rsTB = con.Execute("Select * from tbbarang where kode = '" & rsHistory!kode & "'")
                    If Not rsTB.EOF Then
                        con.Execute ("update tbbarang_history set harga_modal = '" & rsTB!harga_modal & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                    End If
                End If
            ElseIf rsHistory!harga_jual = 0 Then
                Set rsTB = con.Execute("select * from tbjual where kode = '" & rsHistory!kode & "' and tglbukti >= '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' order by tglbukti asc limit 1")
                If Not rsTB.EOF Then
                    con.Execute ("update tbbarang_history set harga_jual = '" & rsTB!harga_jual & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                Else
                    Set rsTB = con.Execute("Select * from tbbarang where kode = '" & rsHistory!kode & "'")
                    If Not rsTB.EOF Then
                        con.Execute ("update tbbarang_history set harga_jual = '" & rsTB!harga_jual & "' where tanggal = '" & Format(rsHistory!tanggal, "yyyy-MM-dd") & "' and kode = '" & rsHistory!kode & "'")
                    End If
                End If
            End If
            rsHistory.MoveNext
        Loop
        Set rsTB = Nothing
        Set rsHistory = Nothing
        MsgBox "Update 4 berhasil"
    End If
End Sub

Private Sub warning_Click(index As Integer)
    Form_Warning.Show (1)
End Sub

Public Function ReadTextFile(sFilePath As String) As String
   On Error Resume Next
   
   Dim handle As Integer
   If LenB(Dir$(sFilePath)) > 0 Then
   
      handle = FreeFile
      Open sFilePath For Binary As #handle
      ReadTextFile = Space$(LOF(handle))
      Get #handle, , ReadTextFile
      Close #handle
      
   End If
   
End Function
