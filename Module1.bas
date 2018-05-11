Attribute VB_Name = "Module1"
Public fMainForm As FrmMain
Public con As New ADODB.Connection
Public rsSupplier As ADODB.Recordset
Public status
Public username As String
Public Setting_Object As Object
' **********************************************
' Posiflex usbpd.dll DLL
' **********************************************
Public Declare Function WritePD _
    Lib "usbpd.dll" _
    (ByVal data As String, ByVal Length As Long) _
As Long

Public Declare Function WritePD80 _
    Lib "usbpd.dll" Alias "WritePD" _
    (ByRef data As Any, ByVal Length As Long) _
As Long

Public Declare Function PdState _
    Lib "usbpd.dll" _
    () _
As Long

Public Declare Function OpenUSBpd _
    Lib "usbpd.dll" _
    () _
As Long

Public Declare Function CloseUSBpd _
    Lib "usbpd.dll" _
    () _
As Long

Declare Sub Sleep Lib "kernel32" _
   (ByVal dwMilliseconds As Long)
   
Public Sub connect()
'    con.ConnectionString = "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=data"
    con.ConnectionString = "Provider=MSDASQL.1;Password=" & Setting_Object("DB_Pw") & ";Persist Security Info=True;User ID=" & Setting_Object("DB_Id") & ";Data Source=" & Setting_Object("DB_Name")
    con.Open
End Sub

Public Function getSupplier(kode As String) As Boolean
    If rsSupplier Is Nothing Then
        Set rsSupplier = con.Execute("select * from tbsuplier")
    End If
    
    Dim found As Boolean
    found = False
    If Not rsSupplier.EOF Then
        rsSupplier.MoveFirst
        Do While Not rsSupplier.EOF
            If kode = rsSupplier!kdsuplier Then
                found = True
                Exit Do
            End If
            rsSupplier.MoveNext
        Loop
        If found = False Then rsSupplier.MoveFirst
    End If
    getSupplier = found
End Function

Public Function priceToNum(price As String) As Currency
    price = Replace(price, ".", "")
    price = Replace(price, ",", "")
    priceToNum = Val(price)
End Function

Public Function isMaster() As Boolean
    isMaster = (status = "Master")
End Function

Public Function isSPV() As Boolean
    isSPV = (status = "Supervisor")
End Function

Sub Main()
    Set fMainForm = New FrmMain
    fMainForm.Show
End Sub

''tambahan fingerprint
'Function arraytohex(arr() As Byte) As String
'    Dim templatestr As String
'    Dim tempstr As String
'    Dim i As Integer
'    templatestr = ""
'    For i = LBound(arr) To UBound(arr)
'        tempstr = Hex$(arr(i))
'        If Len(tempstr) = 1 Then tempstr = "0" + tempstr 'padHex
'        templatestr = templatestr + tempstr
'    Next i
'    arraytohex = templatestr
'End Function
'
'Public Sub hextoarray(inphex As String, outarray() As Byte)
'    ReDim outarray(0 To Len(inphex) / 2) As Byte
'    DoEvents
'    Dim i As Integer
'    For i = 1 To Len(inphex) Step 2
'        outarray(((i + 1) / 2) - 1) = Val("&H" + Mid$(inphex, i, 2))
'    Next i
'End Sub
'
'Function Kualitas(X As AISampleQuality)
'    If X = Sq_Good Then
'        Kualitas = "Hasil Bagus"
'    Else
'        Kualitas = "Hasil Jelek"
'    End If
'End Function
''end tambahan fingerprint

Function validateKey(KeyAscii As Integer, lim As Integer)
    If lim = 1 Then 'for number
        Select Case KeyAscii
            Case 48 To 57, 44, 45, 46, 8 '0-9, comma, minus, dot and backspace
            'Let these key codes pass through
            Case Else
            'All others get trapped
            KeyAscii = 0 ' set ascii 0 to trap others input
        End Select
    ElseIf lim = 2 Then 'for password/kode
        Select Case KeyAscii
            Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
            'Let these key codes pass through
            Case Else
            'All others get trapped
            KeyAscii = 0 ' set ascii 0 to trap others input
        End Select
    ElseIf lim = 3 Then 'for general textbox
        Select Case KeyAscii
            Case 8, 32 To 38, 40 To 58, 60 To 126 ' Allow all except ' and ;
            'Let these key codes pass through
            Case Else
            'All others get trapped
            KeyAscii = 0 ' set ascii 0 to trap others input
        End Select
    End If
    validateKey = KeyAscii
End Function
