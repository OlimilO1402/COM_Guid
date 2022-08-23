VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Guid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'  GUID = Globally  Unique Identifier       | the base of CLSID and UUID
' CLSID = Globally  Unique Class Identifier | a CLSID is a GUID that identifies a COM object
'  UUID = Universal Unique Identifier       | used in IDL files

'

'https://stackoverflow.com/questions/10687505/uuid-format-8-4-4-4-12-why
'https://www.ietf.org/rfc/rfc4122.txt
'see chapter: "4.1.2. Layout and Byte Order"
'  The internal representation of a UUID is a specific sequence of
'  bits in memory, as described in Section 4.  To accurately
'  represent a UUID as a URN, it is necessary to convert the bit
'  sequence to a string representation.
'
'  Each field is treated as an integer and has its value printed as a
'  zero-filled hexadecimal digit string with the most significant
'  digit first.  The hexadecimal values "a" through "f" are output as
'  lower case characters and are case insensitive on input.
'
'  The formal definition of the UUID string representation is
'  provided by the following ABNF [7]:
'
'  UUID                   = time-low "-" time-mid "-"
'                           time-high-and-version "-"
'                           clock-seq-and-reserved
'                           clock-seq-low "-" node
'  time-low               = 4hexOctet
'  time-mid               = 2hexOctet
'  time-high-and-version  = 2hexOctet
'  clock-seq-and-reserved = hexOctet
'  clock -seq - low = hexOctet
'  node                   = 6hexOctet
'  hexOctet               = hexDigit hexDigit
'  hexDigit =
'        "0" / "1" / "2" / "3" / "4" / "5" / "6" / "7" / "8" / "9" /
'        "a" / "b" / "c" / "d" / "e" / "f" /
'        "A" / "B" / "C" / "D" / "E" / "F"
Private Type TInt16
    Value As Integer
End Type
Private Type TInt162
    Value0 As Integer
    value1 As Integer
End Type
Private Type TInt32
    Value As Long
End Type
Private Type TB2
    Value(0 To 1) As Byte ' 2
End Type
Private Type TB4
    Value(0 To 3) As Byte ' 4
End Type
Private Type TB8
    Value(0 To 7) As Byte ' 8
End Type
Private Type TB16
    Value(0 To 15) As Byte
End Type
Private Type TGuid
    Data1 As Long          ' 4
    Data2 As Integer       ' 2
    Data3 As Integer       ' 2
    Data4(0 To 7) As Byte  ' 8
    Data5 As Long
End Type              'Sum: 16

Private m_Guid As TGuid

Private Declare Sub CoCreateGuid Lib "ole32" (ByRef PGuid As Any)

Private Declare Function StringFromCLSID Lib "ole32" (ByRef pCLSID As Any, ByRef pOleStr As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal pString As Long, ByRef pCLSID As Any) As Long

Private Declare Function StringFromGUID2 Lib "ole32" (ByRef PGuid As Any, ByVal lpsz As Long, ByVal cchMax As Long) As Long
'Declare Function GUIDFromString Lib "shell32" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long

' Windows crypto RNG API
Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (ByVal pBuff As Long, ByVal dwLen As Long) As Long

Friend Sub New_(ByVal a As Long, ByVal b As Integer, ByVal c As Integer, _
                ByVal d0 As Byte, ByVal d1 As Byte, ByVal d2 As Byte, ByVal d3 As Byte, _
                ByVal d4 As Byte, ByVal d5 As Byte, ByVal d6 As Byte, ByVal d7 As Byte)
    With m_Guid
        .Data1 = a: .Data2 = b: .Data3 = c
        .Data4(0) = d0: .Data4(1) = d1: .Data4(2) = d2: .Data4(3) = d3
        .Data4(4) = d4: .Data4(5) = d5: .Data4(6) = d6: .Data4(7) = d7
    End With
End Sub

Friend Sub NewD(ByVal a As Long, ByVal b As Integer, ByVal c As Integer, d() As Byte)
    New_ a, b, c, d(0), d(1), d(2), d(3), d(4), d(5), d(6), d(7)
End Sub

Friend Sub NewCo()
    CoCreateGuid m_Guid
End Sub

Friend Sub NewPK(ByVal sGuid As String, ByVal PID As Long)
    CLSIDFromString StrPtr(sGuid), ByVal VarPtr(m_Guid)
    m_Guid.Data5 = PID
End Sub

' ############################## '
'Public Function GuidCo() As Guid
'    Set Guid = New Guid: Guid.NewCo
'End Function
'Public Function Guid(ByVal a As Long, ByVal b As Integer, ByVal c As Integer, _
'                     ByVal d0 As Byte, ByVal d1 As Byte, ByVal d2 As Byte, ByVal d3 As Byte, _
'                     ByVal d4 As Byte, ByVal d5 As Byte, ByVal d6 As Byte, ByVal d7 As Byte) As Guid
'    Set Guid = New Guid: Guid.New_ a, b, c, d0, d1, d2, d3, d4, d5, d6, d7
'End Function
'Public Function GuidS(s As String) As Guid
'    Set GuidS = New Guid
'    If Not GuidS.Parse(s) Then Set GuidS = Nothing
'End Function
'Public Function GuidD(a As Long, b As Integer, c As Integer, d() As Byte) As Guid
'    Set GuidD = New Guid: GuidD.NewD a, b, c, d
'End Function
'Public Function UUID() As Guid
'    Set UUID = New Guid: UUID.Parse UUID.UUID
'End Function
'Public Function GuidPK(ByVal sGuid As String, ByVal PID As Long) As Guid
'    Set GuidPK = New Guid: GuidPK.NewPK sGuid, PID
'End Function
' ############################## '

Friend Function Clone() As Guid
    With m_Guid
        Set Clone = New Guid: Clone.NewD .Data1, .Data2, .Data3, .Data4
    End With
End Function

Public Property Let Data1(ByVal Value As Long)
    m_Guid.Data1 = Value
End Property
Public Property Get Data1() As Long
    Data1 = m_Guid.Data1
End Property

Public Property Let Data2(ByVal Value As Integer)
    m_Guid.Data2 = Value
End Property
Public Property Get Data2() As Integer
    Data2 = m_Guid.Data2
End Property

Public Property Let Data3(ByVal Value As Integer)
    m_Guid.Data3 = Value
End Property
Public Property Get Data3() As Integer
    Data3 = m_Guid.Data3
End Property

Public Property Let Data4(ByVal Index As Long, ByVal Value As Byte)
    If Index < 0 Or 7 < Index Then Exit Property
    m_Guid.Data4(Index) = Value
End Property
Public Property Get Data4(ByVal Index As Long) As Byte
    If Index < 0 Or 7 < Index Then Exit Property
    Data4 = m_Guid.Data4(Index)
End Property

Public Property Let Data5(ByVal Value As Long)
    m_Guid.Data5 = Value
End Property
Public Property Get Data5() As Long
    Data5 = m_Guid.Data5
End Property

Public Property Get Ptr()
    Ptr = VarPtr(m_Guid)
End Property

Friend Property Get IsEqual(other As Guid) As Boolean
    With m_Guid
        If .Data1 <> other.Data1 Then Exit Property
        If .Data2 <> other.Data2 Then Exit Property
        If .Data3 <> other.Data3 Then Exit Property
        Dim i As Long
        For i = 0 To 7
            If .Data4(i) <> other.Data4(i) Then Exit Property
        Next
    End With
    IsEqual = True
End Property

Friend Property Get IsSame(other As Guid) As Boolean
    IsSame = Me.Ptr = other.Ptr
End Property

Private Function Hex2(ByVal Value As Byte) As String
    Hex2 = Hex(Value): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
End Function

Private Function Hex4(ByVal Value As Integer) As String
    Hex4 = Hex(Value): If Len(Hex4) < 4 Then Hex4 = String(4 - Len(Hex4), "0") & Hex4
End Function

Private Function Hex8(ByVal Value As Long) As String
    Hex8 = Hex(Value): If Len(Hex8) < 8 Then Hex8 = String(8 - Len(Hex8), "0") & Hex8
End Function

Friend Function ToStr() As String
    'The function ToStr in pure VB
    With m_Guid
        ToStr = "{" & LCase(Hex8(.Data1) & "-" & Hex4(.Data2) & "-" & Hex4(.Data3) & "-" & Hex2(.Data4(0)) & Hex2(.Data4(1)) & "-" & Hex2(.Data4(2)) & Hex2(.Data4(3)) & Hex2(.Data4(4)) & Hex2(.Data4(5)) & Hex2(.Data4(6)) & Hex2(.Data4(7))) & "}"
    End With
End Function

Friend Function ToStr2() As String
    'The function ToStr by using WinAPI
    Dim s As String: s = String$(38, vbNullChar)
    Dim hr As Long: hr = StringFromGUID2(m_Guid, StrPtr(s), LenB(s))
    ToStr2 = s
End Function

Public Function Parse(ByVal s As String) As Boolean
    'try to have only hex-string at the end
    'filter everything non hex string
    s = Replace(s, "{", ""): s = Replace(s, "}", "")
    s = Replace(s, "(", ""): s = Replace(s, ")", "")
    s = Replace(s, "[", ""): s = Replace(s, "]", "")
    s = Replace(s, """", "")
    s = Replace(s, vbNullChar, "")
    
    Dim sa() As String, u As Long
    Dim ati As TInt16, atl As TInt32, atb4 As TB4, atb2 As TB2
    
    If InStr(1, s, "-") > 0 Then
        sa = Split(s, "-")
        u = UBound(sa)
        With m_Guid
            If u >= 0 Then .Data1 = CLng("&H" & sa(0))
            If u >= 1 Then .Data2 = CInt("&H" & sa(1))
            If u >= 2 Then .Data3 = CInt("&H" & sa(2))
            If u >= 3 Then
                ati.Value = CInt("&H" & sa(3))
                LSet atb2 = ati
                .Data4(0) = atb2.Value(1)
                .Data4(1) = atb2.Value(0)
                If u >= 4 Then
                    ati.Value = CInt("&H" & Left(sa(4), 4))
                    LSet atb2 = ati
                    .Data4(2) = atb2.Value(1)
                    .Data4(3) = atb2.Value(0)
                    atl.Value = CLng("&H" & Mid(sa(4), 5))
                    LSet atb4 = atl
                    .Data4(4) = atb4.Value(3)
                    .Data4(5) = atb4.Value(2)
                    .Data4(6) = atb4.Value(1)
                    .Data4(7) = atb4.Value(0)
                End If
            End If
        End With
        Parse = True
    ElseIf InStr(1, s, ",") > 0 Then
        sa = Split(s, ",")
        u = UBound(sa)
        With m_Guid
            If u >= 0 Then .Data1 = ParseHexOrDecLng(sa(0))
            If u >= 1 Then .Data2 = ParseHexOrDecInt(sa(1))
            If u >= 2 Then .Data3 = ParseHexOrDecInt(sa(2))
            If u >= 3 Then .Data4(0) = CByte(ParseHexOrDecInt(sa(3)))
            If u >= 4 Then .Data4(1) = CByte(ParseHexOrDecInt(sa(4)))
            If u >= 5 Then .Data4(2) = CByte(ParseHexOrDecInt(sa(5)))
            If u >= 6 Then .Data4(3) = CByte(ParseHexOrDecInt(sa(6)))
            If u >= 7 Then .Data4(4) = CByte(ParseHexOrDecInt(sa(7)))
            If u >= 8 Then .Data4(5) = CByte(ParseHexOrDecInt(sa(8)))
            If u >= 9 Then .Data4(6) = CByte(ParseHexOrDecInt(sa(9)))
            If u >= 10 Then .Data4(7) = CByte(ParseHexOrDecInt(sa(10)))
        End With
        Parse = True
    End If
End Function

Private Function ParseHexOrDecLng(ByVal s As String) As Long
    s = Trim(s)
    If Left(s, 2) = "0x" Then s = Replace(s, "0x", "&H")
    If Left(s, 2) = "&H" Then
        ParseHexOrDecLng = CLng(s)
    Else
        If IsNumeric(s) Then
            ParseHexOrDecLng = CLng(s)
        ElseIf IsHex(s) Then
            ParseHexOrDecLng = CLng("&H" & s)
        End If
    End If
End Function

Private Function ParseHexOrDecInt(ByVal s As String) As Long
    s = Trim(s)
    If Left(s, 2) = "0x" Then s = Replace(s, "0x", "&H")
    If Left(s, 2) = "&H" Then
        ParseHexOrDecInt = CInt(s)
    Else
        If IsNumeric(s) Then
            ParseHexOrDecInt = CInt(s)
        ElseIf IsHex(s) Then
            ParseHexOrDecInt = CInt("&H" & s)
        End If
    End If
End Function

Private Function IsHex(s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        Select Case Asc(Mid(s, i, 1))
        Case 48 To 57:  ' 0 - 9 OK weiter
        Case 65 To 70:  ' A - F OK weiter
        Case 97 To 102: ' a - f OK weiter
        Case Else: Exit Function
        End Select
    Next
    IsHex = True
End Function

Public Function ParseCLSID(ByVal s As String) As Boolean
    If Right(s, 1) <> vbNullChar Then s = s & vbNullChar ' create null-terminated OLESTR
    Dim lBuf As Long: lBuf = CLSIDFromString(StrPtr(s), m_Guid)
    ParseCLSID = lBuf > 0
End Function

' ############################## '    UUID    ' ############################## '

' =========================================================
' generates a v4 UUID (http://www.ietf.org/rfc/rfc4122.txt)
' =========================================================
Public Function UUID(Optional ByVal bBrk As Boolean = False, Optional ByVal bSep As Boolean = True) As String
    Dim sDigit As String, sUUID As String * 36, sRes As String
    
    ' initialize the code
    Dim cbVal As Byte
    Dim lPtr  As Long:  lPtr = VarPtr(cbVal)
    Dim lSize As Long: lSize = LenB(cbVal)
    
    ' UUID lenght is 36 chars (including separators)
    Dim i As Integer
    For i = 0 To 35
        If (i = 8) Or (i = 13) Or (i = 18) Or (i = 23) Then
            ' digit group separator
            If bSep Then
                sDigit = "-"
            Else
                sDigit = ""
            End If
        ElseIf (i = 14) Then
            ' bits 12-15 of "time_hi_and_version field" to 0010
            sDigit = "4"
        Else
            ' generate a random value between 0 and 15 (0 to F)
            RtlGenRandom lPtr, lSize
            cbVal = (cbVal Mod 16)
            If (i = 19) Then
                ' bits 6-7 of "clock_seq_hi_and_reserved" to 01
                cbVal = ((cbVal And 3) Or 8)
            End If
            sDigit = Hex(cbVal)
        End If
        ' update UUID
        Mid(sUUID, i + 1, 1) = sDigit
    Next
    
    ' trim spaces and convert to lcase as for RFC
    sRes = LCase(Trim(sUUID))
    
    ' need to add enclosing brackets (MS style) ?
    If bBrk Then
        sRes = "{" & sRes & "}"
    End If
    
    ' return generated and formatted UUID
    UUID = sRes
End Function

