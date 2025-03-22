Attribute VB_Name = "MVBGuid"
Option Explicit

Public Const sIID_IUnknown     As String = "{00000000-0000-0000-C000-000000000046}"
Public Const sIID_IDispatch    As String = "{00020400-0000-0000-C000-000000000046}"
Public Const sIID_IEnumVariant As String = "{00020404-0000-0000-C000-000000000046}"

Public Type VBGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data5(0 To 7) As Byte
End Type

Private Type TByteHiLo
    Lo As Byte
    Hi As Byte
End Type

Private Type TInteger
    Value As Integer
End Type

Public IID_IUnknown     As VBGuid
Public IID_IDispatch    As VBGuid
Public IID_IEnumVariant As VBGuid

'https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateguid
Private Declare Sub CoCreateGuid Lib "ole32" (ByRef pGuid_out As Any)

'https://learn.microsoft.com/de-de/windows/win32/api/combaseapi/nf-combaseapi-stringfromclsid
Private Declare Function StringFromCLSID Lib "ole32" (ByRef pCLSID As Any, ByRef pOleStr As LongPtr) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/combaseapi/nf-combaseapi-clsidfromstring
Private Declare Function CLSIDFromString Lib "ole32" (ByVal pString As LongPtr, ByRef pCLSID As Any) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/combaseapi/nf-combaseapi-stringfromguid2
Private Declare Function StringFromGUID2 Lib "ole32" (ByRef pGuid As Any, ByVal lpsz As LongPtr, ByVal cchMax As Long) As Long

'Private Declare Function GUIDFromString Lib "shell32" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long

' Windows crypto RNG API
Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (ByVal pBuff As Long, ByVal dwLen As Long) As Long

Public Sub Init()
    IID_IUnknown = New_VBGuidS(sIID_IUnknown)
    IID_IDispatch = New_VBGuidS(sIID_IDispatch)
    IID_IEnumVariant = New_VBGuidS(sIID_IEnumVariant)
End Sub

'VBGuid
Public Function New_VBGuid(ByVal Data1 As Long, ByVal Data2 As Integer, ByVal Data3 As Integer, _
                           ByVal Data50 As Byte, ByVal Data51 As Byte, ByVal Data52 As Byte, ByVal Data53 As Byte, _
                           ByVal Data54 As Byte, ByVal Data55 As Byte, ByVal Data56 As Byte, ByVal Data57 As Byte) As VBGuid
    With New_VBGuid: .Data1 = Data1: .Data2 = Data2: .Data3 = Data3
        .Data5(0) = Data50: .Data5(1) = Data51: .Data5(2) = Data52: .Data5(3) = Data53
        .Data5(4) = Data54: .Data5(5) = Data55: .Data5(6) = Data56: .Data5(7) = Data57
    End With
End Function

Public Function New_VBGuid4(ByVal Data1 As Long, ByVal Data2 As Integer, ByVal Data3 As Integer, ByVal Data4 As Integer, _
                            ByVal Data52 As Byte, ByVal Data53 As Byte, ByVal Data54 As Byte, _
                            ByVal Data55 As Byte, ByVal Data56 As Byte, ByVal Data57 As Byte) As VBGuid
    Dim i4 As TInteger: i4.Value = Data4
    Dim b2 As TByteHiLo: LSet b2 = i4
    New_VBGuid4 = New_VBGuid(Data1, Data2, Data3, b2.Hi, b2.Lo, Data52, Data53, Data54, Data55, Data56, Data57)
End Function

Public Function New_VBGuidS4(ByVal shx1 As String, ByVal shx2 As String, ByVal shx3 As String, ByVal shx4 As String, _
                             ByVal shx52 As String, ByVal shx53 As String, ByVal shx54 As String, ByVal shx55 As String, ByVal shx56 As String, ByVal shx57 As String) As VBGuid
    Dim Data1 As Long:    Data1 = CLng("&H" & shx1)
    Dim Data2 As Integer: Data2 = CInt("&H" & shx2)
    Dim Data3 As Integer: Data3 = CInt("&H" & shx3)
    Dim Data4 As Integer: Data4 = CInt("&H" & shx4)
    Dim Data5(2 To 7) As Byte
    Data5(2) = CByte("&H" & shx52): Data5(3) = CByte("&H" & shx53): Data5(4) = CByte("&H" & shx54)
    Data5(5) = CByte("&H" & shx55): Data5(6) = CByte("&H" & shx56): Data5(7) = CByte("&H" & shx57)
    New_VBGuidS4 = New_VBGuid4(Data1, Data2, Data3, Data4, Data5(2), Data5(3), Data5(4), Data5(5), Data5(6), Data5(7))
End Function

Public Function New_VBGuidS45(ByVal shx1 As String, ByVal shx2 As String, ByVal shx3 As String, ByVal shx4 As String, ByVal shx5 As String) As VBGuid
    Dim l As Long: l = Len(shx5)
    Dim shx5_(2 To 7) As String
    If 2 <= l Then shx5_(2) = Mid(shx5, 1, 2)
    If 4 <= l Then shx5_(3) = Mid(shx5, 3, 2)
    If 6 <= l Then shx5_(4) = Mid(shx5, 5, 2)
    If 8 <= l Then shx5_(5) = Mid(shx5, 7, 2)
    If 10 <= l Then shx5_(6) = Mid(shx5, 9, 2)
    If 12 <= l Then shx5_(7) = Mid(shx5, 11, 2)
    New_VBGuidS45 = New_VBGuidS4(shx1, shx2, shx3, shx4, shx5_(2), shx5_(3), shx5_(4), shx5_(5), shx5_(6), shx5_(7))
End Function

Public Function New_VBGuidS(ByVal sIID As String) As VBGuid
    VBGuid_Parse New_VBGuidS, sIID
'    Dim hr As Long: hr = CLSIDFromString(StrPtr(sIID), New_VBGuidS)
'    If hr = 0 Then
'        MsgBox "Error creating guid from string: '" & sIID & "'"
'    End If
'    sIID = Trim(sIID): Dim s As String
'    s = Left(sIID, 1):  If s = "{" Then sIID = Mid(sIID, 2)
'    s = Right(sIID, 1): If s = "}" Then sIID = Left(sIID, Len(sIID) - 1)
'    Dim sa() As String: sa = Split(sIID, "-")
'    Dim u As Long: u = UBound(sa)
'    Dim Data1 As String: If 0 <= u Then Data1 = sa(0)
'    Dim Data2 As String: If 1 <= u Then Data2 = sa(1)
'    Dim Data3 As String: If 2 <= u Then Data3 = sa(2)
'    Dim Data4 As String: If 3 <= u Then Data4 = sa(3)
'    Dim Data5 As String: If 4 <= u Then Data5 = sa(4)
'    New_VBGuidS = New_VBGuidS45(Data1, Data2, Data3, Data4, Data5) '(2), Data5(2), Data5(2), Data5(2), Data5(2), Data5(2))
End Function

Public Sub VBGuid_Parse(this As VBGuid, ByVal sIID As String)
    'sIID = String(40, 0)
    Dim hr As Long: hr = CLSIDFromString(StrPtr(sIID), this)
    If hr <> 0 Then
        MsgBox "Error creating guid from string: '" & sIID & "'"
    End If
'    sIID = Trim(sIID): Dim s As String
'    s = Left(sIID, 1):  If s = "{" Then sIID = Mid(sIID, 2)
'    s = Right(sIID, 1): If s = "}" Then sIID = Left(sIID, Len(sIID) - 1)
'    Dim sa() As String: sa = Split(sIID, "-")
'    Dim u As Long: u = UBound(sa)
'    Dim Data1 As String: If 0 <= u Then Data1 = sa(0)
'    Dim Data2 As String: If 1 <= u Then Data2 = sa(1)
'    Dim Data3 As String: If 2 <= u Then Data3 = sa(2)
'    Dim Data4 As String: If 3 <= u Then Data4 = sa(3)
'    Dim Data5 As String: If 4 <= u Then Data5 = sa(4)
'    this = New_VBGuidS45(Data1, Data2, Data3, Data4, Data5) '(2), Data5(2), Data5(2), Data5(2), Data5(2), Data5(2))
End Sub

Public Function VBGuid_ToStr(this As VBGuid) As String
    VBGuid_ToStr = String(40, 0)
    Dim hr As Long: hr = StringFromGUID2(this, StrPtr(VBGuid_ToStr), 40)
    VBGuid_ToStr = MString.Trim0(VBGuid_ToStr)
'    Dim i As Long, s As String: s = "{"
'    With this
'        s = s & Hex8(.Data1) & "-"
'        s = s & Hex4(.Data2) & "-"
'        s = s & Hex4(.Data3) & "-"
'        Dim b4 As TByteHiLo: b4.Hi = .Data5(0): b4.Lo = .Data5(1)
'        Dim Data4 As TInteger: LSet Data4 = b4
'        s = s & Hex4(Data4.Value) & "-"
'        For i = 2 To UBound(.Data5): s = s & Hex2(.Data5(i)): Next
'    End With
'    VBGuid_ToStr = s & "}"
End Function

Public Function VBGuid_Equals(this As VBGuid, other As VBGuid) As Boolean
    With other
        If .Data1 <> this.Data1 Then Exit Function
        If .Data2 <> this.Data2 Then Exit Function
        If .Data3 <> this.Data3 Then Exit Function
        If .Data5(0) <> this.Data5(0) Then Exit Function: If .Data5(1) <> this.Data5(1) Then Exit Function
        If .Data5(2) <> this.Data5(2) Then Exit Function: If .Data5(3) <> this.Data5(3) Then Exit Function
        If .Data5(4) <> this.Data5(4) Then Exit Function: If .Data5(5) <> this.Data5(5) Then Exit Function
        If .Data5(6) <> this.Data5(6) Then Exit Function: If .Data5(7) <> this.Data5(7) Then Exit Function
    End With
    VBGuid_Equals = True
End Function

'Private Function Hex2(ByVal b As Byte) As String
'    Hex2 = Hex(b): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
'End Function
'Private Function Hex4(ByVal i As Integer) As String
'    Hex4 = Hex(i): If Len(Hex4) < 4 Then Hex4 = String(4 - Len(Hex4), "0") & Hex4
'End Function
'Private Function Hex8(ByVal l As Long) As String
'    Hex8 = Hex(l): If Len(Hex8) < 8 Then Hex8 = String(8 - Len(Hex8), "0") & Hex8
'End Function


