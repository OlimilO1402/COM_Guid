Attribute VB_Name = "MNew"
Option Explicit

Public Function GuidCo() As Guid
    Set GuidCo = New Guid: GuidCo.NewCo
End Function

Public Function Guid(ByVal a As Long, ByVal b As Integer, ByVal c As Integer, _
                     ByVal d0 As Byte, ByVal d1 As Byte, ByVal d2 As Byte, ByVal d3 As Byte, _
                     ByVal d4 As Byte, ByVal d5 As Byte, ByVal d6 As Byte, ByVal d7 As Byte) As Guid
    Set Guid = New Guid: Guid.New_ a, b, c, d0, d1, d2, d3, d4, d5, d6, d7
End Function

Public Function GuidS(s As String) As Guid
    Set GuidS = New Guid
    If Not GuidS.Parse(s) Then Set GuidS = Nothing
End Function

Public Function GuidD(a As Long, b As Integer, c As Integer, d() As Byte) As Guid
    Set GuidD = New Guid: GuidD.NewD a, b, c, d
End Function

Public Function UUID() As Guid
    Set UUID = New Guid: UUID.Parse UUID.UUID
End Function
