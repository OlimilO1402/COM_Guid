Attribute VB_Name = "MNew"
Option Explicit

Public Function GuidCo() As Guid
    Set GuidCo = New Guid: GuidCo.NewCo
End Function

Public Function Guid(ByVal Data1 As Long, ByVal Data2 As Integer, ByVal Data3 As Integer, _
                     ByVal Data50 As Byte, ByVal Data51 As Byte, ByVal Data52 As Byte, ByVal Data53 As Byte, _
                     ByVal Data54 As Byte, ByVal Data55 As Byte, ByVal Data56 As Byte, ByVal Data57 As Byte) As Guid
    Set Guid = New Guid: Guid.New_ Data1, Data2, Data3, Data50, Data51, Data52, Data53, Data54, Data55, Data56, Data57
End Function

Public Function GuidS(s As String) As Guid
    Set GuidS = New Guid
    If Not GuidS.Parse(s) Then Set GuidS = Nothing
End Function

Public Function GuidD(Data1 As Long, Data2 As Integer, Data3 As Integer, d() As Byte) As Guid
    Set GuidD = New Guid: GuidD.NewD Data1, Data2, Data3, d
End Function

Public Function UUID() As Guid
    Set UUID = New Guid: UUID.Parse UUID.UUID
End Function

Public Function GuidPK(ByVal sGuid As String, ByVal PID As Long) As Guid
    Set GuidPK = New Guid: GuidPK.NewPK sGuid, PID
End Function
