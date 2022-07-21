VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = "Testing Class Guid " & App.Major & "." & App.Minor & "." & App.Revision
    With List1
        .Clear
        .AddItem """{64440490-4c8b-11d1-8b70-080036b11a03}"""
        .AddItem "{64440490-4C8B-11D1-8B70-080036B11A03}"
        '64440490-4C8B-11D1-8B70-080036B11A03
        .AddItem "0x64440490, 0x4C8B, 0x11D1, 0x8B, 0x70, 0x08, 0x00, 0x36, 0xB1, 0x1A, 0x03"
        '.AddItem ""
        '.AddItem ""
    End With
End Sub

Private Sub Command1_Click()
    
    Dim s As String: s = "{64440490-4c8b-11d1-8b70-080036b11a03}"
    
    Dim g1 As Guid: Set g1 = MNew.GuidS(s)
    Dim s1 As String:   s1 = g1.ToStr
    
    Dim g2 As Guid: Set g2 = MNew.GuidS(s1)
    Dim s2 As String:   s2 = g2.ToStr2
    
    MsgBox s & vbCrLf & s1 & vbCrLf & s2
    
End Sub

Private Sub Command2_Click()
    
    Dim g1 As Guid: Set g1 = MNew.GuidCo
    Dim s1 As String:   s1 = g1.ToStr2
    
    Dim g2 As Guid: Set g2 = MNew.GuidS(s1)
    Dim s2 As String:   s2 = g2.ToStr
    
    MsgBox s1 & vbCrLf & s2 & vbCrLf & "g1 = g2: " & g1.IsEqual(g2)
    
End Sub

Private Sub Command3_Click()
    Dim g1 As Guid: Set g1 = MNew.GuidCo
    Dim s1 As String:   s1 = g1.ToStr2
    '
    
End Sub

Private Sub Command4_Click()
    'Dim g As Guid: Set g = MNew
End Sub

Private Sub List1_Click()
    
    Dim i As Long:   i = List1.ListIndex
    Dim s As String: s = List1.List(i)
    
    Dim g1 As Guid: Set g1 = MNew.GuidS(s)
    Dim s1 As String:   s1 = g1.ToStr
    
    Dim g2 As Guid: Set g2 = MNew.GuidS(s1)
    Dim s2 As String: s2 = g1.ToStr & " = " & g2.ToStr
    
    Label1.Caption = s & vbCrLf & s1 & vbCrLf & s2 & vbCrLf & g1.IsEqual(g2)
    
End Sub
