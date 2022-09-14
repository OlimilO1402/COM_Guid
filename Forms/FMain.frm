VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "GUID/CLSID/UUID"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbDecHex 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FMain.frx":0000
      Left            =   4440
      List            =   "FMain.frx":000A
      TabIndex        =   21
      Text            =   "Hex"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton BtnCreateUUID 
      Caption         =   "Create UUID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnCreateCLSID 
      Caption         =   "Create CLSID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnCreateGUID 
      Caption         =   "Create GUID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "GUID / CLSID / UUID: 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7455
      Begin VB.TextBox TxtData56 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6720
         TabIndex        =   17
         Text            =   "255"
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox TxtData55 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6240
         TabIndex        =   16
         Text            =   "255"
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox TxtData54 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5760
         TabIndex        =   15
         Text            =   "255"
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox TxtData53 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5280
         TabIndex        =   14
         Text            =   "255"
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox TxtData52 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4800
         TabIndex        =   13
         Text            =   "255"
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox TxtData51 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         TabIndex        =   11
         Text            =   "255"
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox TxtData42 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3720
         TabIndex        =   10
         Text            =   "255"
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox TxtData41 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3240
         TabIndex        =   8
         Text            =   "255"
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox TxtData3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         TabIndex        =   6
         Text            =   "65536"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtData2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         TabIndex        =   4
         Text            =   "65536"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtData1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   2
         Text            =   "2147483647"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label LblData5 
         AutoSize        =   -1  'True
         Caption         =   "Data5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4320
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LblData4 
         AutoSize        =   -1  'True
         Caption         =   "Data4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3240
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LblData3 
         AutoSize        =   -1  'True
         Caption         =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LblData2 
         AutoSize        =   -1  'True
         Caption         =   "Data2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LblData1 
         AutoSize        =   -1  'True
         Caption         =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   7455
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_List As Collection
Private m_Guid As Guid
Private m_Indx As Long

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set m_List = New Collection
    BtnCreateGUID_Click
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = List1.Left
    Dim T As Single: T = List1.Top
    Dim W As Single: W = Me.ScaleWidth - 2 * L
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move L, T, W, H
End Sub

Private Sub BtnCreateGUID_Click()
    Set m_Guid = MNew.GuidCo
    m_List.Add m_Guid
    List1.AddItem m_Guid.ToStr
    UpdateView
End Sub

Private Sub BtnCreateCLSID_Click()
    Set m_Guid = MNew.GuidCo
    m_List.Add m_Guid
    List1.AddItem m_Guid.ToStr
    m_Indx = m_List.Count - 1
    UpdateView
End Sub

Private Sub BtnCreateUUID_Click()
    Set m_Guid = MNew.UUID
    m_List.Add m_Guid
    List1.AddItem m_Guid.ToStr
    m_Indx = m_List.Count - 1
    UpdateView
End Sub

Private Sub CmbDecHex_Click()
    UpdateView
End Sub

Private Sub List1_Click()
    m_Indx = List1.ListIndex
    Set m_Guid = m_List.Item(m_Indx + 1)
    UpdateViewAllTBs
End Sub

Public Sub UpdateView()
    UpdateViewListBox
    UpdateViewAllTBs
End Sub
Private Sub UpdateViewListBox()
    'Dim i As Long: i = m_Indx
    'List1.List(m_Indx) = m_Guid.ToStr
    Dim i As Long: i = List1.ListIndex
    List1.Clear
    Dim g As Guid
    For Each g In m_List
        List1.AddItem g.ToStr
    Next
    'If i >= 0 Then List1.Selected(i) = True
    'List1.ListIndex = i
End Sub
Private Sub UpdateViewAllTBs()
    If CmbDecHex.Text = "Dec" Then
        With m_Guid
            TxtData1.Text = MUnsigned.UInt32_ToStr(.Data1)
            TxtData2.Text = MUnsigned.UInt16_ToStr(.Data2)
            TxtData3.Text = MUnsigned.UInt16_ToStr(.Data3)
            TxtData41.Text = .Data4(0)
            TxtData42.Text = .Data4(1)
            TxtData51.Text = .Data4(2)
            TxtData52.Text = .Data4(3)
            TxtData53.Text = .Data4(4)
            TxtData54.Text = .Data4(5)
            TxtData55.Text = .Data4(6)
            TxtData56.Text = .Data4(7)
        End With
    ElseIf CmbDecHex.Text = "Hex" Then
        With m_Guid
            TxtData1.Text = Hex(.Data1)
            TxtData2.Text = Hex(.Data2)
            TxtData3.Text = Hex(.Data3)
            TxtData41.Text = Hex(.Data4(0))
            TxtData42.Text = Hex(.Data4(1))
            TxtData51.Text = Hex(.Data4(2))
            TxtData52.Text = Hex(.Data4(3))
            TxtData53.Text = Hex(.Data4(4))
            TxtData54.Text = Hex(.Data4(5))
            TxtData55.Text = Hex(.Data4(6))
            TxtData56.Text = Hex(.Data4(7))
        End With
    End If
End Sub

' ##################### '   All TxtData_LostFocus    ' ##################### '
Private Sub TxtData1_LostFocus()
    Dim s As String: s = TxtData1.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Long
    If Not MUnsigned.UInt32_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data1 = v
    UpdateView
End Sub

Private Sub TxtData2_LostFocus()
    Dim s As String: s = TxtData2.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Integer
    If Not MUnsigned.UInt16_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data2 = v
    UpdateView
End Sub

Private Sub TxtData3_LostFocus()
    Dim s As String: s = TxtData3.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Integer
    If Not MUnsigned.UInt16_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data3 = v
    UpdateView
End Sub

Private Sub TxtData41_LostFocus()
    Dim s As String: s = TxtData41.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Integer
    If Not MUnsigned.UInt16_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data4(0) = CByte(v)
    UpdateView
End Sub

Private Sub TxtData42_LostFocus()
    Dim s As String: s = TxtData42.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Integer
    If Not MUnsigned.UInt16_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data4(1) = CByte(v)
    UpdateView
End Sub

Private Sub TxtData51_LostFocus()
    Dim s As String: s = TxtData51.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Integer
    If Not MUnsigned.UInt16_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data4(2) = CByte(v)
    UpdateView
End Sub

Private Sub TxtData52_LostFocus()
    Dim s As String: s = TxtData52.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Integer
    If Not MUnsigned.UInt16_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data4(3) = CByte(v)
    UpdateView
End Sub

Private Sub TxtData53_LostFocus()
    Dim s As String: s = TxtData53.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Integer
    If Not MUnsigned.UInt16_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data4(4) = CByte(v)
    UpdateView
End Sub

Private Sub TxtData54_LostFocus()
    Dim s As String: s = TxtData54.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Integer
    If Not MUnsigned.UInt16_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data4(5) = CByte(v)
    UpdateView
End Sub

Private Sub TxtData55_LostFocus()
    Dim s As String: s = TxtData55.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Integer
    If Not MUnsigned.UInt16_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data4(6) = CByte(v)
    UpdateView
End Sub

Private Sub TxtData56_LostFocus()
    Dim s As String: s = TxtData56.Text
    Dim r As Long:   r = IIf(CmbDecHex.ListIndex = 0, 10, 16)
    Dim v As Integer
    If Not MUnsigned.UInt16_TryParse(s, v, r) Then Exit Sub
    m_Guid.Data4(7) = CByte(v)
    UpdateView
End Sub


