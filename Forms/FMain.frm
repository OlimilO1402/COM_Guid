VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "GUID/CLSID/UUID"
   ClientHeight    =   3975
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnSomeTests 
      Caption         =   "IsEqual?"
      Height          =   375
      Left            =   6360
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnTestVBGuid 
      Caption         =   "Test VBGuid"
      Height          =   615
      Left            =   5280
      TabIndex        =   23
      Top             =   0
      Width           =   1095
   End
   Begin VB.ComboBox CmbDecHex 
      Height          =   375
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
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnCreateCLSID 
      Caption         =   "Create CLSID"
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnCreateGUID 
      Caption         =   "Create GUID"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "GUID / CLSID / UUID: 1"
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
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   360
         Width           =   510
      End
      Begin VB.Label LblData4 
         AutoSize        =   -1  'True
         Caption         =   "Data4"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   360
         Width           =   510
      End
      Begin VB.Label LblData3 
         AutoSize        =   -1  'True
         Caption         =   "Data3"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   510
      End
      Begin VB.Label LblData2 
         AutoSize        =   -1  'True
         Caption         =   "Data2"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   510
      End
      Begin VB.Label LblData1 
         AutoSize        =   -1  'True
         Caption         =   "Data1"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   510
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
      MultiSelect     =   2  'Erweitert
      TabIndex        =   0
      Top             =   1800
      Width           =   7455
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditClone 
         Caption         =   "Clone"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
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

Private Sub BtnSomeTests_Click()
    'Testing the functions Clone, IsEQual and IsSame
    Dim g As Guid, g1 As Guid, g2 As Guid
    
    Dim i As Long
    For i = 1 To m_List.Count
        Set g = m_List.Item(i)
        If List1.Selected(i - 1) Then
            If g1 Is Nothing Then
                Set g1 = g
            Else
                Set g2 = g
                Exit For
            End If
        End If
    Next
    
    If g1 Is Nothing Then
        MsgBox "Please select 2 guids first"
        Exit Sub
    End If
    If g2 Is Nothing Then
        Set g2 = m_Guid
    End If
    MsgBox CheckSameOrEQual(g1, g2)
    
End Sub
Private Function CheckSameOrEQual(g1 As Guid, g2 As Guid) As String
    Dim s As String
    If g1.IsEqual(g2) Then
        s = "The data of the two Guids is equal:"
        If g1.IsSame(g2) Then
            s = "The two Guids are variables of the same Object:"
        End If
    Else
        s = "The data of the two Guids is NOT equal:"
    End If
    CheckSameOrEQual = s & vbCrLf & g1.ToStr & vbCrLf & g2.ToStr
End Function

Private Sub BtnTestVBGuid_Click()
    MVBGuid.Init
    MsgBox "IID_IUnknown.Equals( _" & vbCrLf & VBGuid_ToStr(MVBGuid.IID_IUnknown) & ", _" & vbCrLf & sIID_IUnknown & ") = " & MVBGuid.VBGuid_Equals(MVBGuid.IID_IUnknown, MVBGuid.New_VBGuidS(sIID_IUnknown))
    MsgBox "IID_IDispatch.Equals( _" & vbCrLf & VBGuid_ToStr(MVBGuid.IID_IDispatch) & ", _" & vbCrLf & sIID_IDispatch & ") = " & MVBGuid.VBGuid_Equals(MVBGuid.IID_IDispatch, MVBGuid.New_VBGuidS(sIID_IDispatch))
    MsgBox "IID_IEnumVariant.Equals( _" & vbCrLf & VBGuid_ToStr(MVBGuid.IID_IEnumVariant) & ", _" & vbCrLf & sIID_IEnumVariant & ") = " & MVBGuid.VBGuid_Equals(MVBGuid.IID_IEnumVariant, MVBGuid.New_VBGuidS(sIID_IEnumVariant))
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set m_List = New Collection
    BtnCreateGUID_Click
End Sub

Private Sub Form_Resize()
    Dim l As Single: l = List1.Left
    Dim T As Single: T = List1.Top
    Dim W As Single: W = Me.ScaleWidth - 2 * l
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move l, T, W, H
End Sub

Private Sub BtnCreateGUID_Click()
    Set m_Guid = MNew.GuidCo
    m_List.Add m_Guid
    m_Indx = m_List.Count - 1
    List1.AddItem m_Guid.ToStr
    UpdateViewAllTBs
End Sub

Private Sub BtnCreateCLSID_Click()
    Set m_Guid = MNew.GuidCo
    m_List.Add m_Guid
    m_Indx = m_List.Count - 1
    List1.AddItem m_Guid.ToStr
    UpdateViewAllTBs
End Sub

Private Sub BtnCreateUUID_Click()
    Set m_Guid = MNew.UUID
    m_List.Add m_Guid
    m_Indx = m_List.Count - 1
    List1.AddItem m_Guid.ToStr
    UpdateViewAllTBs
End Sub

Private Sub CmbDecHex_Click()
    UpdateView
End Sub

Private Sub List1_Click()
    m_Indx = List1.ListIndex
    If m_Indx < 0 Then Exit Sub
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
            TxtData41.Text = .Data5(0)
            TxtData42.Text = .Data5(1)
            TxtData51.Text = .Data5(2)
            TxtData52.Text = .Data5(3)
            TxtData53.Text = .Data5(4)
            TxtData54.Text = .Data5(5)
            TxtData55.Text = .Data5(6)
            TxtData56.Text = .Data5(7)
        End With
    ElseIf CmbDecHex.Text = "Hex" Then
        With m_Guid
            TxtData1.Text = Hex(.Data1)
            TxtData2.Text = Hex(.Data2)
            TxtData3.Text = Hex(.Data3)
            TxtData41.Text = Hex(.Data5(0))
            TxtData42.Text = Hex(.Data5(1))
            TxtData51.Text = Hex(.Data5(2))
            TxtData52.Text = Hex(.Data5(3))
            TxtData53.Text = Hex(.Data5(4))
            TxtData54.Text = Hex(.Data5(5))
            TxtData55.Text = Hex(.Data5(6))
            TxtData56.Text = Hex(.Data5(7))
        End With
    End If
End Sub

Private Sub mnuEditClone_Click()
    Dim g As Guid: Set g = m_Guid.Clone
    m_List.Add g
    UpdateView
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.Clear
    Clipboard.SetText m_Guid.ToStr
End Sub

Private Sub mnuEditPaste_Click()
    Dim s As String
    If Not Clipboard.GetFormat(ClipBoardConstants.vbCFText) Then
        MsgBox "Only paste string"
        Exit Sub
    End If
    s = Clipboard.GetText
    Set m_Guid = MNew.GuidS(s)
    m_List.Add m_Guid
    UpdateView
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


