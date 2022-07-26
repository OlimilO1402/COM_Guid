VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9540
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   6
      Top             =   600
      Width           =   7335
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   9
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
         TabIndex        =   7
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
         TabIndex        =   17
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
         TabIndex        =   14
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
         TabIndex        =   12
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
         TabIndex        =   10
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
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   5655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   6000
      TabIndex        =   1
      Top             =   2400
      Width           =   480
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Guid As Guid

Private Sub Form_Load()
    
    'CmbDecHex.Text = "Hex"
    BtnCreateGUID_Click
    
    'Set m_Guid = MNew.GuidCo
'
'    Me.Caption = "Testing Class Guid " & App.Major & "." & App.Minor & "." & App.Revision
'    With List1
'        .Clear
'        .AddItem """{64440490-4c8b-11d1-8b70-080036b11a03}"""
'        .AddItem "{64440490-4C8B-11D1-8B70-080036B11A03}"
'        '64440490-4C8B-11D1-8B70-080036B11A03
'        .AddItem "0x64440490, 0x4C8B, 0x11D1, 0x8B, 0x70, 0x08, 0x00, 0x36, 0xB1, 0x1A, 0x03"
'        '.AddItem ""
'        '.AddItem ""
'    End With
End Sub

Private Sub BtnCreateGUID_Click()
    Set m_Guid = MNew.GuidCo
    UpdateView
End Sub

Private Sub BtnCreateCLSID_Click()
    Set m_Guid = MNew.GuidCo
    UpdateView
End Sub

Private Sub BtnCreateUUID_Click()
    Set m_Guid = MNew.UUID
    UpdateView
End Sub

Private Sub CmbDecHex_Click()
    UpdateView
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

Public Sub UpdateView()
    If CmbDecHex.Text = "Dec" Then
        With m_Guid
            TxtData1.Text = .Data1
            TxtData2.Text = .Data2
            TxtData3.Text = .Data3
            TxtData41.Text = .Data4(0)
            TxtData42.Text = .Data4(1)
            TxtData51.Text = .Data4(2)
            TxtData52.Text = .Data4(3)
            TxtData53.Text = .Data4(4)
            TxtData54.Text = .Data4(5)
            TxtData55.Text = .Data4(6)
            TxtData56.Text = .Data4(7)
        End With
    Else
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
