VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "電腦硬體裝修乙級第一站第1題"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   20.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   9060
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   3120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   2400
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   4320
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Red LED"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   4800
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Green LED"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   360
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Shape G 
      Height          =   375
      Index           =   7
      Left            =   240
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape G 
      Height          =   375
      Index           =   6
      Left            =   768
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape G 
      Height          =   375
      Index           =   5
      Left            =   1296
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape G 
      Height          =   375
      Index           =   4
      Left            =   1824
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape G 
      Height          =   375
      Index           =   3
      Left            =   2352
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape G 
      Height          =   375
      Index           =   2
      Left            =   2880
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape G 
      Height          =   375
      Index           =   1
      Left            =   3408
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape G 
      Height          =   375
      Index           =   0
      Left            =   3936
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape R 
      Height          =   375
      Index           =   7
      Left            =   4464
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape R 
      Height          =   375
      Index           =   6
      Left            =   4992
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape R 
      Height          =   375
      Index           =   5
      Left            =   5520
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape R 
      Height          =   375
      Index           =   4
      Left            =   6048
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape R 
      Height          =   375
      Index           =   3
      Left            =   6576
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape R 
      Height          =   375
      Index           =   2
      Left            =   7104
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape R 
      Height          =   375
      Index           =   1
      Left            =   7632
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape R 
      Height          =   375
      Index           =   0
      Left            =   8160
      Shape           =   3  '圓形
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b(99), c As Integer


Private Sub Command1_Click(Index As Integer)
a = Index
c = 0
End Sub

Private Sub display(no)
For i = 0 To 7
    If no Mod 2 = 1 And a = 1 Then G(i).FillColor = RGB(0, 255, 0)
    If no Mod 2 = 1 And a = 2 Then R(i).FillColor = RGB(255, 0, 0)
    no = no \ 2
Next i
End Sub

Private Sub Timer1_Timer()
b(0) = 1
b(1) = 2
b(2) = 4
b(3) = 8
b(4) = &H10
b(5) = &H20
b(6) = &H40
b(7) = &H80

Label1.Caption = vbCrLf & "Current Time:" & Time$

For i = 0 To 7
    G(i).FillStyle = 1
    R(i).FillStyle = 1
Next i

If OpenUsbDevice(VendorID, ProductID) Then
    For i = 0 To 7
        G(i).FillStyle = 0: G(i).FillColor = RGB(0, 128, 0)
        R(i).FillStyle = 0: R(i).FillColor = RGB(128, 0, 0)
    Next i
    OutDataCtrl 0, 0
    OutDataCtrl 0, 16
    
    If a = 1 Then OutDataCtrl b(c), 0: display (b(c))
    If a = 2 And c <= 7 Then
        OutDataCtrl 2 ^ c, 32
        OutDataCtrl 2 ^ c, 48
        display (2 ^ c)
    End If
End If

If a = 3 Then CloseUsbDevice: End
If c > 15 Then c = 15 Else c = c + 1





End Sub
