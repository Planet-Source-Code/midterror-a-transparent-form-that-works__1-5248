VERSION 5.00
Begin VB.Form frmTrans 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   230
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4605
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Text            =   "A TextBox!"
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A Button"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   4560
      Shape           =   2  'Oval
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   255
   End
   Begin VB.Line Line6 
      X1              =   264
      X2              =   216
      Y1              =   104
      Y2              =   80
   End
   Begin VB.Line Line5 
      X1              =   264
      X2              =   312
      Y1              =   104
      Y2              =   80
   End
   Begin VB.Line Line4 
      X1              =   232
      X2              =   264
      Y1              =   216
      Y2              =   136
   End
   Begin VB.Line Line3 
      X1              =   264
      X2              =   264
      Y1              =   72
      Y2              =   136
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Move Me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   304
      X2              =   264
      Y1              =   216
      Y2              =   136
   End
   Begin VB.Label Label1 
      Caption         =   "A Label!"
      Height          =   555
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2145
   End
End
Attribute VB_Name = "frmTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Move form without a border declarations
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112




Private Sub Command2_Click()
    'unloads the form
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = "This is just a sample of what this can do. Make sure to have a borderless form."
    MakeTransparent frmTrans
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This code makes the form move when the mouse
    'is down on the label
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_SYSCOMMAND, &HF012, 0
    End If
End Sub

