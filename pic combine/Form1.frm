VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Picture Combine"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   495
      Left            =   5400
      TabIndex        =   14
      Top             =   6600
      Width           =   1695
   End
   Begin ComctlLib.Slider slide 
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   5160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   327682
      LargeChange     =   1
      Max             =   255
      SelStart        =   128
      Value           =   128
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3435
      ScaleWidth      =   3675
      TabIndex        =   10
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4560
      Top             =   6000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fade It"
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   6120
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cc 
      Left            =   4320
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Combine Now!"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   3495
      Left            =   4320
      ScaleHeight     =   3435
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   240
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3435
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "128"
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Alpha Value:"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture 2"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture 1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare Api
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Dim pic1 As String
Dim pic2 As String
Private Sub Command1_Click()
'Open dialog box for picture 1
cc.Filter = "BITMAP|*.bmp|JPEG|*.jpg|GIF|*.gif"
cc.ShowOpen
Picture1.Picture = LoadPicture(cc.FileName)
pic1 = cc.FileName
End Sub

Private Sub Command2_Click()
'Open dialog box for picture 2
cc.Filter = "BITMAP|*.bmp|JPEG|*.jpg|GIF|*.gif"
cc.ShowOpen
Picture2.Picture = LoadPicture(cc.FileName)
pic2 = cc.FileName
End Sub

Private Sub Command3_Click()
    Dim blend As BLENDFUNCTION, blendl As Long
    'pixels
    Picture1.ScaleMode = vbPixels
    Picture2.ScaleMode = vbPixels
    Picture3.ScaleMode = vbPixels
    'parameters
    With blend
        .BlendOp = &H0
        .BlendFlags = 0
        .SourceConstantAlpha = slide.Value
        .AlphaFormat = 0
    End With
    'set the structure to a long
    RtlMoveMemory blendl, blend, 4
    'Copy the picture to picture3 so that it will combine in the result
    Picture3.Picture = Picture1.Picture
    'Combine picture from picture1 over picture3
    AlphaBlend Picture3.hdc, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, blendl
    Picture3.Refresh
End Sub

Private Sub Command4_Click()
    Dim blend As BLENDFUNCTION, blendl As Long
    'pixels
    Picture1.ScaleMode = vbPixels
    Picture2.ScaleMode = vbPixels
    Picture3.ScaleMode = vbPixels
    'parameters
    'Let it to fade from alpha 1 to 255
    'Change the value for alpha transparent
    For i = slide.Value To 255
    slide.Value = i
    With blend
        .BlendOp = &H0
        .BlendFlags = 0
        .SourceConstantAlpha = i
        .AlphaFormat = 0
    End With
    'set the structure to a long
    RtlMoveMemory blendl, blend, 4
    'Copy the picture to picture3 so that it will combine in the result
    Picture3.Picture = Picture1.Picture
    'Combine picture from picture1 over picture3
    AlphaBlend Picture3.hdc, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, blendl
    Picture3.Refresh
    DoEvents
    Label5.Caption = slide.Value
    Next i

End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
MsgBox "Create by Alvin Chua, 15/8/2004"
End Sub

Private Sub slide_Scroll()
'Value for alpha blend
Label5.Caption = slide.Value
End Sub
