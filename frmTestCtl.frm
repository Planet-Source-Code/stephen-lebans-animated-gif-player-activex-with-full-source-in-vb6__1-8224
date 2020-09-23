VERSION 5.00
Object = "{5F3A0583-2C54-11D4-9E31-00902715CDA7}#5.0#0"; "AnimatedGif.ocx"
Begin VB.Form frmTestCtl 
   Caption         =   "Animated Gif ActiveX Control"
   ClientHeight    =   5904
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   8880
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5904
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin AnimatedGif.AnimatedGifCtl AnimatedGifCtl2 
      Height          =   2004
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   6972
      _ExtentX        =   12298
      _ExtentY        =   3535
      BorderStyle     =   1
   End
   Begin AnimatedGif.AnimatedGifCtl AnimatedGifCtl1 
      Height          =   2004
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   6972
      _ExtentX        =   12298
      _ExtentY        =   3535
      BorderStyle     =   1
   End
   Begin VB.CommandButton CmdColor2 
      Caption         =   "Background Color"
      Height          =   372
      Left            =   7080
      TabIndex        =   0
      Top             =   3720
      Width           =   1692
   End
   Begin VB.CommandButton CmdColor 
      Caption         =   "Background Color"
      Height          =   372
      Left            =   7080
      TabIndex        =   3
      Top             =   1560
      Width           =   1692
   End
   Begin VB.CommandButton CmdLoad2 
      Caption         =   "Load Animated Gif"
      Height          =   372
      Left            =   7080
      TabIndex        =   7
      Top             =   2400
      Width           =   1692
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Load Animated Gif"
      Height          =   372
      Left            =   7080
      TabIndex        =   6
      Top             =   240
      Width           =   1692
   End
   Begin VB.CommandButton CmdPlay2 
      Caption         =   "Play Animated Gif"
      Height          =   372
      Left            =   7080
      TabIndex        =   5
      Top             =   2844
      Width           =   1692
   End
   Begin VB.CommandButton CmdStop2 
      Caption         =   "Stop"
      Height          =   372
      Left            =   7080
      TabIndex        =   4
      Top             =   3276
      Width           =   1692
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Stop"
      Height          =   372
      Left            =   7080
      TabIndex        =   2
      Top             =   1116
      Width           =   1692
   End
   Begin VB.CommandButton CmdPlay 
      Caption         =   "Play Animated Gif"
      Height          =   372
      Left            =   7080
      TabIndex        =   1
      Top             =   684
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1452
      Left            =   0
      TabIndex        =   8
      Top             =   4440
      Width           =   8892
   End
End
Attribute VB_Name = "frmTestCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdColor_Click()
Me.AnimatedGifCtl1.StopGif
Me.AnimatedGifCtl1.setbackColor
Me.AnimatedGifCtl1.LoadGif False
' If you want the  Gif to start
' automatically then
' uncomment out the line below.
Me.AnimatedGifCtl1.StartTimer
End Sub

Private Sub CmdLoad_Click()
Me.AnimatedGifCtl1.LoadGif True
' If you want the  Gif to start
' automatically upon load then
' uncomment out the line below.
'Me.AnimatedGifCtl1.StartTimer
DoEvents
Me.Refresh
End Sub

Private Sub CmdPlay_Click()
Me.AnimatedGifCtl1.StartTimer
End Sub

Private Sub CmdStop_Click()
Me.AnimatedGifCtl1.StopGif
End Sub

Private Sub CmdLoad2_Click()
' Load for the second control
Me.AnimatedGifCtl2.LoadGif True
' If you want the  Gif to start
' automatically upon load then
' uncomment out the line below.
'Me.AnimatedGifCtl2.StartTimer
DoEvents
Me.Refresh
End Sub

Private Sub CmdPlay2_Click()
' Play for the second control
Me.AnimatedGifCtl2.StartTimer
End Sub

Private Sub CmdStop2_Click()
' Stop for the second control
Me.AnimatedGifCtl2.StopGif
End Sub

Private Sub CmdColor2_Click()
' Set BackColor for second control
Me.AnimatedGifCtl2.StopGif
Me.AnimatedGifCtl2.setbackColor
Me.AnimatedGifCtl2.LoadGif False
' If you want the  Gif to start
' automatically then
' uncomment out the line below.
Me.AnimatedGifCtl2.StartTimer

End Sub


Private Sub Form_Load()
Dim StrTemp As String
StrTemp = "Copyright Lebans Holdings 1999 Ltd. " '& vbCrLf
StrTemp = StrTemp & "You may freely use and redistribute this code providing the one line copyright notice is left intact. " '& vbCrLf
StrTemp = StrTemp & "You may not resell this code as part of a collection or by itself. Please feel free to use this code within your own applications, " ' & vbCrLf
StrTemp = StrTemp & "whether private or commercial,  without cost or obligation. " & vbCrLf
StrTemp = StrTemp & "Contact   Stephen@lebans.com    or visit    www.lebans.com"
Label1.Caption = StrTemp

End Sub

