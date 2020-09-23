VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ZZZZZZZOOOOOOOOOOMMMMMMMMMMM"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdZoom 
      Caption         =   "!!!ZOOM!!!"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.PictureBox picZoom 
      Height          =   8775
      Left            =   1440
      ScaleHeight     =   581
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   629
      TabIndex        =   1
      Top             =   120
      Width           =   9495
   End
   Begin VB.PictureBox picPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   0
      Top             =   120
      Width           =   1140
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdZoom_Click()
    Dim X    As Integer
    Dim Y    As Integer
    Dim Zoom As Integer

    Zoom = 6

    For X = 0 To picPicture.ScaleWidth - 1
        For Y = 0 To picPicture.ScaleHeight - 1
            picZoom.Line ((X * Zoom), (Y * Zoom))-(((X * Zoom) + Zoom), ((Y * Zoom) + Zoom)), _
                picPicture.Point(X, Y), BF
        Next
    Next
End Sub

