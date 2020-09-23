VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " Scroll Keys"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   6000
      Left            =   8445
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   150
      Width           =   195
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   195
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6150
      Width           =   8295
   End
   Begin VB.PictureBox Picture1 
      Height          =   6260
      Left            =   120
      ScaleHeight     =   6195
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   120
      Width           =   8530
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   11520
         Left            =   0
         Picture         =   "Form1.frx":000C
         ScaleHeight     =   11520
         ScaleWidth      =   15360
         TabIndex        =   1
         Top             =   0
         Width           =   15360
         Begin VB.Label Label1 
            Height          =   315
            Left            =   8280
            TabIndex        =   4
            Top             =   6000
            Width           =   315
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Now with scrollbars inside picture frame

Option Explicit

Dim LblLeft As Long
Dim LblTop As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If VScroll1.Value - VScroll1.SmallChange < 0 Then
                VScroll1.Value = 0
                'This will prevent overscrolling (Error 380)
            Else
                VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
            End If
        Case vbKeyDown
            If VScroll1.Value + VScroll1.SmallChange > VScroll1.Max Then
                VScroll1.Value = VScroll1.Max
                'This will prevent overscrolling (Error 380)
            Else
                VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
            End If
        Case vbKeyLeft
            If HScroll1.Value - HScroll1.SmallChange < 0 Then
                HScroll1.Value = 0
                'This will prevent overscrolling (Error 380)
            Else
                HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
            End If
        Case vbKeyRight
                If HScroll1.Value + HScroll1.SmallChange > HScroll1.Max Then
                    HScroll1.Value = HScroll1.Max
                'This will prevent overscrolling (Error 380)
                Else
                    HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
                End If
    End Select
End Sub

Private Sub Form_Load()

    'VERY IMPORTANT
    'Form1.KeyPreview = True
    'Picture2.AutoSize = True

    'Set VScrollBar Max
    VScroll1.Max = Picture2.Height - Picture1.Height

    'Set VScrollBar LargeChange and SmallChange
    VScroll1.LargeChange = VScroll1.Max \ 10
    VScroll1.SmallChange = VScroll1.Max \ 50
    'You can set it to any value

    'Set HScrollBar Max
    HScroll1.Max = Picture2.Width - Picture1.Width

    'Set HScrollBar LargeChange and SmallChange
    HScroll1.LargeChange = HScroll1.Max \ 10
    HScroll1.SmallChange = HScroll1.Max \ 50
    'You can set it to any value

    'Save cover label position
    LblTop = Label1.Top
    LblLeft = Label1.Left
    
End Sub

Private Sub HScroll1_Change()
    Picture2.Left = -HScroll1.Value
    Label1.Left = LblLeft + HScroll1.Value
End Sub

Private Sub VScroll1_Change()
    Picture2.Top = -VScroll1.Value
    Label1.Top = LblTop + VScroll1.Value
End Sub
