VERSION 5.00
Begin VB.Form FormUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penghitung Karakter"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4125
   Icon            =   "FormUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMD_Reset 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton CMD_About 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton CMD_Keluar 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox TextKarakter 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FormUtama.frx":000C
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label LabelKarakter 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   480
   End
End
Attribute VB_Name = "FormUTama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pesan As Integer

Private Sub CMD_About_Click()
    MsgBox "Penghitung Jumlah Karakter by Rizky Hafitsyah", vbInformation + vbOKOnly, "About"
    TextKarakter.SetFocus
End Sub

Private Sub CMD_Keluar_Click()
    Pesan = MsgBox("Anda yakin ingin keluar dari program?", vbQuestion + vbYesNo, "Keluar")
        If Pesan = vbYes Then
            End
        ElseIf Pesan = vbNo Then
            TextKarakter.SetFocus
        End If
End Sub

Private Sub CMD_Reset_Click()
    With TextKarakter
        .Text = ""
        .SetFocus
    End With
End Sub

Private Sub Form_Load()
    With Me
        .TextKarakter.Text = ""
        .TextKarakter.MaxLength = 10000
        .CMD_About.Caption = "Tentang"
        .CMD_Keluar.Caption = "Keluar"
        .LabelKarakter.Caption = "0 Karakter"
        .CMD_Reset.Caption = "Reset"
    End With
End Sub

Private Sub TextKarakter_Change()
    LabelKarakter.Caption = Len(TextKarakter.Text) & " Karakter"
    If TextKarakter.Text = "" Then
        CMD_Reset.Enabled = False
    Else
        CMD_Reset.Enabled = True
    End If
End Sub

