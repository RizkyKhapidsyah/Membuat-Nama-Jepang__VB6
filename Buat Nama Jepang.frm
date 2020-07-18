VERSION 5.00
Begin VB.Form Utama 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nama Jepang"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800080&
   Icon            =   "Buat Nama Jepang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Keluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Hasil 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Baca 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1440
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Top             =   270
      Width           =   3615
   End
   Begin VB.CommandButton Lakukan 
      Caption         =   "Buat Nama"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label NamaJepang 
      Caption         =   "Nama Jepang:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label NamaAnda 
      Caption         =   "Nama anda:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
    KontrolXP 'Meload fungsi pada modul
End Sub

Private Sub Lakukan_Click()
    Dim ambil As Integer
    Dim Huruf As String
    Dim nama As String
    '^deklarasi variabel

    For ambil = 1 To Len(Baca)
        Huruf = Asc(Mid(Baca, ambil, 1))
            Select Case Huruf
                Case 32:
                    nama = nama & " "
                Case 39:
                    nama = nama & "'"
                Case 45:
                    nama = nama & "-"
                Case 65:
                    nama = nama & "ka"
                Case 66:
                    nama = nama & "zu"
                Case 67:
                    nama = nama & "mi"
                Case 68:
                    nama = nama & "te"
                Case 69:
                    nama = nama & "ku"
                Case 70:
                    nama = nama & "lu"
                Case 71:
                    nama = nama & "ji"
                Case 72:
                    nama = nama & "ri"
                Case 73:
                    nama = nama & "ki"
                Case 74:
                    nama = nama & "zu"
                Case 75:
                    nama = nama & "me"
                Case 76:
                    nama = nama & "ta"
                Case 77:
                    nama = nama & "rin"
                Case 78:
                    nama = nama & "to"
                Case 79:
                    nama = nama & "mo"
                Case 80:
                    nama = nama & "no"
                Case 81:
                    nama = nama & "ke"
                Case 82:
                    nama = nama & "shi"
                Case 83:
                    nama = nama & "ari"
                Case 84:
                    nama = nama & "chi"
                Case 85:
                    nama = nama & "do"
                Case 86:
                    nama = nama & "ru"
                Case 87:
                    nama = nama & "ko"
                Case 88:
                    nama = nama & "na"
                Case 89:
                    nama = nama & "su"
                Case 90:
                    nama = nama & "ro"
                Case 97:
                    nama = nama & "ka"
                Case 98:
                    nama = nama & "zu"
                Case 99:
                    nama = nama & "mi"
                Case 100:
                    nama = nama & "te"
                Case 101:
                    nama = nama & "ku"
                Case 102:
                    nama = nama & "lu"
                Case 103:
                    nama = nama & "ji"
                Case 104:
                    nama = nama & "ri"
                Case 105:
                    nama = nama & "ki"
                Case 106:
                    nama = nama & "zu"
                Case 107:
                    nama = nama & "me"
                Case 108:
                    nama = nama & "ta"
                Case 109:
                    nama = nama & "rin"
                Case 110:
                    nama = nama & "to"
                Case 111:
                    nama = nama & "mo"
                Case 112:
                    nama = nama & "no"
                Case 113:
                    nama = nama & "ke"
                Case 114:
                    nama = nama & "shi"
                Case 115:
                    nama = nama & "ari"
                Case 116:
                    nama = nama & "chi"
                Case 117:
                    nama = nama & "do"
                Case 118:
                    nama = nama & "ru"
                Case 119:
                    nama = nama & "ko"
                Case 120:
                    nama = nama & "na"
                Case 121:
                    nama = nama & "su"
                Case 122:
                    nama = nama & "ro"
                Case Else:
                    MsgBox "Input hanya huruf", vbExclamation, "Input salah"
            End Select
        Hasil.Text = StrConv(nama, 3)
    Next
End Sub

Private Sub Keluar_Click()
    End
End Sub



Private Sub Link_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Link.ForeColor = vbBlue
End Sub

Private Sub Link_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Link.ForeColor = &H800080
End Sub
