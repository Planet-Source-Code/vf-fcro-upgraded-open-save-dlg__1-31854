VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "AUTHOR:Vanja Fuckar,Zagreb,Croatia   EMAIL:INGA@VIP.HR"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   5775
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   5775
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   960
         Left            =   0
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   900
         ScaleWidth      =   1125
         TabIndex        =   3
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "New Look For a Standard OPEN/SAVE Dialog"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "What ya' think???"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   4335
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "GetSaveFileName"
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "GetOpenFileName"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const sFilter = "Samo EXE (*.exe)" & vbNullChar & "*.exe"
'Dim sFilter As String 'filter fileova
Dim sFile As String 'ime filea
Dim sPath As String 'ime filea+PATH

Private Sub Command1_Click()
aa = GetOpenFilePath(hwnd, sFilter, 0, sFile, "", "Uèitaj File", sPath)
If aa = False Then Exit Sub
End Sub

Private Sub Command2_Click()
aa = GetSaveFilePath(hwnd, sFilter, 0, sFilter, "", "", "Snimi File", sPath)
If aa = False Then Exit Sub
End Sub


Private Sub Form_Load()
InsertCtrl Picture1.hwnd, hwnd, 10, 235
SetWidth = 330
End Sub

Private Sub Picture1_Click()
If MsgBox("Želite li zatvoriti dialog?" & vbCrLf & "Would you like to close a Dialogbox?", vbOKCancel, "Potvrda/Confirm") = vbOK Then
CloseDLG
End If
End Sub
