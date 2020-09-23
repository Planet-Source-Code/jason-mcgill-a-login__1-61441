VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Test Page"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4500
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogout 
      Caption         =   "&Logout"
      Default         =   -1  'True
      Height          =   1935
      Left            =   3240
      Picture         =   "frmmain.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Test Page.  Replace with your program."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   9015
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogout_Click()
'Close form
Dim strValue As String
strValue = MsgBox("Are you sure you want to logout?", vbQuestion + vbYesNo, "Logout?")
If strValue = vbYes Then Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Open login form
frmlogin.Show
End Sub

