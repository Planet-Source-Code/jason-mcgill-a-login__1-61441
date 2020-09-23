VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register"
      Height          =   350
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   2760
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "If you do not have a username and password type them in the spaces below and click register."
      Height          =   735
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmlogin.frx":0442
      Stretch         =   -1  'True
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Plese enter your username and password in the space provided below to login."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If you have any questions about how to use this program or ideas on how to
'better this program please contact me (Jason) at jaymcgill@gmail.com.
Option Explicit
Dim cn As New ADODB.Connection, strCNString As String
Dim rs As New ADODB.Recordset
Dim Txt As String

Private Sub cmdOK_Click()

On Error GoTo ErrHandler
'Connect to database
strCNString = "Data Source=" & App.Path & "\dbpassword.mdb"
cn.Provider = "Microsoft Jet 4.0 OLE DB Provider"
cn.ConnectionString = strCNString
cn.Properties("Jet OLEDB:Database Password") = "jason"
cn.Open
'Open recordsource
With rs
   
         .Open "Select * from tblUsers where Username='" & txtname.Text & "' and Password='" & txtpass.Text & "'", cn, adOpenDynamic, adLockOptimistic
        'Check username and password
        If .EOF Then
            MsgBox "Access Denied...Please enter correct password!", vbOKOnly + vbCritical, "Security Login"
               txtname.Text = ""
               txtpass.Text = ""
               txtname.SetFocus
               cn.Close
        Else
           Txt = "" & " " & UCase$(txtname.Text) & ""
            MsgBox "Welcome!!!" & Txt, vbOKOnly + vbExclamation, "Security Login"
            cn.Close
            Unload Me
            Main.Show
            
        End If
    End With

     Exit Sub
     
ErrHandler:
MsgBox Err.Description, vbCritical, "Login"
cn.Close
End Sub

Private Sub cmdCancel_Click()
'Close form
Unload Me
End Sub

Private Sub cmdRegister_Click()
'Register a new username and password
On Error Resume Next
'Keep user from saving a blank username
If txtname.Text = "" Then GoTo message
'Connect to database
strCNString = "Data Source=" & App.Path & "\dbpassword.mdb"
cn.Provider = "Microsoft Jet 4.0 OLE DB Provider"
cn.ConnectionString = strCNString
cn.Properties("Jet OLEDB:Database Password") = "jason"
cn.Open
'Open recordsource
rs.Open "Select * from tblUsers", cn, adOpenDynamic, adLockOptimistic
'Ready recordsource for adding username and password
rs.AddNew
'Assign test boxes on form to their appropriate field in the recordsource
rs(0) = txtname.Text
rs(1) = txtpass.Text
'Save record
rs.Save
'Close connections to database
cn.Close
rs.Close
MsgBox "User Name and Password Created.", vbInformation, "Confirmation"
Exit Sub
message:
    MsgBox "You must enter a User Name and Password.", vbCritical, "Error"
End Sub

Private Sub Form_Load()
'Disable Register button until data is entered into both text boxes
cmdRegister.Enabled = False
End Sub

Private Sub txtpass_Change()
'Enable Register button to allow user to save record
cmdRegister.Enabled = True
End Sub
