VERSION 5.00
Begin VB.Form TelaLogin 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Tela de Login"
   ClientHeight    =   4725
   ClientLeft      =   8445
   ClientTop       =   6345
   ClientWidth     =   4260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H008080FF&
      Caption         =   "Sair"
      Height          =   480
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Login"
      Height          =   480
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "•"
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblCriarConta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Criar conta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   2400
      TabIndex        =   8
      Top             =   4200
      Width           =   795
   End
   Begin VB.Label lblTxtPosuir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Não possui uma conta?"
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Top             =   4200
      Width           =   1650
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Tela de Login"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label lblUsuário 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   540
   End
End
Attribute VB_Name = "TelaLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const user = "Luiz"
Const password = "1234"

Private Sub cmdLogin_Click()
    If txtUser.Text = "" Then
    MsgBox "Favor preencher o campo usuário!", vbExclamation
    txtUser.SetFocus
    Exit Sub
    End If
    
    If txtPassword.Text = "" Then
    MsgBox "Favor preencher o campo Senha!", vbExclamation
    txtPassword.SetFocus
    Exit Sub
    End If
    
    If txtUser.Text <> user Then
    MsgBox "Usuário e/ou Senha incorretos!", vbExclamation
    Exit Sub
    End If
    
    If txtPassword.Text <> password Then
    MsgBox "Usuário e/ou Senha incorretos!", vbExclamation
    Exit Sub
    End If
 
 
    If txtUser.Text = user And txtPassword.Text = password Then
    frmHello.Show
    Unload Me
    End If
    
End Sub

Private Sub cmdSair_Click()
    End
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    txtUser.SetFocus
    End If
    
    If KeyAscii = 13 Then
    cmdLogin.SetFocus
    End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtPassword.SetFocus
    End If
End Sub

Private Sub lblCriarConta_Click()
    TelaCadastro.Show
    Unload Me
End Sub

