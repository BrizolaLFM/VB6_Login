VERSION 5.00
Begin VB.Form TelaCadastro 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4185
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
   ScaleHeight     =   6045
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtUser2 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtPassword2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "•"
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton cmdCadastrar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cadastre-se"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblVoltar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Voltar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   8
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblBack 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   60
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label lblUsuário2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label lblSenha2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   3360
      Width           =   450
   End
   Begin VB.Label lblTitleCadastro 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tela de Cadastro"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "TelaCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCadastrar_Click()
    
    If txtUser2.Text = "" Then
    MsgBox "Favor preencher todos os campos!", vbExclamation
    txtUser2.SetFocus
    Exit Sub
    End If
    
    
    If txtEmail.Text = "" Then
    MsgBox "Favor preencher todos os campos!", vbExclamation
    txtEmail.SetFocus
    Exit Sub
    End If
    
    
    If txtPassword2.Text = "" Then
    MsgBox "Favor preencher todos os campos!", vbExclamation
    txtPassword2.SetFocus
    Exit Sub
    End If


On Error Resume Next

Set conn = New ADODB.Connection
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\luizf\OneDrive\Documentos\Projetos_VB6\BDcadastroVB.mdb"

Dim rs As New ADODB.Recordset
Dim sql As String


    If Err.Number <> 0 Then
    MsgBox "Luiz erro ao conectar ao banco de dados: " & Err.Description
    
    Else
    Dim usuario As String
    Dim email As String
    Dim senha As String
    
    usuario = txtUser2.Text
    email = txtEmail.Text
    senha = txtPassword2.Text
    
    sql = "INSERT INTO Users (usuario, email, senha) VALUES ('" & usuario & "', '" & email & "', '" & senha & "')"
    
    conn.Execute sql
    
    rs.Close
    conn.Close
    
    MsgBox "Cadastro realizado com sucesso!", vbInformation
    End If

    End Sub

Private Sub lblVoltar_Click()
TelaCadastro.Hide
TelaLogin.Show
End Sub

Private Sub txtUser2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtEmail.SetFocus
    End If
    Exit Sub
End Sub


Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtPassword2.SetFocus
    End If
    
    If KeyAscii = 27 Then
    txtUser2.SetFocus
    End If
End Sub

Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    txtEmail.SetFocus
    End If
    
    If KeyAscii = 13 Then
    cmdCadastrar.SetFocus
    End If
End Sub
