VERSION 5.00
Begin VB.Form frmHello 
   BackColor       =   &H00FF8080&
   Caption         =   "Wellcome"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3840
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblOla 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seja Bem-vindo!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1770
      TabIndex        =   0
      Top             =   1560
      Width           =   5535
   End
End
Attribute VB_Name = "frmHello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
End
End Sub
