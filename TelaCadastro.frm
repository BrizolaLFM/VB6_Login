VERSION 5.00
Begin VB.Form frmHello 
   Caption         =   "Hello"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7020
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
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblBemVindo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Bem vindo"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1635
      TabIndex        =   0
      Top             =   1320
      Width           =   1560
   End
End
Attribute VB_Name = "frmHello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
