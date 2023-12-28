VERSION 5.00
Begin VB.Form FormLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle Financeiro - Versão 1.0.0"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEntrar 
      Appearance      =   0  'Flat
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   2745
      Width           =   1470
   End
   Begin VB.TextBox txtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2310
      Width           =   2370
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   1890
      Width           =   2370
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1185
      TabIndex        =   1
      Top             =   2325
      Width           =   555
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1065
      TabIndex        =   0
      Top             =   1905
      Width           =   675
   End
   Begin VB.Image ImageLogo 
      Height          =   1500
      Left            =   2160
      Picture         =   "FormLogin.frx":0000
      Top             =   285
      Width           =   1500
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEntrar_Click()
100     Dim conectaBD As clsConexaoBanco
101     Set conectaBD = New clsConexaoBanco
    
102     If VerificaPreenchimentoCampos(FormLogin.Name, "entrar") = True Then
    
103         If conectaBD.conectaBancoDados(txtUsuario.Text, txtSenha.Text) = False Then
104             Set conectaBD = Nothing
105             Exit Sub
106         End If
107         Set conectaBD = Nothing

108         Unload Me
109         FormMenu.Show 1

110     Else
        
111         MsgBox "Preencha os campos de Usuário e Senha para fazer o login.", vbExclamation, "AVISO"
112         Exit Sub
    
113     End If
End Sub

