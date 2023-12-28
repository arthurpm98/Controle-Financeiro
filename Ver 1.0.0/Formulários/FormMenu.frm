VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form FormMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle Financeiro - Versão 1.0.0"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TrueDBGrid60.TDBGrid TDBMenu 
      Height          =   7065
      Left            =   60
      OleObjectBlob   =   "FormMenu.frx":0000
      TabIndex        =   0
      Top             =   1200
      Width           =   11895
   End
   Begin VB.Label lblMenuDespesas 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Despesas Pendentes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4605
      TabIndex        =   1
      Top             =   735
      Width           =   2940
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "Cadastros"
      Index           =   0
      Begin VB.Menu mnuCadastroReceitas 
         Caption         =   "Receitas"
         Index           =   1
      End
      Begin VB.Menu mnuCadastroDespesas 
         Caption         =   "Despesas"
         Index           =   2
      End
      Begin VB.Menu mnuCadastroCategorias 
         Caption         =   "Categorias"
         Index           =   3
      End
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
100     Dim vMatriz As New XArrayDB
101     Dim rs       As ADODB.Recordset
102     Dim contador As Integer
103     Dim sMensagem As String

104     Set rs = New ADODB.Recordset

105     vMatriz.ReDim 1, vMatriz.Count(1) + 1, 0, 4
106     contador = 0
        
107     csql = "SELECT codigo_pagamento, data_pagamento, descricao_pagamento, valor_pagamento, observacao_pagamento "
108     csql = csql & " FROM pagamentos_contas WHERE conta_paga = 0"
109     rs.CursorLocation = adUseClient
110     rs.Open csql, cn, adOpenStatic, adLockReadOnly
111     Do While Not rs.EOF
112         contador = contador + 1

113         vMatriz.ReDim 1, vMatriz.Count(1) + 1, 0, 4

114         vMatriz(contador, 0) = GetInteger(rs!codigo_pagamento)
115         vMatriz(contador, 1) = GetDate(rs!data_pagamento)
116         vMatriz(contador, 2) = GetString(rs!descricao_pagamento)
117         vMatriz(contador, 3) = GetCurrency(rs!valor_pagamento)
118         vMatriz(contador, 4) = GetString(rs!observacao_pagamento)
        
119         If rs!data_pagamento <= Date Then
120             If sMensagem = "" Then
121                 sMensagem = "Atenção! Você têm alguma(s) conta(s) para pagar: " & vbCrLf & GetString(rs!descricao_pagamento) & " - " & GetCurrency(rs!valor_pagamento) & " Reais - " & GetDate(rs!data_pagamento) & vbCrLf
122             Else
123                sMensagem = sMensagem & "Atenção! Você têm alguma(s) conta(s) para pagar: " & vbCrLf & GetString(rs!descricao_pagamento) & " - " & GetCurrency(rs!valor_pagamento) & " Reais - " & GetDate(rs!data_pagamento) & vbCrLf
124             End If
125         End If
            
126         rs.MoveNext
127         DoEvents
128     Loop
129     rs.Close

130     TDBMenu.Array = vMatriz
131     TDBMenu.ReBind
132     TDBMenu.Refresh
        
133     If sMensagem <> "" Then MsgBox sMensagem, vbInformation
End Sub

Private Sub mnuCadastroCategorias_Click(Index As Integer)
100     FormMenu.Enabled = False
101     FormCategorias.Show 1
End Sub

Private Sub mnuCadastroDespesas_Click(Index As Integer)
100     FormMenu.Enabled = False
101     FormDespesas.Show 1
End Sub

Private Sub mnuCadastroReceitas_Click(Index As Integer)
100     FormMenu.Enabled = False
101     FormReceitas.Show 1
End Sub
