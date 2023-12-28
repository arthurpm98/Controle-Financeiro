VERSION 5.00
Begin VB.Form FormCategorias 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Categorias"
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3540
      TabIndex        =   14
      Top             =   8505
      Width           =   1125
   End
   Begin VB.Frame FraCategorias 
      Appearance      =   0  'Flat
      Caption         =   "Categorias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8040
      Left            =   150
      TabIndex        =   4
      Top             =   405
      Width           =   11700
      Begin VB.ListBox ListCategorias 
         ForeColor       =   &H8000000D&
         Height          =   2205
         ItemData        =   "FormCategorias.frx":0000
         Left            =   135
         List            =   "FormCategorias.frx":0002
         TabIndex        =   13
         Top             =   5730
         Width           =   11445
      End
      Begin VB.TextBox txtCodigo 
         Height          =   300
         Left            =   1305
         TabIndex        =   9
         Top             =   480
         Width           =   1380
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1290
         TabIndex        =   8
         Top             =   990
         Width           =   9480
      End
      Begin VB.TextBox txtObservacao 
         Height          =   285
         Left            =   1275
         TabIndex        =   7
         Top             =   1500
         Width           =   9495
      End
      Begin VB.OptionButton OptDespesa 
         Appearance      =   0  'Flat
         Caption         =   "Despesa"
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   1230
         TabIndex        =   6
         Top             =   2010
         Width           =   915
      End
      Begin VB.OptionButton OptReceita 
         Appearance      =   0  'Flat
         Caption         =   "Receita"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   2205
         TabIndex        =   5
         Top             =   2145
         Width           =   1020
      End
      Begin VB.Label lblCódigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   690
         TabIndex        =   12
         Top             =   510
         Width           =   540
      End
      Begin VB.Label lblDescrição 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   450
         TabIndex        =   11
         Top             =   1005
         Width           =   780
      End
      Begin VB.Label lblObservação 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   285
         TabIndex        =   10
         Top             =   1515
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdLimpar 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10725
      TabIndex        =   3
      Top             =   8505
      Width           =   1125
   End
   Begin VB.CommandButton cmdExcluir 
      Appearance      =   0  'Flat
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      TabIndex        =   2
      Top             =   8505
      Width           =   1125
   End
   Begin VB.CommandButton cmdGravar 
      Appearance      =   0  'Flat
      Caption         =   "Gravar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1260
      TabIndex        =   1
      Top             =   8505
      Width           =   1125
   End
   Begin VB.CommandButton cmdEditar 
      Appearance      =   0  'Flat
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   135
      TabIndex        =   0
      Top             =   8505
      Width           =   1110
   End
   Begin VB.Menu mnuCategorias 
      Caption         =   "Categorias"
      Index           =   0
      Begin VB.Menu mnuCategoriasNova 
         Caption         =   "Nova Categoria"
         Index           =   1
      End
   End
End
Attribute VB_Name = "FormCategorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
100     PreparaCampos FormCategorias.Name, "cancelar"
101     LimpaCampos FormCategorias.Name, "limpar"
End Sub

Private Sub cmdEditar_Click()
100     Dim rs As ADODB.Recordset
101     Set rs = New ADODB.Recordset

102     If Trim(txtCodigo.Text) <> "" Then

103         If VerificaPreenchimentoCampos(FormCategorias.Name, "editar") = True Then
104             csql = "SELECT codigo_categoria FROM categorias WHERE codigo_categoria = " & Val(txtCodigo.Text) & " AND descricao_categoria = '" & UCase$(Trim(txtDescricao.Text)) & "'"

105             rs.CursorLocation = adUseClient
106             rs.Open csql, cn, adOpenStatic, adLockReadOnly
                
107             If Not rs.EOF Then
108                 PreparaCampos FormCategorias.Name, "editar"
109             End If
110             rs.Close
111         Else
112             MsgBox "Preencha os campos: Código, Descrição e Tipo de Categoria para editar o cadastro.", vbInformation
113             Exit Sub
114         End If
115     Else
116         MsgBox "Código ou Descrição da categoria não estão preenchidos.", vbInformation
117         Exit Sub
118     End If
End Sub

Private Sub cmdExcluir_Click()
100     Dim rs As ADODB.Recordset
101     Set rs = New ADODB.Recordset
102     If VerificaPreenchimentoCampos(FormCategorias.Name, "excluir") = True Then
103         csql = "SELECT codigo_categoria FROM categorias WHERE codigo_categoria = " & Val(txtCodigo.Text)
104         csql = csql & " AND descricao_categoria = '" & UCase$(Trim(txtDescricao.Text)) & "'"
            
105         rs.CursorLocation = adUseClient
106         rs.Open csql, cn, adOpenStatic, adLockReadOnly
107         If Not rs.EOF Then
108             If MsgBox("Deseja excluir o registro: " & Trim(txtCodigo.Text) & " - " & UCase$(Trim(txtDescricao.Text)) & " ?", vbYesNo) = vbYes Then
109                 cn.Execute "DELETE FROM categorias WHERE codigo_categoria = " & Val(txtCodigo.Text) & " AND descricao_categoria = '" & UCase$(Trim(txtDescricao.Text)) & "'"
110                 MsgBox "Categoria excluída! Código: " & Trim(txtCodigo.Text), vbInformation
111             End If
112             LimpaCampos FormCategorias.Name, "limpar"
113             PreparaCampos FormCategorias.Name, "inicio"
114             Call PreencheListBox
115         Else
116             MsgBox "Categoria não encontrada no banco de dados.", vbInformation
117             rs.Close
118             Exit Sub
119         End If
120         rs.Close
121     Else
122         MsgBox "Código ou Descrição da categoria não estão preenchidos.", vbInformation
123         Exit Sub
124     End If
End Sub

Private Sub cmdGravar_Click()
100     Dim rs As ADODB.Recordset
101     Set rs = New ADODB.Recordset
102     If VerificaPreenchimentoCampos(FormCategorias.Name, "gravar") = True Then
103         If FraCategorias.Caption = "Nova Categoria" Then

104             csql = "INSERT INTO categorias (codigo_categoria, descricao_categoria, observacao_categoria, receita_ou_despesa_categoria) VALUES("
105             csql = csql & Val(txtCodigo.Text) & ", '" & UCase$(Trim(txtDescricao.Text)) & "'"
106             If Trim(txtObservacao.Text) <> "" Then
107                 csql = csql & ",'" & UCase$(Trim(txtObservacao.Text)) & "'"
108             Else
109                 csql = csql & ",''"
110             End If
111             If OptReceita.Value = True Then
112                 csql = csql & ", 1)"
113             Else
114                 csql = csql & ", 0)"
115             End If
            
116             cn.Execute csql

117         ElseIf FraCategorias.Caption = "Editar Categoria" Then
118             csql = "SELECT codigo_categoria, descricao_categoria, observacao_categoria, receita_ou_despesa_categoria FROM categorias WHERE codigo_categoria = " & Val(txtCodigo.Text)

119             rs.CursorLocation = adUseClient
120             rs.Open csql, cn, adOpenStatic, adLockReadOnly
121             csql = ""
122             Auxiliar = ""
123             If Not rs.EOF Then
                
124                 csql = "UPDATE categorias SET "
125                 If UCase$(Trim$(rs!descricao_categoria)) <> UCase(Trim$(txtDescricao.Text)) Then
126                     Auxiliar = " descricao_categoria = '" & UCase(Trim$(txtDescricao.Text)) & "'"
127                 End If
128                 If UCase$(Trim$(rs!observacao_categoria)) <> UCase(Trim$(txtObservacao.Text)) Then
129                     If Auxiliar = "" Then
130                         Auxiliar = " observacao_categoria = '" & UCase(Trim$(txtObservacao.Text)) & "'"
131                     Else
132                         Auxiliar = Auxiliar & " , observacao_categoria = '" & UCase(Trim$(txtObservacao.Text)) & "'"
133                     End If
134                 End If
135                 If OptReceita.Value = True Then
136                     If Val(rs!receita_ou_despesa_categoria) = 0 Then
137                         If Auxiliar = "" Then
138                             Auxiliar = " receita_ou_despesa_categoria = " & OptReceita.Value
139                         Else
140                             Auxiliar = Auxiliar & " , receita_ou_despesa_categoria = " & OptReceita.Value
141                         End If
142                     End If
143                 Else
144                     If Val(rs!receita_ou_despesa_categoria) = 1 Then
145                         If Auxiliar = "" Then
146                             Auxiliar = " receita_ou_despesa_categoria = " & OptReceita.Value
147                         Else
148                             Auxiliar = Auxiliar & " , receita_ou_despesa_categoria = " & OptReceita.Value
149                         End If
150                     End If
151                 End If
152                 csql = csql & Auxiliar
153                 csql = csql & " WHERE codigo_categoria = " & Val(rs!codigo_categoria)
154             End If
155             rs.Close
156             cn.Execute csql
157         End If

158         MsgBox "Categoria cadastrada! Código: " & Trim(txtCodigo.Text), vbInformation
159         LimpaCampos FormCategorias.Name, "limpar"
160         PreparaCampos FormCategorias.Name, "inicio"
161         Call PreencheListBox
162     Else
163         MsgBox "Preencha os campos: Código, Descrição e Tipo de Categoria para salvar o cadastro.", vbInformation
164         Exit Sub
165     End If
End Sub

Private Sub cmdLimpar_Click()
100     txtDescricao.Text = ""
101     txtObservacao.Text = ""
102     OptDespesa.Value = True
103     OptReceita.Value = False
End Sub

Private Sub Form_Load()
100     PreparaCampos FormCategorias.Name, "inicio"
101     Call PreencheListBox
End Sub

Private Sub mnuCategoriasNova_Click(Index As Integer)
100     Dim ProxCod As Integer
101     Dim rs As ADODB.Recordset
102     Set rs = New ADODB.Recordset

103     csql = "SELECT MAX(codigo_categoria) FROM categorias"
104     rs.CursorLocation = adUseClient
105     rs.Open csql, cn, adOpenStatic, adLockReadOnly
106     If Not rs.EOF Then
107         ProxCod = IIf(IsNull(rs(0)) = True, 0, rs(0)) + 1
108     End If
109     rs.Close
        
110     txtCodigo.Text = ProxCod
111     ProxCod = 0
        
112     PreparaCampos FormCategorias.Name, "novo"
113     LimpaCampos FormCategorias.Name, "novo"
        
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
100     Dim rs As ADODB.Recordset
101     Set rs = New ADODB.Recordset

102     If KeyCode = 13 Then

103         csql = "SELECT codigo_categoria, descricao_categoria, observacao_categoria, receita_ou_despesa_categoria FROM categorias WHERE codigo_categoria = " & Val(txtCodigo.Text)
104         rs.CursorLocation = adUseClient
105         rs.Open csql, cn, adOpenStatic, adLockReadOnly

106         If Not rs.EOF Then
107             txtCodigo.Text = GetInteger(rs!codigo_categoria)
108             txtDescricao.Text = GetString(rs!descricao_categoria)
109             txtObservacao.Text = GetString(rs!observacao_categoria)
110             If CBool(rs!receita_ou_despesa_categoria) = True Then
111                 OptReceita.Value = True 'Receita
112             Else
113                 OptDespesa.Value = True 'Despesa
114             End If
115         Else
116             MsgBox "Código " & txtCodigo.Text & " não encontrado no banco de dados.", vbInformation
117             LimpaCampos FormCategorias.Name, "limpar"
118             cmdEditar.Enabled = False
119             cmdExcluir.Enabled = False
120             rs.Close
121             Exit Sub
122         End If
123         rs.Close
124         cmdEditar.Enabled = True
125         cmdExcluir.Enabled = True
126     End If
End Sub

Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
100     txtCodigo.Text = ObrigaNumerosInteiros(txtCodigo.Text)
End Sub

Private Sub PreencheListBox()
100     Dim bReceitaDespesas As Boolean
101     Dim rs As ADODB.Recordset
102     Set rs = New ADODB.Recordset

103     ListCategorias.Clear
        'Preenche o ListBox
104     csql = "SELECT codigo_categoria, descricao_categoria, receita_ou_despesa_categoria FROM categorias order by codigo_categoria"
105     rs.CursorLocation = adUseClient
106     rs.Open csql, cn, adOpenStatic, adLockReadOnly

107     Do While Not rs.EOF
108         If CBool(rs!receita_ou_despesa_categoria) = True Then
109             bReceitaDespesas = True 'Receita
110         Else
111             bReceitaDespesas = False 'Despesa
112         End If
113         Auxiliar = "Código: " & rs!codigo_categoria & " | Descrição: " & UCase(rs!descricao_categoria) & " | "
114         If bReceitaDespesas = True Then
115             Auxiliar = Auxiliar & "Receita"
116         Else
117             Auxiliar = Auxiliar & "Despesa"
118         End If
119         ListCategorias.AddItem Auxiliar
120         rs.MoveNext
121         DoEvents
122     Loop
123     rs.Close

End Sub
