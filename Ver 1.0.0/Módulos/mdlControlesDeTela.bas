Attribute VB_Name = "mdlControlesDeTela"
Option Explicit

Public Function VerificaPreenchimentoCampos(tela As String, acao As String) As Boolean
100     If tela = "FormLogin" Then
101         Select Case acao
        
                Case "entrar"
        
102                 If FormLogin.txtUsuario.Text <> "" And FormLogin.txtSenha.Text <> "" Then
103                     VerificaPreenchimentoCampos = True
104                 Else
105                     VerificaPreenchimentoCampos = False
106                 End If

107         End Select
108     End If
        
109     If tela = "FormCategorias" Then
110         Select Case acao
            
                Case "gravar"
                
111                 If FormCategorias.txtCodigo.Text <> "" And FormCategorias.txtDescricao.Text <> "" Then
112                     VerificaPreenchimentoCampos = True
113                 Else
114                     VerificaPreenchimentoCampos = False
115                 End If
                    
116             Case "excluir"
117                 If FormCategorias.txtCodigo.Text <> "" And FormCategorias.txtDescricao.Text <> "" Then
118                     VerificaPreenchimentoCampos = True
119                 Else
120                     VerificaPreenchimentoCampos = False
121                 End If
                
122             Case "editar"
123                 If FormCategorias.txtCodigo.Text <> "" And FormCategorias.txtDescricao.Text <> "" Then
124                     VerificaPreenchimentoCampos = True
125                 Else
126                     VerificaPreenchimentoCampos = False
127                 End If
128         End Select
129     End If

End Function

Public Sub PreparaCampos(tela As String, acao As String)
100     If tela = "FormCategorias" Then
101         Select Case acao
            
                Case "inicio":
102                 FormCategorias.FraCategorias.Caption = "Consulta Categoria"
103                 FormCategorias.txtCodigo.Enabled = True
104                 FormCategorias.txtDescricao.Enabled = False
105                 FormCategorias.txtObservacao.Enabled = False
106                 FormCategorias.OptDespesa.Enabled = False
107                 FormCategorias.OptReceita.Enabled = False
108                 FormCategorias.cmdEditar.Enabled = False
109                 FormCategorias.cmdGravar.Enabled = False
110                 FormCategorias.cmdExcluir.Enabled = False
111                 FormCategorias.cmdLimpar.Enabled = False
112                 FormCategorias.cmdCancelar.Enabled = False
                
                Case "novo":
                
113                 FormCategorias.FraCategorias.Caption = "Nova Categoria"
114                 FormCategorias.txtCodigo.Enabled = False
115                 FormCategorias.txtDescricao.Enabled = True
116                 FormCategorias.txtObservacao.Enabled = True
117                 FormCategorias.OptDespesa.Enabled = True
118                 FormCategorias.OptReceita.Enabled = True
119                 FormCategorias.cmdGravar.Enabled = True
120                 FormCategorias.cmdLimpar.Enabled = True
121                 FormCategorias.cmdCancelar.Enabled = True
122                 FormCategorias.cmdEditar.Enabled = False
123                 FormCategorias.cmdExcluir.Enabled = False
                
                Case "cancelar":
124                 FormCategorias.FraCategorias.Caption = "Consulta Categoria"
125                 FormCategorias.txtCodigo.Enabled = True
126                 FormCategorias.txtDescricao.Enabled = False
127                 FormCategorias.txtObservacao.Enabled = False
128                 FormCategorias.OptDespesa.Enabled = False
129                 FormCategorias.OptReceita.Enabled = False
130                 FormCategorias.cmdEditar.Enabled = False
131                 FormCategorias.cmdGravar.Enabled = False
132                 FormCategorias.cmdExcluir.Enabled = False
133                 FormCategorias.cmdLimpar.Enabled = False
134                 FormCategorias.cmdCancelar.Enabled = False
          
                Case "editar":
135                 FormCategorias.FraCategorias.Caption = "Editar Categoria"
136                 FormCategorias.txtCodigo.Enabled = False
137                 FormCategorias.txtDescricao.Enabled = True
138                 FormCategorias.txtObservacao.Enabled = True
139                 FormCategorias.OptDespesa.Enabled = True
140                 FormCategorias.OptReceita.Enabled = True
141                 FormCategorias.cmdEditar.Enabled = False
142                 FormCategorias.cmdGravar.Enabled = True
143                 FormCategorias.cmdExcluir.Enabled = False
144                 FormCategorias.cmdLimpar.Enabled = True
145                 FormCategorias.cmdCancelar.Enabled = True
146         End Select
147     End If

148     If tela = "FormDespesas" Then
149         Select Case acao
            
                Case "inicio":
150                 FormDespesas.FraDespesas.Caption = "Consulta Despesa"
151                 FormDespesas.txtCodigo.Enabled = True
152                 FormDespesas.txtDescricao.Enabled = False
153                 FormDespesas.txtObservacao.Enabled = False
154                 FormDespesas.ComCategoria.Enabled = False
155                 FormDespesas.ComCategoria.ListIndex = -1
156                 FormDespesas.txtValor.Enabled = False
157                 FormDespesas.OptPago.Enabled = False
158                 FormDespesas.OptPendente.Enabled = False
159                 FormDespesas.TDBDataDespesa.Enabled = False
160                 FormDespesas.cmdEditar.Enabled = False
161                 FormDespesas.cmdGravar.Enabled = False
162                 FormDespesas.cmdExcluir.Enabled = False
163                 FormDespesas.cmdLimpar.Enabled = False
164                 FormDespesas.cmdCancelar.Enabled = False
165         End Select
166     End If
End Sub

Public Sub LimpaCampos(tela As String, acao As String)
100     If tela = "FormCategorias" Then
101         Select Case acao
            
                Case "novo":
102                 FormCategorias.txtDescricao.Text = ""
103                 FormCategorias.txtObservacao.Text = ""
104                 FormCategorias.OptDespesa.Value = True
105                 FormCategorias.OptReceita.Value = False

                Case "limpar":
106                 FormCategorias.txtCodigo.Text = ""
107                 FormCategorias.txtDescricao.Text = ""
108                 FormCategorias.txtObservacao.Text = ""
109                 FormCategorias.OptDespesa.Value = True
110                 FormCategorias.OptReceita.Value = False
111         End Select
112     End If
End Sub
