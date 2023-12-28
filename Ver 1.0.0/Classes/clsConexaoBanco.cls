VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConexaoBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Function strConexaoBanco(Usuario_ As String, Senha_ As String)
100     strConexaoBanco = "DRIVER={MySQL ODBC 8.0 UNICODE DRIVER};Server=ns148.hostgator.com.br;Port=3306;Database=rhinfo16_cf;UID=" & Usuario_ & ";PWD=" & Senha_ & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
End Function
Public Function conectaBancoDados(Usuario_ As String, Senha_ As String) As Boolean
100     conectaBancoDados = False
102     On Error GoTo ErroConexao
104     Set cn = New ADODB.Connection
106     With cn
108         .CursorLocation = adUseClient
110         .ConnectionString = strConexaoBanco(Usuario_, Senha_)
112         .Open
114     End With
116     conectaBancoDados = True

        'Cria uma tabela de teste.
118     Call CriarTabeladeTeste
        
        'Cria todas as tabelas do banco (PRIMEIRO ACESSO AO SISTEMA)
        'Call CriarEstruturaBanco

120     Exit Function
ErroConexao:
122     MsgBox "Aconteceu um erro ao conectar ao banco de dados, verifique.", vbExclamation, "AVISO"
End Function

Private Sub CriarTabeladeTeste()
        'cria tabela no banco se nao existir
100     cn.Execute "DROP TABLE IF EXISTS `teste`;"
102     csql = csql & " CREATE TABLE IF NOT EXISTS `teste` ("
104     csql = csql & "`id` INT(5) NOT NULL AUTO_INCREMENT,"
106     csql = csql & "`nome` VARCHAR(80) NOT NULL DEFAULT '' COLLATE 'latin1_swedish_ci',"
108     csql = csql & "`nascimento` DATE NULL DEFAULT NULL,"
110     csql = csql & "`salario` DECIMAL(16,2) NULL DEFAULT '0.00',"
112     csql = csql & "PRIMARY KEY (`id`) USING BTREE)"
114     csql = csql & " ENGINE=InnoDB DEFAULT CHARSET=utf8"
116     cn.Execute csql
End Sub
