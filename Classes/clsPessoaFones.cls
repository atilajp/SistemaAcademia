VERSION 1.0 CLASS
Begin
	MultiUse = -1  'True
End
Attribute VB_Name = "clsPessoaFones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'Booleano que registra alterações nos atributos chave da classe
Private booChaveAlterada As Boolean

'Booleano que registra alterações nos atributos da classe
Private booAtributoAlterado As Boolean


'Atributos da Classe

'Atributo de backup e atributo identificador da Classe
Private bkpIdPessoaFone As Variant
Private lngIdPessoaFone As Variant

Private strDDD As Variant

Private strDetalhe As Variant

Private lngIdOperadora As Variant

Private lngIdPessoa As Variant

Private lngNumero As Variant


'Métodos Get, Set e Let da Classe

Property Get chaveAlterada() As Variant

	chaveAlterada = booChaveAlterada

End Property

Property Get atributoAlterado() As Variant

	atributoAlterado = booAtributoAlterado

End Property

Property Get IdPessoaFone() As Variant

	IdPessoaFone = lngIdPessoaFone

End Property

Property Let IdPessoaFone(argIdPessoaFone As Variant)

	lngIdPessoaFone = argIdPessoaFone
	booAtributoAlterado = True
	booChaveAlterada = True

	If IsEmpty(bkpIdPessoaFone) Then
		bkpIdPessoaFone = lngIdPessoaFone
	End If

End Property

Property Get DDD() As Variant

	DDD = strDDD

End Property

Property Let DDD(argDDD As Variant)

	strDDD = argDDD
	booAtributoAlterado = True

End Property

Property Get Detalhe() As Variant

	Detalhe = strDetalhe

End Property

Property Let Detalhe(argDetalhe As Variant)

	strDetalhe = argDetalhe
	booAtributoAlterado = True

End Property

Property Get IdOperadora() As Variant

	IdOperadora = lngIdOperadora

End Property

Property Let IdOperadora(argIdOperadora As Variant)

	lngIdOperadora = argIdOperadora
	booAtributoAlterado = True

End Property

Property Get IdPessoa() As Variant

	IdPessoa = lngIdPessoa

End Property

Property Let IdPessoa(argIdPessoa As Variant)

	lngIdPessoa = argIdPessoa
	booAtributoAlterado = True

End Property

Property Get Numero() As Variant

	Numero = lngNumero

End Property

Property Let Numero(argNumero As Variant)

	lngNumero = argNumero
	booAtributoAlterado = True

End Property

Private Sub class_Initialize()
On Error Goto Err_Initialize

	'Insira aqui o código que deve ser executado quando a classe for criada...


Exit_Initialize:
	Exit Sub

Err_Initialize:
	Goto Exit_Initialize

End Sub

Private Sub class_Terminate()
On Error Goto Err_Terminate

	'Insira aqui o código que deve ser executado quando a classe for liberada...


Exit_Terminate:
	Exit Sub

Err_Terminate:
	Goto Exit_Terminate

End Sub

'Método Limpar - Limpa os valores dos atributos
Sub limpar()
On Error GoTo Err_limpar

	'Limpa os valores dos atributos
	IdPessoaFone = Null
	DDD = Null
	Detalhe = Null
	IdOperadora = Null
	IdPessoa = Null
	Numero = Null

	booChaveAlterada = False
	booAtributoAlterado = False

Exit_limpar:
	Exit Sub

Err_limpar:
	GoTo Exit_limpar

End Sub

'Método Existe [Com conhecimento de SQL]
'Verifica a existência do objeto PessoaFones na tabela correspondente no Banco de Dados
Function existe(argIdPessoaFone As Variant) As Boolean
On Error GoTo Err_existe

	Dim objCon As New aclConexaoBD
	Dim rstExiste As Recordset
	Dim strSql As String

	existe = False

	strSql = "Select * " & _
		"From PessoaFones " & _
		"Where IdPessoaFone = " & objCon.valorSql(argIdPessoaFone)

	Set rstExiste = objCon.consulta(strSql)

	If rstExiste.RecordCount > 0 Then
		existe = True
	End If

	'Fecha o Recordset existe
	rstExiste.close

Exit_existe:
	Set rstExiste = Nothing
	Exit Function

Err_existe:
	existe = False
	GoTo Exit_existe

End Function

'Método Incluir [Com conhecimento de SQL]
'Inclui um novo objeto na tabela correspondente dentro do Banco de dados
Function incluir() As Boolean
On Error Goto Err_incluir

	Dim objCon As New aclConexaoBD
	Dim strSql As String

	strSql = "Insert Into " & _
		"PessoaFones(IdPessoaFone,DDD,Detalhe,IdOperadora,IdPessoa,Numero) " & _
		"Values(" & objCon.valorSql(IdPessoaFone) & "," & objCon.valorSql(DDD) & "," & objCon.valorSql(Detalhe) & "," & objCon.valorSql(IdOperadora) & "," & objCon.valorSql(IdPessoa) & "," & objCon.valorSql(Numero) & ")"

	incluir = (objCon.executa(strSql) > 0)

	If incluir Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza os campos de backup
		bkpIdPessoaFone = IdPessoaFone
	End If

Exit_incluir:
	Exit Function

Err_incluir:
	incluir = False
	Goto Exit_incluir

End Function

'Método Excluir [Com conhecimento de SQL]
'Exclui o objeto atual na tabela correspondente dentro do Banco de dados
Function excluir() As Boolean
On Error Goto Err_excluir

	Dim objCon As New aclConexaoBD
	Dim strSql As String

	strSql = "Delete From PessoaFones " & _
		"Where IdPessoaFone = " & objCon.valorSql(IdPessoaFone)

		excluir = (objCon.executa(strSql) > 0)

	If excluir Then
		Call limpar
		booChaveAlterada = False
		booAtributoAlterado = False
	End If

Exit_excluir:
	Exit Function

Err_excluir:
	excluir = false
	Goto Exit_excluir

End Function

'Método Obter [Com conhecimento de SQL]
'Recupera o objeto PessoaFones através dos argumentos informados
Function obter(argIdPessoaFone As Variant) As Boolean
On Error GoTo Err_obter

	Dim objCon As New aclConexaoBD
	Dim rstObter As Recordset
	Dim strSql As String

	strSql = "Select * " & _
		"From PessoaFones " & _
		"Where IdPessoaFone = " & objCon.valorSql(argIdPessoaFone)

	Set rstObter = objCon.consulta(strSql)

	If rstObter.RecordCount = 0 Then
		obter = False
		Exit Function
	End If

	'Atualiza os campos de backup e os identificadores
	IdPessoaFone = argIdPessoaFone
	bkpIdPessoaFone = argIdPessoaFone

	'Atualiza os campos restantes
	DDD = rstObter.Fields("DDD")
	Detalhe = rstObter.Fields("Detalhe")
	IdOperadora = rstObter.Fields("IdOperadora")
	IdPessoa = rstObter.Fields("IdPessoa")
	Numero = rstObter.Fields("Numero")

	obter = True

	If obter Then
		booChaveAlterada = False
		booAtributoAlterado = False
	End If

	'Fecha o Recordset obter
	rstObter.close

Exit_obter:
	Set rstObter = Nothing
	Exit Function

Err_obter:
	obter = False
	GoTo Exit_obter

End Function

'Método Salvar [Com conhecimento de SQL]
'Salva o objeto atual na tabela correspondente dentro do Banco de dados
Function salvar() As Boolean
On Error Goto Err_salvar

	Dim objCon As New aclConexaoBD
	Dim strSql As String

	If existe(bkpIdPessoaFone) Then
		strSql = "Update PessoaFones " & _
			"Set IdPessoaFone = " & objCon.valorSql(IdPessoaFone) & ", DDD = " & objCon.valorSql(DDD) & ", Detalhe = " & objCon.valorSql(Detalhe) & ", IdOperadora = " & objCon.valorSql(IdOperadora) & ", IdPessoa = " & objCon.valorSql(IdPessoa) & ", Numero = " & objCon.valorSql(Numero) & _
			" Where IdPessoaFone = " & objCon.valorSql(bkpIdPessoaFone)

		salvar = (objCon.executa(strSql) > 0)
	Else
		salvar = incluir
	End If

	If salvar Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza as variáveis de backup com o novo valor da chave
		bkpIdPessoaFone = IdPessoaFone
	End If

Exit_salvar:
	Exit Function

Err_salvar:
	salvar = False
	Goto Exit_salvar

End Function


'Métodos Adicionais

'Método que retorna um Vetor de Registros da classe
Public Function vetorPessoaFones(argAtributo As String, argValor As Long) As Recordset

	Dim strSql As String

	Dim objCon As New aclConexaoBD

	strSql = "Select * From PessoaFones Where " & argAtributo & " = " & CStr(argValor)

	Set vetorPessoaFones = objCon.consulta(strSql)

End Function

'Fim da classe...