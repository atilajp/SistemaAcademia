VERSION 1.0 CLASS
Begin
	MultiUse = -1  'True
End
Attribute VB_Name = "clsBairros"
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
Private bkpIdBairro As Variant
Private lngIdBairro As Variant

Private lngIdCidades As Variant

Private strNomeBairro As Variant


'Métodos Get, Set e Let da Classe

Property Get chaveAlterada() As Variant

	chaveAlterada = booChaveAlterada

End Property

Property Get atributoAlterado() As Variant

	atributoAlterado = booAtributoAlterado

End Property

Property Get IdBairro() As Variant

	IdBairro = lngIdBairro

End Property

Property Let IdBairro(argIdBairro As Variant)

	lngIdBairro = argIdBairro
	booAtributoAlterado = True
	booChaveAlterada = True

	If IsEmpty(bkpIdBairro) Then
		bkpIdBairro = lngIdBairro
	End If

End Property

Property Get IdCidades() As Variant

	IdCidades = lngIdCidades

End Property

Property Let IdCidades(argIdCidades As Variant)

	lngIdCidades = argIdCidades
	booAtributoAlterado = True

End Property

Property Get NomeBairro() As Variant

	NomeBairro = strNomeBairro

End Property

Property Let NomeBairro(argNomeBairro As Variant)

	strNomeBairro = argNomeBairro
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
	IdBairro = Null
	IdCidades = Null
	NomeBairro = Null

	booChaveAlterada = False
	booAtributoAlterado = False

Exit_limpar:
	Exit Sub

Err_limpar:
	GoTo Exit_limpar

End Sub

'Método Existe [Com conhecimento de SQL]
'Verifica a existência do objeto Bairros na tabela correspondente no Banco de Dados
Function existe(argIdBairro As Variant) As Boolean
On Error GoTo Err_existe

	Dim objCon As New aclConexaoBD
	Dim rstExiste As Recordset
	Dim strSql As String

	existe = False

	strSql = "Select * " & _
		"From Bairros " & _
		"Where IdBairro = " & objCon.valorSql(argIdBairro)

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
		"Bairros(IdBairro,IdCidades,NomeBairro) " & _
		"Values(" & objCon.valorSql(IdBairro) & "," & objCon.valorSql(IdCidades) & "," & objCon.valorSql(NomeBairro) & ")"

	incluir = (objCon.executa(strSql) > 0)

	If incluir Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza os campos de backup
		bkpIdBairro = IdBairro
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

	strSql = "Delete From Bairros " & _
		"Where IdBairro = " & objCon.valorSql(IdBairro)

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
'Recupera o objeto Bairros através dos argumentos informados
Function obter(argIdBairro As Variant) As Boolean
On Error GoTo Err_obter

	Dim objCon As New aclConexaoBD
	Dim rstObter As Recordset
	Dim strSql As String

	strSql = "Select * " & _
		"From Bairros " & _
		"Where IdBairro = " & objCon.valorSql(argIdBairro)

	Set rstObter = objCon.consulta(strSql)

	If rstObter.RecordCount = 0 Then
		obter = False
		Exit Function
	End If

	'Atualiza os campos de backup e os identificadores
	IdBairro = argIdBairro
	bkpIdBairro = argIdBairro

	'Atualiza os campos restantes
	IdCidades = rstObter.Fields("IdCidades")
	NomeBairro = rstObter.Fields("NomeBairro")

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

	If existe(bkpIdBairro) Then
		strSql = "Update Bairros " & _
			"Set IdBairro = " & objCon.valorSql(IdBairro) & ", IdCidades = " & objCon.valorSql(IdCidades) & ", NomeBairro = " & objCon.valorSql(NomeBairro) & _
			" Where IdBairro = " & objCon.valorSql(bkpIdBairro)

		salvar = (objCon.executa(strSql) > 0)
	Else
		salvar = incluir
	End If

	If salvar Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza as variáveis de backup com o novo valor da chave
		bkpIdBairro = IdBairro
	End If

Exit_salvar:
	Exit Function

Err_salvar:
	salvar = False
	Goto Exit_salvar

End Function


'Métodos Adicionais

'Método que retorna um Vetor de Registros da classe
Public Function vetorBairros(argAtributo As String, argValor As Long) As Recordset

	Dim strSql As String

	Dim objCon As New aclConexaoBD

	strSql = "Select * From Bairros Where " & argAtributo & " = " & CStr(argValor)

	Set vetorBairros = objCon.consulta(strSql)

End Function

'Fim da classe...