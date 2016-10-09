VERSION 1.0 CLASS
Begin
	MultiUse = -1  'True
End
Attribute VB_Name = "clsUfs"
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

'Recordset e Objeto de referência para a classe de Vetor de Registro

'Vetor de Registro para a clsCidades
Public rstCidades As Recordset
Private objCidades As New clsCidades

'Atributo de backup e atributo identificador da Classe
Private bkpIdUF As Variant
Private lngIdUF As Variant

Private strNome As Variant

Private strUF As Variant


'Métodos Get, Set e Let da Classe

Property Get chaveAlterada() As Variant

	chaveAlterada = booChaveAlterada

End Property

Property Get atributoAlterado() As Variant

	atributoAlterado = booAtributoAlterado

End Property

Property Get IdUF() As Variant

	IdUF = lngIdUF

End Property

Property Let IdUF(argIdUF As Variant)

	lngIdUF = argIdUF
	booAtributoAlterado = True
	'Atribui o Vetor Recordset retornado pelo método da clsCidades
	Set rstCidades = objCidades.vetorCidades("IdUf", IdUF)
	booChaveAlterada = True

	If IsEmpty(bkpIdUF) Then
		bkpIdUF = lngIdUF
	End If

End Property

Property Get Nome() As Variant

	Nome = strNome

End Property

Property Let Nome(argNome As Variant)

	strNome = argNome
	booAtributoAlterado = True

End Property

Property Get UF() As Variant

	UF = strUF

End Property

Property Let UF(argUF As Variant)

	strUF = argUF
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
	IdUF = Null
	Nome = Null
	UF = Null

	booChaveAlterada = False
	booAtributoAlterado = False

Exit_limpar:
	Exit Sub

Err_limpar:
	GoTo Exit_limpar

End Sub

'Método Existe [Com conhecimento de SQL]
'Verifica a existência do objeto Ufs na tabela correspondente no Banco de Dados
Function existe(argIdUF As Variant) As Boolean
On Error GoTo Err_existe

	Dim objCon As New aclConexaoBD
	Dim rstExiste As Recordset
	Dim strSql As String

	existe = False

	strSql = "Select * " & _
		"From Ufs " & _
		"Where IdUF = " & objCon.valorSql(argIdUF)

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
		"Ufs(IdUF,Nome,UF) " & _
		"Values(" & objCon.valorSql(IdUF) & "," & objCon.valorSql(Nome) & "," & objCon.valorSql(UF) & ")"

	incluir = (objCon.executa(strSql) > 0)

	If incluir Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza os campos de backup
		bkpIdUF = IdUF
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

	strSql = "Delete From Ufs " & _
		"Where IdUF = " & objCon.valorSql(IdUF)

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
'Recupera o objeto Ufs através dos argumentos informados
Function obter(argIdUF As Variant) As Boolean
On Error GoTo Err_obter

	Dim objCon As New aclConexaoBD
	Dim rstObter As Recordset
	Dim strSql As String

	strSql = "Select * " & _
		"From Ufs " & _
		"Where IdUF = " & objCon.valorSql(argIdUF)

	Set rstObter = objCon.consulta(strSql)

	If rstObter.RecordCount = 0 Then
		obter = False
		Exit Function
	End If

	'Atualiza os campos de backup e os identificadores
	IdUF = argIdUF
	bkpIdUF = argIdUF

	'Atualiza os campos restantes
	Nome = rstObter.Fields("Nome")
	UF = rstObter.Fields("UF")

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

	If existe(bkpIdUF) Then
		strSql = "Update Ufs " & _
			"Set IdUF = " & objCon.valorSql(IdUF) & ", Nome = " & objCon.valorSql(Nome) & ", UF = " & objCon.valorSql(UF) & _
			" Where IdUF = " & objCon.valorSql(bkpIdUF)

		salvar = (objCon.executa(strSql) > 0)
	Else
		salvar = incluir
	End If

	If salvar Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza as variáveis de backup com o novo valor da chave
		bkpIdUF = IdUF
	End If

Exit_salvar:
	Exit Function

Err_salvar:
	salvar = False
	Goto Exit_salvar

End Function

'Fim da classe...