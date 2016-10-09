VERSION 1.0 CLASS
Begin
	MultiUse = -1  'True
End
Attribute VB_Name = "clsCidades"
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

'Vetor de Registro para a clsBairros
Public rstBairros As Recordset
Private objBairros As New clsBairros

'Atributo de backup e atributo identificador da Classe
Private bkpIdCidades As Variant
Private lngIdCidades As Variant

Private lngIdUf As Variant

Private strNomeCidades As Variant


'Métodos Get, Set e Let da Classe

Property Get chaveAlterada() As Variant

	chaveAlterada = booChaveAlterada

End Property

Property Get atributoAlterado() As Variant

	atributoAlterado = booAtributoAlterado

End Property

Property Get IdCidades() As Variant

	IdCidades = lngIdCidades

End Property

Property Let IdCidades(argIdCidades As Variant)

	lngIdCidades = argIdCidades
	booAtributoAlterado = True
	'Atribui o Vetor Recordset retornado pelo método da clsBairros
	Set rstBairros = objBairros.vetorBairros("IdCidades", IdCidades)
	booChaveAlterada = True

	If IsEmpty(bkpIdCidades) Then
		bkpIdCidades = lngIdCidades
	End If

End Property

Property Get IdUf() As Variant

	IdUf = lngIdUf

End Property

Property Let IdUf(argIdUf As Variant)

	lngIdUf = argIdUf
	booAtributoAlterado = True

End Property

Property Get NomeCidades() As Variant

	NomeCidades = strNomeCidades

End Property

Property Let NomeCidades(argNomeCidades As Variant)

	strNomeCidades = argNomeCidades
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
	IdCidades = Null
	IdUf = Null
	NomeCidades = Null

	booChaveAlterada = False
	booAtributoAlterado = False

Exit_limpar:
	Exit Sub

Err_limpar:
	GoTo Exit_limpar

End Sub

'Método Existe [Com conhecimento de SQL]
'Verifica a existência do objeto Cidades na tabela correspondente no Banco de Dados
Function existe(argIdCidades As Variant) As Boolean
On Error GoTo Err_existe

	Dim objCon As New aclConexaoBD
	Dim rstExiste As Recordset
	Dim strSql As String

	existe = False

	strSql = "Select * " & _
		"From Cidades " & _
		"Where IdCidades = " & objCon.valorSql(argIdCidades)

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
		"Cidades(IdCidades,IdUf,NomeCidades) " & _
		"Values(" & objCon.valorSql(IdCidades) & "," & objCon.valorSql(IdUf) & "," & objCon.valorSql(NomeCidades) & ")"

	incluir = (objCon.executa(strSql) > 0)

	If incluir Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza os campos de backup
		bkpIdCidades = IdCidades
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

	strSql = "Delete From Cidades " & _
		"Where IdCidades = " & objCon.valorSql(IdCidades)

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
'Recupera o objeto Cidades através dos argumentos informados
Function obter(argIdCidades As Variant) As Boolean
On Error GoTo Err_obter

	Dim objCon As New aclConexaoBD
	Dim rstObter As Recordset
	Dim strSql As String

	strSql = "Select * " & _
		"From Cidades " & _
		"Where IdCidades = " & objCon.valorSql(argIdCidades)

	Set rstObter = objCon.consulta(strSql)

	If rstObter.RecordCount = 0 Then
		obter = False
		Exit Function
	End If

	'Atualiza os campos de backup e os identificadores
	IdCidades = argIdCidades
	bkpIdCidades = argIdCidades

	'Atualiza os campos restantes
	IdUf = rstObter.Fields("IdUf")
	NomeCidades = rstObter.Fields("NomeCidades")

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

	If existe(bkpIdCidades) Then
		strSql = "Update Cidades " & _
			"Set IdCidades = " & objCon.valorSql(IdCidades) & ", IdUf = " & objCon.valorSql(IdUf) & ", NomeCidades = " & objCon.valorSql(NomeCidades) & _
			" Where IdCidades = " & objCon.valorSql(bkpIdCidades)

		salvar = (objCon.executa(strSql) > 0)
	Else
		salvar = incluir
	End If

	If salvar Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza as variáveis de backup com o novo valor da chave
		bkpIdCidades = IdCidades
	End If

Exit_salvar:
	Exit Function

Err_salvar:
	salvar = False
	Goto Exit_salvar

End Function


'Métodos Adicionais

'Método que retorna um Vetor de Registros da classe
Public Function vetorCidades(argAtributo As String, argValor As Long) As Recordset

	Dim strSql As String

	Dim objCon As New aclConexaoBD

	strSql = "Select * From Cidades Where " & argAtributo & " = " & CStr(argValor)

	Set vetorCidades = objCon.consulta(strSql)

End Function

'Fim da classe...