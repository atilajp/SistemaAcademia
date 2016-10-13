VERSION 1.0 CLASS
Begin
	MultiUse = -1  'True
End
Attribute VB_Name = "clsEmpresa"
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
Private bkpIdEmpresa As Variant
Private lngIdEmpresa As Variant

Private strNomeEmpresa As Variant


'Métodos Get, Set e Let da Classe

Property Get chaveAlterada() As Variant

	chaveAlterada = booChaveAlterada

End Property

Property Get atributoAlterado() As Variant

	atributoAlterado = booAtributoAlterado

End Property

Property Get IdEmpresa() As Variant

	IdEmpresa = lngIdEmpresa

End Property

Property Let IdEmpresa(argIdEmpresa As Variant)

	lngIdEmpresa = argIdEmpresa
	booAtributoAlterado = True
	booChaveAlterada = True

	If IsEmpty(bkpIdEmpresa) Then
		bkpIdEmpresa = lngIdEmpresa
	End If

End Property

Property Get NomeEmpresa() As Variant

	NomeEmpresa = strNomeEmpresa

End Property

Property Let NomeEmpresa(argNomeEmpresa As Variant)

	strNomeEmpresa = argNomeEmpresa
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
	IdEmpresa = Null
	NomeEmpresa = Null

	booChaveAlterada = False
	booAtributoAlterado = False

Exit_limpar:
	Exit Sub

Err_limpar:
	GoTo Exit_limpar

End Sub

'Método Existe [Com conhecimento de SQL]
'Verifica a existência do objeto Empresa na tabela correspondente no Banco de Dados
Function existe(argIdEmpresa As Variant) As Boolean
On Error GoTo Err_existe

	Dim objCon As New aclConexaoBD
	Dim rstExiste As Recordset
	Dim strSql As String

	existe = False

	strSql = "Select * " & _
		"From Empresa " & _
		"Where IdEmpresa = " & objCon.valorSql(argIdEmpresa)

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
		"Empresa(IdEmpresa,NomeEmpresa) " & _
		"Values(" & objCon.valorSql(IdEmpresa) & "," & objCon.valorSql(NomeEmpresa) & ")"

	incluir = (objCon.executa(strSql) > 0)

	If incluir Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza os campos de backup
		bkpIdEmpresa = IdEmpresa
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

	strSql = "Delete From Empresa " & _
		"Where IdEmpresa = " & objCon.valorSql(IdEmpresa)

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
'Recupera o objeto Empresa através dos argumentos informados
Function obter(argIdEmpresa As Variant) As Boolean
On Error GoTo Err_obter

	Dim objCon As New aclConexaoBD
	Dim rstObter As Recordset
	Dim strSql As String

	strSql = "Select * " & _
		"From Empresa " & _
		"Where IdEmpresa = " & objCon.valorSql(argIdEmpresa)

	Set rstObter = objCon.consulta(strSql)

	If rstObter.RecordCount = 0 Then
		obter = False
		Exit Function
	End If

	'Atualiza os campos de backup e os identificadores
	IdEmpresa = argIdEmpresa
	bkpIdEmpresa = argIdEmpresa

	'Atualiza os campos restantes
	NomeEmpresa = rstObter.Fields("NomeEmpresa")

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

	If existe(bkpIdEmpresa) Then
		strSql = "Update Empresa " & _
			"Set IdEmpresa = " & objCon.valorSql(IdEmpresa) & ", NomeEmpresa = " & objCon.valorSql(NomeEmpresa) & _
			" Where IdEmpresa = " & objCon.valorSql(bkpIdEmpresa)

		salvar = (objCon.executa(strSql) > 0)
	Else
		salvar = incluir
	End If

	If salvar Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza as variáveis de backup com o novo valor da chave
		bkpIdEmpresa = IdEmpresa
	End If

Exit_salvar:
	Exit Function

Err_salvar:
	salvar = False
	Goto Exit_salvar

End Function

'Fim da classe...
	End If

	If salvar Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza as variáveis de backup com o novo valor da chave
		bkpIdEmpresa = IdEmpresa
	End If

Exit_salvar:
	Exit Function

Err_salvar:
	salvar = False
	Goto Exit_salvar

End Function

'Fim da classe...