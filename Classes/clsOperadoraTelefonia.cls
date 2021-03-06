VERSION 1.0 CLASS
Begin
	MultiUse = -1  'True
End
Attribute VB_Name = "clsOperadoraTelefonia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'Booleano que registra altera��es nos atributos chave da classe
Private booChaveAlterada As Boolean

'Booleano que registra altera��es nos atributos da classe
Private booAtributoAlterado As Boolean


'Atributos da Classe

'Atributo de backup e atributo identificador da Classe
Private bkpIdOperadora As Variant
Private lngIdOperadora As Variant

Private strNomeOperadora As Variant


'M�todos Get, Set e Let da Classe

Property Get chaveAlterada() As Variant

	chaveAlterada = booChaveAlterada

End Property

Property Get atributoAlterado() As Variant

	atributoAlterado = booAtributoAlterado

End Property

Property Get IdOperadora() As Variant

	IdOperadora = lngIdOperadora

End Property

Property Let IdOperadora(argIdOperadora As Variant)

	lngIdOperadora = argIdOperadora
	booAtributoAlterado = True
	booChaveAlterada = True

	If IsEmpty(bkpIdOperadora) Then
		bkpIdOperadora = lngIdOperadora
	End If

End Property

Property Get NomeOperadora() As Variant

	NomeOperadora = strNomeOperadora

End Property

Property Let NomeOperadora(argNomeOperadora As Variant)

	strNomeOperadora = argNomeOperadora
	booAtributoAlterado = True

End Property

Private Sub class_Initialize()
On Error Goto Err_Initialize

	'Insira aqui o c�digo que deve ser executado quando a classe for criada...


Exit_Initialize:
	Exit Sub

Err_Initialize:
	Goto Exit_Initialize

End Sub

Private Sub class_Terminate()
On Error Goto Err_Terminate

	'Insira aqui o c�digo que deve ser executado quando a classe for liberada...


Exit_Terminate:
	Exit Sub

Err_Terminate:
	Goto Exit_Terminate

End Sub

'M�todo Limpar - Limpa os valores dos atributos
Sub limpar()
On Error GoTo Err_limpar

	'Limpa os valores dos atributos
	IdOperadora = Null
	NomeOperadora = Null

	booChaveAlterada = False
	booAtributoAlterado = False

Exit_limpar:
	Exit Sub

Err_limpar:
	GoTo Exit_limpar

End Sub

'M�todo Existe [Com conhecimento de SQL]
'Verifica a exist�ncia do objeto OperadoraTelefonia na tabela correspondente no Banco de Dados
Function existe(argIdOperadora As Variant) As Boolean
On Error GoTo Err_existe

	Dim objCon As New aclConexaoBD
	Dim rstExiste As Recordset
	Dim strSql As String

	existe = False

	strSql = "Select * " & _
		"From OperadoraTelefonia " & _
		"Where IdOperadora = " & objCon.valorSql(argIdOperadora)

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

'M�todo Incluir [Com conhecimento de SQL]
'Inclui um novo objeto na tabela correspondente dentro do Banco de dados
Function incluir() As Boolean
On Error Goto Err_incluir

	Dim objCon As New aclConexaoBD
	Dim strSql As String

	strSql = "Insert Into " & _
		"OperadoraTelefonia(IdOperadora,NomeOperadora) " & _
		"Values(" & objCon.valorSql(IdOperadora) & "," & objCon.valorSql(NomeOperadora) & ")"

	incluir = (objCon.executa(strSql) > 0)

	If incluir Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza os campos de backup
		bkpIdOperadora = IdOperadora
	End If

Exit_incluir:
	Exit Function

Err_incluir:
	incluir = False
	Goto Exit_incluir

End Function

'M�todo Excluir [Com conhecimento de SQL]
'Exclui o objeto atual na tabela correspondente dentro do Banco de dados
Function excluir() As Boolean
On Error Goto Err_excluir

	Dim objCon As New aclConexaoBD
	Dim strSql As String

	strSql = "Delete From OperadoraTelefonia " & _
		"Where IdOperadora = " & objCon.valorSql(IdOperadora)

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

'M�todo Obter [Com conhecimento de SQL]
'Recupera o objeto OperadoraTelefonia atrav�s dos argumentos informados
Function obter(argIdOperadora As Variant) As Boolean
On Error GoTo Err_obter

	Dim objCon As New aclConexaoBD
	Dim rstObter As Recordset
	Dim strSql As String

	strSql = "Select * " & _
		"From OperadoraTelefonia " & _
		"Where IdOperadora = " & objCon.valorSql(argIdOperadora)

	Set rstObter = objCon.consulta(strSql)

	If rstObter.RecordCount = 0 Then
		obter = False
		Exit Function
	End If

	'Atualiza os campos de backup e os identificadores
	IdOperadora = argIdOperadora
	bkpIdOperadora = argIdOperadora

	'Atualiza os campos restantes
	NomeOperadora = rstObter.Fields("NomeOperadora")

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

'M�todo Salvar [Com conhecimento de SQL]
'Salva o objeto atual na tabela correspondente dentro do Banco de dados
Function salvar() As Boolean
On Error Goto Err_salvar

	Dim objCon As New aclConexaoBD
	Dim strSql As String

	If existe(bkpIdOperadora) Then
		strSql = "Update OperadoraTelefonia " & _
			"Set IdOperadora = " & objCon.valorSql(IdOperadora) & ", NomeOperadora = " & objCon.valorSql(NomeOperadora) & _
			" Where IdOperadora = " & objCon.valorSql(bkpIdOperadora)

		salvar = (objCon.executa(strSql) > 0)
	Else
		salvar = incluir
	End If

	If salvar Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza as vari�veis de backup com o novo valor da chave
		bkpIdOperadora = IdOperadora
	End If

Exit_salvar:
	Exit Function

Err_salvar:
	salvar = False
	Goto Exit_salvar

End Function

'Fim da classe...