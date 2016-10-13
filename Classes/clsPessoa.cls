VERSION 1.0 CLASS
Begin
	MultiUse = -1  'True
End
Attribute VB_Name = "clsPessoa"
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

'Vetor de Registro para a clsPessoaFones
Public rstPessoaFones As Recordset
Private objPessoaFones As New clsPessoaFones

'Atributo de backup e atributo identificador da Classe
Private bkpIdPessoa As Variant
Private lngIdPessoa As Variant

Private lngBairroPessoa As Variant

Private lngCEPPessoa As Variant

Private lngCidadePessoa As Variant

Private strComplementoPontoDeRe As Variant

Private strContato As Variant

Private lngCPFCNPJ As Variant

Private dtmDataNascimento As Variant

Private lngDependenteDe As Variant

'Objeto da classe de Descrição de Tipo
Private objetoPessoa As New clsPessoa

Private booDesativado As Variant

Private strEmail As Variant

Private strEnderecoPessoa As Variant

Private strFantasiaApelido As Variant

Private strFonePrincipa As Variant

Private lngIdEmpresa As Variant

'Objeto da classe de Descrição de Tipo
Private objetoEmpresa As New clsEmpresa

Private lngIdPessoaTipo As Variant

'Objeto da classe de Descrição de Tipo
Private objetoPessoaTipo As New clsPessoaTipo

Private strNomeRazao As Variant

Private strNumPessoa As Variant

Private strRGIE As Variant

Private lngUFPessoa As Variant


'Métodos Get, Set e Let da Classe

Property Get chaveAlterada() As Variant

	chaveAlterada = booChaveAlterada

End Property

Property Get atributoAlterado() As Variant

	atributoAlterado = booAtributoAlterado

End Property

Property Get IdPessoa() As Variant

	IdPessoa = lngIdPessoa

End Property

Property Let IdPessoa(argIdPessoa As Variant)

	lngIdPessoa = argIdPessoa
	booAtributoAlterado = True
	'Atribui o Vetor Recordset retornado pelo método da clsPessoaFones
	Set rstPessoaFones = objPessoaFones.vetorPessoaFones("IdPessoa", IdPessoa)
	booChaveAlterada = True

	If IsEmpty(bkpIdPessoa) Then
		bkpIdPessoa = lngIdPessoa
	End If

End Property

Property Get BairroPessoa() As Variant

	BairroPessoa = lngBairroPessoa

End Property

Property Let BairroPessoa(argBairroPessoa As Variant)

	lngBairroPessoa = argBairroPessoa
	booAtributoAlterado = True

End Property

Property Get CEPPessoa() As Variant

	CEPPessoa = lngCEPPessoa

End Property

Property Let CEPPessoa(argCEPPessoa As Variant)

	lngCEPPessoa = argCEPPessoa
	booAtributoAlterado = True

End Property

Property Get CidadePessoa() As Variant

	CidadePessoa = lngCidadePessoa

End Property

Property Let CidadePessoa(argCidadePessoa As Variant)

	lngCidadePessoa = argCidadePessoa
	booAtributoAlterado = True

End Property

Property Get ComplementoPontoDeRe() As Variant

	ComplementoPontoDeRe = strComplementoPontoDeRe

End Property

Property Let ComplementoPontoDeRe(argComplementoPontoDeRe As Variant)

	strComplementoPontoDeRe = argComplementoPontoDeRe
	booAtributoAlterado = True

End Property

Property Get Contato() As Variant

	Contato = strContato

End Property

Property Let Contato(argContato As Variant)

	strContato = argContato
	booAtributoAlterado = True

End Property

Property Get CPFCNPJ() As Variant

	CPFCNPJ = lngCPFCNPJ

End Property

Property Let CPFCNPJ(argCPFCNPJ As Variant)

	lngCPFCNPJ = argCPFCNPJ
	booAtributoAlterado = True

End Property

Property Get DataNascimento() As Variant

	DataNascimento = dtmDataNascimento

End Property

Property Let DataNascimento(argDataNascimento As Variant)

	dtmDataNascimento = argDataNascimento
	booAtributoAlterado = True

End Property

Property Get DependenteDe() As Variant

	DependenteDe = lngDependenteDe

End Property

Property Let DependenteDe(argDependenteDe As Variant)

	lngDependenteDe = argDependenteDe
	booAtributoAlterado = True

End Property

Property Get Desativado() As Variant

	Desativado = booDesativado

End Property

Property Let Desativado(argDesativado As Variant)

	booDesativado = argDesativado
	booAtributoAlterado = True

End Property

Property Get Email() As Variant

	Email = strEmail

End Property

Property Let Email(argEmail As Variant)

	strEmail = argEmail
	booAtributoAlterado = True

End Property

Property Get EnderecoPessoa() As Variant

	EnderecoPessoa = strEnderecoPessoa

End Property

Property Let EnderecoPessoa(argEnderecoPessoa As Variant)

	strEnderecoPessoa = argEnderecoPessoa
	booAtributoAlterado = True

End Property

Property Get FantasiaApelido() As Variant

	FantasiaApelido = strFantasiaApelido

End Property

Property Let FantasiaApelido(argFantasiaApelido As Variant)

	strFantasiaApelido = argFantasiaApelido
	booAtributoAlterado = True

End Property

Property Get FonePrincipa() As Variant

	FonePrincipa = strFonePrincipa

End Property

Property Let FonePrincipa(argFonePrincipa As Variant)

	strFonePrincipa = argFonePrincipa
	booAtributoAlterado = True

End Property

Property Get IdEmpresa() As Variant

	IdEmpresa = lngIdEmpresa

End Property

Property Let IdEmpresa(argIdEmpresa As Variant)

	lngIdEmpresa = argIdEmpresa
	booAtributoAlterado = True

End Property

Property Get IdPessoaTipo() As Variant

	IdPessoaTipo = lngIdPessoaTipo

End Property

Property Let IdPessoaTipo(argIdPessoaTipo As Variant)

	lngIdPessoaTipo = argIdPessoaTipo
	booAtributoAlterado = True

End Property

Property Get NomeRazao() As Variant

	NomeRazao = strNomeRazao

End Property

Property Let NomeRazao(argNomeRazao As Variant)

	strNomeRazao = argNomeRazao
	booAtributoAlterado = True

End Property

Property Get NumPessoa() As Variant

	NumPessoa = strNumPessoa

End Property

Property Let NumPessoa(argNumPessoa As Variant)

	strNumPessoa = argNumPessoa
	booAtributoAlterado = True

End Property

Property Get RGIE() As Variant

	RGIE = strRGIE

End Property

Property Let RGIE(argRGIE As Variant)

	strRGIE = argRGIE
	booAtributoAlterado = True

End Property

Property Get UFPessoa() As Variant

	UFPessoa = lngUFPessoa

End Property

Property Let UFPessoa(argUFPessoa As Variant)

	lngUFPessoa = argUFPessoa
	booAtributoAlterado = True

End Property

Property Get objPessoa() As Variant

	If Not IsEmpty(DependenteDe) Then
		If objetoPessoa.obter(DependenteDe) Then
			Set objPessoa = objetoPessoa
		End If
	End If

End Property

Property Get objEmpresa() As Variant

	If Not IsEmpty(IdEmpresa) Then
		If objetoEmpresa.obter(IdEmpresa) Then
			Set objEmpresa = objetoEmpresa
		End If
	End If

End Property

Property Get objPessoaTipo() As Variant

	If Not IsEmpty(IdPessoaTipo) Then
		If objetoPessoaTipo.obter(IdPessoaTipo) Then
			Set objPessoaTipo = objetoPessoaTipo
		End If
	End If

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
	IdPessoa = Null
	BairroPessoa = Null
	CEPPessoa = Null
	CidadePessoa = Null
	ComplementoPontoDeRe = Null
	Contato = Null
	CPFCNPJ = Null
	DataNascimento = Null
	DependenteDe = Null
	Desativado = Null
	Email = Null
	EnderecoPessoa = Null
	FantasiaApelido = Null
	FonePrincipa = Null
	IdEmpresa = Null
	IdPessoaTipo = Null
	NomeRazao = Null
	NumPessoa = Null
	RGIE = Null
	UFPessoa = Null

	booChaveAlterada = False
	booAtributoAlterado = False

Exit_limpar:
	Exit Sub

Err_limpar:
	GoTo Exit_limpar

End Sub

'Método Existe [Com conhecimento de SQL]
'Verifica a existência do objeto Pessoa na tabela correspondente no Banco de Dados
Function existe(argIdPessoa As Variant) As Boolean
On Error GoTo Err_existe

	Dim objCon As New aclConexaoBD
	Dim rstExiste As Recordset
	Dim strSql As String

	existe = False

	strSql = "Select * " & _
		"From Pessoa " & _
		"Where IdPessoa = " & objCon.valorSql(argIdPessoa)

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
		"Pessoa(IdPessoa,BairroPessoa,CEPPessoa,CidadePessoa,ComplementoPontoDeRe,Contato,CPFCNPJ,DataNascimento,DependenteDe,Desativado,Email,EnderecoPessoa,FantasiaApelido,FonePrincipa,IdEmpresa,IdPessoaTipo,NomeRazao,NumPessoa,RGIE,UFPessoa) " & _
		"Values(" & objCon.valorSql(IdPessoa) & "," & objCon.valorSql(BairroPessoa) & "," & objCon.valorSql(CEPPessoa) & "," & objCon.valorSql(CidadePessoa) & "," & objCon.valorSql(ComplementoPontoDeRe) & "," & objCon.valorSql(Contato) & "," & objCon.valorSql(CPFCNPJ) & "," & objCon.valorSql(DataNascimento) & "," & objCon.valorSql(DependenteDe) & "," & objCon.valorSql(Desativado) & "," & objCon.valorSql(Email) & "," & objCon.valorSql(EnderecoPessoa) & "," & objCon.valorSql(FantasiaApelido) & "," & objCon.valorSql(FonePrincipa) & "," & objCon.valorSql(IdEmpresa) & "," & objCon.valorSql(IdPessoaTipo) & "," & objCon.valorSql(NomeRazao) & "," & objCon.valorSql(NumPessoa) & "," & objCon.valorSql(RGIE) & "," & objCon.valorSql(UFPessoa) & ")"

	incluir = (objCon.executa(strSql) > 0)

	If incluir Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza os campos de backup
		bkpIdPessoa = IdPessoa
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

	strSql = "Delete From Pessoa " & _
		"Where IdPessoa = " & objCon.valorSql(IdPessoa)

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
'Recupera o objeto Pessoa através dos argumentos informados
Function obter(argIdPessoa As Variant) As Boolean
On Error GoTo Err_obter

	Dim objCon As New aclConexaoBD
	Dim rstObter As Recordset
	Dim strSql As String

	strSql = "Select * " & _
		"From Pessoa " & _
		"Where IdPessoa = " & objCon.valorSql(argIdPessoa)

	Set rstObter = objCon.consulta(strSql)

	If rstObter.RecordCount = 0 Then
		obter = False
		Exit Function
	End If

	'Atualiza os campos de backup e os identificadores
	IdPessoa = argIdPessoa
	bkpIdPessoa = argIdPessoa

	'Atualiza os campos restantes
	BairroPessoa = rstObter.Fields("BairroPessoa")
	CEPPessoa = rstObter.Fields("CEPPessoa")
	CidadePessoa = rstObter.Fields("CidadePessoa")
	ComplementoPontoDeRe = rstObter.Fields("ComplementoPontoDeRe")
	Contato = rstObter.Fields("Contato")
	CPFCNPJ = rstObter.Fields("CPFCNPJ")
	DataNascimento = rstObter.Fields("DataNascimento")
	DependenteDe = rstObter.Fields("DependenteDe")
	Desativado = rstObter.Fields("Desativado")
	Email = rstObter.Fields("Email")
	EnderecoPessoa = rstObter.Fields("EnderecoPessoa")
	FantasiaApelido = rstObter.Fields("FantasiaApelido")
	FonePrincipa = rstObter.Fields("FonePrincipa")
	IdEmpresa = rstObter.Fields("IdEmpresa")
	IdPessoaTipo = rstObter.Fields("IdPessoaTipo")
	NomeRazao = rstObter.Fields("NomeRazao")
	NumPessoa = rstObter.Fields("NumPessoa")
	RGIE = rstObter.Fields("RGIE")
	UFPessoa = rstObter.Fields("UFPessoa")

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

	If existe(bkpIdPessoa) Then
		strSql = "Update Pessoa " & _
			"Set IdPessoa = " & objCon.valorSql(IdPessoa) & ", BairroPessoa = " & objCon.valorSql(BairroPessoa) & ", CEPPessoa = " & objCon.valorSql(CEPPessoa) & ", CidadePessoa = " & objCon.valorSql(CidadePessoa) & ", ComplementoPontoDeRe = " & objCon.valorSql(ComplementoPontoDeRe) & ", Contato = " & objCon.valorSql(Contato) & ", CPFCNPJ = " & objCon.valorSql(CPFCNPJ) & ", DataNascimento = " & objCon.valorSql(DataNascimento) & ", DependenteDe = " & objCon.valorSql(DependenteDe) & ", Desativado = " & objCon.valorSql(Desativado) & ", Email = " & objCon.valorSql(Email) & ", EnderecoPessoa = " & objCon.valorSql(EnderecoPessoa) & ", FantasiaApelido = " & objCon.valorSql(FantasiaApelido) & ", FonePrincipa = " & objCon.valorSql(FonePrincipa) & ", IdEmpresa = " & objCon.valorSql(IdEmpresa) & ", IdPessoaTipo = " & objCon.valorSql(IdPessoaTipo) & ", NomeRazao = " & objCon.valorSql(NomeRazao) & ", NumPessoa = " & objCon.valorSql(NumPessoa) & ", RGIE = " & objCon.valorSql(RGIE) & ", UFPessoa = " & objCon.valorSql(UFPessoa) & _
			" Where IdPessoa = " & objCon.valorSql(bkpIdPessoa)

		salvar = (objCon.executa(strSql) > 0)
	Else
		salvar = incluir
	End If

	If salvar Then
		booChaveAlterada = False
		booAtributoAlterado = False

		'Atualiza as variáveis de backup com o novo valor da chave
		bkpIdPessoa = IdPessoa
	End If

Exit_salvar:
	Exit Function

Err_salvar:
	salvar = False
	Goto Exit_salvar

End Function

'Fim da classe...