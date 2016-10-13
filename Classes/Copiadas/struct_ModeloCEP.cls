VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "struct_ModeloCEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

'*****************************************************************
'This class was created by the Microsoft Office 2003 Web Services Toolkit.
'
'Created: 1/14/2013 06:25:24 PM
'
'Description:
'This class is a Visual Basic for Applications class representation of the user-defined
'type as defined by http://softcomdevelop.com.br/webService/wsCorreios.asmx?wsdl.
'
'This class only contains the ModeloCEP,
'as defined in the WSDL.
'
'Changes to the code in this class may result in incorrect behavior.
'
'*****************************************************************

Public Status As Long
Public StatusDesc As String
Public Logradouro As String
Public Bairro As String
Public Cidade As String
Public cCidade As String
Public UF As String
Public xUF As String
Public cUF As String
Public Pais As String
Public cPais As String
Public CEP As String

