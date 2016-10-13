VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsws_wsCorreios"
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
'This class is a Visual Basic for Applications class representation of the Web service
'as defined by http://softcomdevelop.com.br/webService/wsCorreios.asmx?wsdl.
'
'To Use:
'Dimension a variable as new clsws_wsCorreios, and then write code to
'use the methods provided by the class.
'Example:
' Dim ExampleVar as New clsws_wsCorreios
' debug.print ExampleVar.wsm_getEnderecoByCEP("Sample Input")
'
'For more information, see Complex Types in Microsoft Office 2003
'Web Services Toolkit Help.
'
'Changes to the code in this class may result in incorrect behavior.
'
'*****************************************************************

'Dimensioning private class variables.
Private sc_wsCorreios As SoapClient30
Private Const c_WSDL_URL As String = "http://softcomdevelop.com.br/webService/wsCorreios.asmx?wsdl"
Private Const c_SERVICE As String = "wsCorreios"
Private Const c_PORT As String = "wsCorreiosSoap"
Private Const c_SERVICE_NAMESPACE As String = "http://tempuri.org/"

Private Sub Class_Initialize()
    '*****************************************************************
    'This subroutine will be called each time the class is instantiated.
    'Creates sc_ComplexTypes as new SoapClient30, and then
    'initializes sc_ComplexTypes.mssoapinit2 with WSDL file found in
    'http://softcomdevelop.com.br/webService/wsCorreios.asmx?wsdl.
    '*****************************************************************

    Dim str_WSML As String
    str_WSML = "<servicemapping>"
    str_WSML = str_WSML & "<service name='wsCorreios'>"
    str_WSML = str_WSML & "<using PROGID='MSOSOAP.GenericCustomTypeMapper30' cachable='0' ID='GCTM'/>"
    str_WSML = str_WSML & "<types>"
    str_WSML = str_WSML & "<type name='ModeloCEP' targetNamespace='http://tempuri.org/' uses='GCTM' targetClassName='struct_ModeloCEP'/>"
    str_WSML = str_WSML & "</types>"
    str_WSML = str_WSML & "</service>"
    str_WSML = str_WSML & "</servicemapping>"

    Set sc_wsCorreios = New SoapClient30

    sc_wsCorreios.MSSoapInit2 c_WSDL_URL, str_WSML, c_SERVICE, c_PORT, c_SERVICE_NAMESPACE
    'Use the proxy server defined in Internet Explorer's LAN settings by
    'setting ProxyServer to <CURRENT_USER>
    sc_wsCorreios.ConnectorProperty("ProxyServer") = "<CURRENT_USER>"
    'Autodetect proxy settings if Internet Explorer is set to autodetect
    'by setting EnableAutoProxy to True
    sc_wsCorreios.ConnectorProperty("EnableAutoProxy") = True

    Set sc_wsCorreios.ClientProperty("GCTMObjectFactory") = New clsof_Factory_wsCorreios

End Sub

Private Sub Class_Terminate()
    '*****************************************************************
    'This subroutine will be called each time the class is destructed.
    'Sets sc_ComplexTypes to Nothing.
    '*****************************************************************

    'Error Trap
    On Error GoTo Class_TerminateTrap

    Set sc_wsCorreios = Nothing

Exit Sub

Class_TerminateTrap:
    wsCorreiosErrorHandler ("Class_Terminate")
End Sub

Private Sub wsCorreiosErrorHandler(str_Function As String)
    '*****************************************************************
    'This subroutine is the class error handler. It can be called from any class subroutine or function
    'when that subroutine or function encounters an error. Then, it will raise the error along with the
    'name of the calling subroutine or function.
    '*****************************************************************

    'SOAP Error
    If sc_wsCorreios.FaultCode <> "" Then
        Err.Raise vbObjectError, str_Function, sc_wsCorreios.FaultString
    'Non SOAP Error
    Else
        Err.Raise Err.Number, str_Function, Err.Description
    End If

End Sub

Public Function wsm_getEnderecoByCEP(ByVal str_CEP As String) As struct_ModeloCEP
    '*****************************************************************
    'Proxy function created from http://softcomdevelop.com.br/webService/wsCorreios.asmx?wsdl.
    '*****************************************************************

    'Error Trap
    On Error GoTo wsm_getEnderecoByCEPTrap

    Set wsm_getEnderecoByCEP = sc_wsCorreios.getEnderecoByCEP(str_CEP)

Exit Function
wsm_getEnderecoByCEPTrap:
    wsCorreiosErrorHandler "wsm_getEnderecoByCEP"
End Function


