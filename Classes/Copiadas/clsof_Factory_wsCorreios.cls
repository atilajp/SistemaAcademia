VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsof_Factory_wsCorreios"
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
'The Generic Type Mapper is a SOAP Toolkit custom type mapper implementation
'that can map (serialize to XML and deserialize from XML) most of a user's COM
'objects. For more information, see Object Factory in Microsoft Office 2003
'Web Services Toolkit Help.
'
'Changes to the code in this class may result in incorrect behavior.
'
'*****************************************************************

Implements IGCTMObjectFactory

Private Function IGCTMObjectFactory_CreateObject(ByVal par_WSMLNode As MSXML2.IXMLDOMNode) As Object
    Dim node As IXMLDOMNode

    On Error GoTo IGCTMObjectFactoryTrap

    Set node = par_WSMLNode.Attributes.getNamedItem("targetClassName")

    Set IGCTMObjectFactory_CreateObject = Nothing

    If Not (node Is Nothing) Then
        Select Case node.nodeValue
            Case "struct_ModeloCEP"
                Set IGCTMObjectFactory_CreateObject = New struct_ModeloCEP
        End Select
    End If

Exit Function

IGCTMObjectFactoryTrap:
    Err.Raise Err.Number, "clsof_Factory_wsCorreios", Err.Description
End Function


