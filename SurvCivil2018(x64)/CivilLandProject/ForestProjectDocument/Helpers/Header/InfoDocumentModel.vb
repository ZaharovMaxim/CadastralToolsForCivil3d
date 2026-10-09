Imports System.Xml.Serialization

''' <summary>
''' Сведения о документе (InfoDocumentType)
''' </summary>
Public Class InfoDocumentModel

    <XmlElement("organizationName")>
    Public Property OrganizationName As String

    <XmlElement("documentName")>
    Public Property DocumentName As String

    <XmlElement("documentNumber")>
    Public Property DocumentNumber As String

    <XmlElement("documentData")>
    Public Property DocumentData As String

    <XmlElement("sign")>
    Public Property Sign As SignerDataModel

End Class