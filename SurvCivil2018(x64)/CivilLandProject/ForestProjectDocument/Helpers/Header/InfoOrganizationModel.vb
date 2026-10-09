Imports System.Xml.Serialization

''' <summary>
''' Сведения о юридическом лице (InfoOrganizationEType)
''' </summary>
Public Class InfoOrganizationModel

    <XmlElement("name")>
    Public Property Name As String

    <XmlElement("regDate")>
    Public Property RegDate As String

    <XmlElement("inn")>
    Public Property Inn As String

    <XmlElement("codeOGRN")>
    Public Property CodeOGRN As String

    <XmlElement("location")>
    Public Property Location As LocationModel

End Class
