Imports System.Xml.Serialization

''' <summary>
''' Сведения об учреждении (InstituteEType)
''' </summary>
Public Class InstituteModel

    <XmlElement("name")>
    Public Property Name As String

    <XmlElement("location")>
    Public Property Location As LocationModel

    <XmlElement("inn")>
    Public Property Inn As String

    <XmlElement("codeOGRN")>
    Public Property CodeOGRN As String

End Class
