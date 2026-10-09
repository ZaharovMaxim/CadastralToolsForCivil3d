Imports System.Xml.Serialization

''' <summary>
''' Файл с геоданными (SpecialFileType)
''' </summary>
Public Class SpecialFileModel

    <XmlElement("filePath")>
    Public Property FilePath As String

    <XmlElement("geopointOpred")>
    Public Property GeopointOpred As KindEModel

    <XmlElement("deltaGeopoint")>
    Public Property DeltaGeopoint As Decimal?

End Class
