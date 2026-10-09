Imports System.Xml.Serialization

''' <summary>
''' Элемент контура (SpatialsElementType)
''' </summary>
Public Class SpatialsElementModel

    <XmlElement("typeUnit")>
    Public Property TypeUnit As KindEModel

    <XmlElement("ordinates")>
    Public Property Ordinates As OrdinatesModel

End Class
