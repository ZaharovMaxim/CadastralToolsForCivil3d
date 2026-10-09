Imports System.Xml.Serialization

''' <summary>
''' Одна координата (OrdinateType)
''' </summary>
Public Class OrdinateModel

    <XmlElement("x")>
    Public Property X As Decimal

    <XmlElement("y")>
    Public Property Y As Decimal

    <XmlElement("z")>
    Public Property Z As Decimal?

    <XmlElement("ordNmb")>
    Public Property OrdNmb As Decimal?

    <XmlElement("geopointOpred")>
    Public Property GeopointOpred As KindEModel

    <XmlElement("deltaGeopoint")>
    Public Property DeltaGeopoint As Decimal?

    <XmlElement("spatialsElementId")>
    Public Property SpatialsElementId As Decimal?

End Class
