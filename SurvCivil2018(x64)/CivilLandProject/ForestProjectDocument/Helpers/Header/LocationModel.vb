Imports System.Xml.Serialization

''' <summary>
''' Почтовый адрес (LocationType)
''' </summary>
Public Class LocationModel

    <XmlElement("okato")>
    Public Property Okato As String

    <XmlElement("kladr")>
    Public Property Kladr As String

    <XmlElement("oktmo")>
    Public Property Oktmo As String

    <XmlElement("postalCode")>
    Public Property PostalCode As String

    <XmlElement("regioncode")>
    Public Property Regioncode As String

    <XmlElement("district")>
    Public Property District As String

    <XmlElement("city")>
    Public Property City As String

    <XmlElement("urbanDistrict")>
    Public Property UrbanDistrict As String

    <XmlElement("sovietVillage")>
    Public Property SovietVillage As String

    <XmlElement("locality")>
    Public Property Locality As String

    <XmlElement("street")>
    Public Property Street As String

    <XmlElement("level1")>
    Public Property Level1 As String

    <XmlElement("level2")>
    Public Property Level2 As String

    <XmlElement("level3")>
    Public Property Level3 As String

    <XmlElement("apartment")>
    Public Property Apartment As String

    <XmlElement("note")>
    Public Property Note As String

End Class
