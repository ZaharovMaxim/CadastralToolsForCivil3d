Imports System.Xml.Serialization

''' <summary>
''' Сведения о геодезической основе (InfoGeodeticBaseEType)
''' </summary>
Public Class InfoGeodeticBaseModel

    <XmlElement("tGeoNetwork")>
    Public Property TGeoNetwork As String

    <XmlElement("nameGeoNetworkPoint")>
    Public Property NameGeoNetworkPoint As String

    <XmlElement("geoNetworkSystem")>
    Public Property GeoNetworkSystem As KindEModel

    <XmlElement("xCoordinate")>
    Public Property XCoordinate As String

    <XmlElement("yCoordinate")>
    Public Property YCoordinate As String

    <XmlElement("dExaminate")>
    Public Property DExaminate As String

    <XmlElement("geoNetworkPointInfStatus")>
    Public Property GeoNetworkPointInfStatus As GeoNetworkPointInfStatusModel

End Class
