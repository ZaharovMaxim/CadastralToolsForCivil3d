Imports System.Xml.Serialization

''' <summary>
''' Контур (ContourType).
''' По нашему соглашению: один выдел = один контур.
''' </summary>
Public Class ContourModel

    <XmlElement("numberPp")>
    Public Property NumberPp As Decimal?

    <XmlElement("area")>
    Public Property Area As AreaModel

    <XmlElement("entitySpatial")>
    Public Property EntitySpatial As EntitySpatialModel

    <XmlElement("definition")>
    Public Property Definition As String

End Class
