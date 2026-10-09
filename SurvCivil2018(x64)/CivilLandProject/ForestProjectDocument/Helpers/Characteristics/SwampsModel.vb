Imports System.Xml.Serialization

''' <summary>
''' Сведения о болотах (SwampsType)
''' </summary>
Public Class SwampsModel

    <XmlElement("swampType")>
    Public Property SwampType As KindEModel

    <XmlElement("typeVegetation")>
    Public Property TypeVegetation As KindEModel

    <XmlElement("thicknessLayer")>
    Public Property ThicknessLayer As ValueUnitMeasurementModel

    <XmlElement("tree")>
    Public Property Tree As KindEModel

    <XmlElement("perOvergrowth")>
    Public Property PerOvergrowth As ValueUnitMeasurementModel

End Class
