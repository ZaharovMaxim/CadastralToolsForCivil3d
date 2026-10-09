Imports System.Xml.Serialization

''' <summary>
''' Сведения о растении (BerryType)
''' </summary>
Public Class BerryModel

    <XmlElement("forestResours")>
    Public Property ForestResours As KindEModel

    <XmlElement("perCoating")>
    Public Property PerCoating As ValueUnitMeasurementModel

End Class
