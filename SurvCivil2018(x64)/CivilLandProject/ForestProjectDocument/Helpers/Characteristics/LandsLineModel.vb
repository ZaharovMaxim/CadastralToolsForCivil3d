Imports System.Xml.Serialization

''' <summary>
''' Сведения о землях линейного протяжения (LandsLineType)
''' </summary>
Public Class LandsLineModel

    <XmlElement("typeRoad")>
    Public Property TypeRoad As KindEModel

    <XmlElement("width")>
    Public Property Width As ValueUnitMeasurementModel

    <XmlElement("length")>
    Public Property Length As ValueUnitMeasurementModel

    <XmlElement("state")>
    Public Property State As String

    <XmlElement("purpose")>
    Public Property Purpose As KindEModel

    <XmlElement("typeCoating")>
    Public Property TypeCoating As String

    <XmlElement("roadwayWidth")>
    Public Property RoadwayWidth As ValueUnitMeasurementModel

    <XmlElement("seasonality")>
    Public Property Seasonality As KindEModel

    <XmlElement("events")>
    Public Property Events As String

End Class
