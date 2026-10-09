Imports System.Xml.Serialization

''' <summary>
''' Сведения о положении лесотаксационного выдела
''' (PositionForestTaxingAllocationType)
''' </summary>
Public Class PositionForestTaxingAllocationModel

    <XmlElement("slopeSteepness")>
    Public Property SlopeSteepness As KindEModel

    <XmlElement("slopeExposure")>
    Public Property SlopeExposure As KindEModel

    <XmlElement("seaHeight")>
    Public Property SeaHeight As ValueUnitMeasurementModel

    <XmlElement("typeErosion")>
    Public Property TypeErosion As KindEModel

    <XmlElement("erosionDegree")>
    Public Property ErosionDegree As KindEModel

End Class
