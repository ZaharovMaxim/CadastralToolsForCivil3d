Imports System.Xml.Serialization

''' <summary>
''' Характеристика почвы (SoilCharacteristicsType)
''' </summary>
Public Class SoilCharacteristicsModel

    <XmlElement("soilTypes")>
    Public Property SoilTypes As KindEModel

    <XmlElement("mechanicalComp")>
    Public Property MechanicalComp As String

    <XmlElement("humidityDegree")>
    Public Property HumidityDegree As String

    <XmlElement("turfingDegree")>
    Public Property TurfingDegree As String

    <XmlElement("humusThickness")>
    Public Property HumusThickness As String

    <XmlElement("perRocks")>
    Public Property PerRocks As ValueUnitMeasurementModel

End Class
