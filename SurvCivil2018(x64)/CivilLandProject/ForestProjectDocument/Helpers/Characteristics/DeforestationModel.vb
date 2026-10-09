Imports System.Xml.Serialization

''' <summary>
''' Сведения о вырубках (DeforestationType)
''' </summary>
Public Class DeforestationModel

    <XmlElement("quantityStumps")>
    Public Property QuantityStumps As ValueUnitMeasurementModel

    <XmlElement("pineQuantityStumps")>
    Public Property PineQuantityStumps As ValueUnitMeasurementModel

    <XmlElement("averageDiameterStumps")>
    Public Property AverageDiameterStumps As ValueUnitMeasurementModel

    <XmlElement("typeDeforestation")>
    Public Property TypeDeforestation As KindEModel

    <XmlElement("deforestationYear")>
    Public Property DeforestationYear As String

End Class
