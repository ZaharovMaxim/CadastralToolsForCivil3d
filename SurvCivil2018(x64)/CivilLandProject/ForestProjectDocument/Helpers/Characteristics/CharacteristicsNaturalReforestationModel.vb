Imports System.Xml.Serialization

''' <summary>
''' Характеристика естественного лесовосстановления
''' (CharacteristicsNaturalReforestationType)
''' </summary>
Public Class CharacteristicsNaturalReforestationModel

    <XmlElement("averageHeight")>
    Public Property AverageHeight As ValueUnitMeasurementModel

    <XmlElement("averageAge")>
    Public Property AverageAge As ValueUnitMeasurementModel

    <XmlElement("quantityTreeUndergrowth")>
    Public Property QuantityTreeUndergrowth As ValueUnitMeasurementModel

    <XmlElement("compositionTreeUndergrowth")>
    Public Property CompositionTreeUndergrowth As String

End Class
