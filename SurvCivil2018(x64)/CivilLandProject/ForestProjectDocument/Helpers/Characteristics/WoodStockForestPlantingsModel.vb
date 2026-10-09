Imports System.Xml.Serialization

''' <summary>
''' Запас древесины лесного насаждения (WoodStockForestPlantingsType)
''' </summary>
Public Class WoodStockForestPlantingsModel

    <XmlElement("averageStock")>
    Public Property AverageStock As ValueUnitMeasurementModel

    <XmlElement("stock")>
    Public Property Stock As ValueUnitMeasurementModel

End Class
