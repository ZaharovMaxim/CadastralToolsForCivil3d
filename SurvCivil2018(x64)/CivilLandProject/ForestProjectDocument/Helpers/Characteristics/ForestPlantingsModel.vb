Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Сведения о ярусах лесного насаждения (ForestPlantingsType)
''' </summary>
Public Class ForestPlantingsModel

    <XmlElement("forestPlanting")>
    Public Property ForestPlantings As List(Of ForestPlantingModel)

End Class
