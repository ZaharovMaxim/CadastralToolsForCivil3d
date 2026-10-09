Imports System.Xml.Serialization

''' <summary>
''' Сведения о лесном квартале (ForestQuarterType)
''' </summary>
Public Class ForestQuarterModel

    <XmlElement("characteristicsForestQuarter")>
    Public Property CharacteristicsForestQuarter As CharacteristicsForestQuarterModel

    <XmlElement("containForestQuarter")>
    Public Property ContainForestQuarter As ContainForestQuarterModel

End Class
