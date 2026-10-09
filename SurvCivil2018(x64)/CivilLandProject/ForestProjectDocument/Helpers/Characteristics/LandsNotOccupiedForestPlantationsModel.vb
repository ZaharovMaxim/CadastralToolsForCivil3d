Imports System.Xml.Serialization

''' <summary>
''' Сведения о не занятых лесными насаждениями землях
''' (LandsNotOccupiedForestPlantationsType)
''' </summary>
Public Class LandsNotOccupiedForestPlantationsModel

    <XmlElement("treeSpecies")>
    Public Property TreeSpecies As SpeciesModel

    <XmlElement("shrubbySpecies")>
    Public Property ShrubbySpecies As SpeciesModel

    <XmlElement("characteristicsNaturalReforestation")>
    Public Property CharacteristicsNaturalReforestation As CharacteristicsNaturalReforestationModel

End Class
