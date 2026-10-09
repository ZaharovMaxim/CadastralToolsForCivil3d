Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Характеристики лесотаксационного выдела (CharacteristicsTaxationPieceType).
''' Самый большой блок схемы — содержит 40+ опциональных полей.
''' </summary>
Public Class CharacteristicsTaxationPieceModel

    ' --- Связи ---
    <XmlElement("relations")>
    Public Property Relations As RelationsModel

    ' --- ГЕОМЕТРИЯ ---
    <XmlElement("contoursLocation")>
    Public Property ContoursLocation As PopdContoursLocationModel

    ' --- Номера ---
    <XmlElement("currentAndPreviousNumbers")>
    Public Property CurrentAndPreviousNumbers As CurrentAndPreviousNumbersModel

    ' --- Даты ---
    <XmlElement("dateEndTaxationPiece")>
    Public Property DateEndTaxationPiece As String

    <XmlElement("documentationDate")>
    Public Property DocumentationDate As String

    ' --- Площадь ---
    <XmlElement("area")>
    Public Property Area As AreaModel

    <XmlElement("calcError")>
    Public Property CalcError As Decimal?

    ' --- Категории земель ---
    <XmlElement("categoryLand")>
    Public Property CategoryLand As KindEModel

    <XmlElement("totalStockForest")>
    Public Property TotalStockForest As ValueUnitMeasurementModel

    <XmlElement("registrationNumberCuttingArea")>
    Public Property RegistrationNumbersCuttingArea As List(Of String)

    <XmlElement("typeLand")>
    Public Property TypeLand As KindEModel

    <XmlElement("typeForestLand")>
    Public Property TypeForestLand As KindEModel

    <XmlElement("typeNonForestLand")>
    Public Property TypeNonForestLand As KindEModel

    ' --- Лес ---
    <XmlElement("typeForest")>
    Public Property TypeForest As KindEModel

    <XmlElement("typeForestGrowingConditions")>
    Public Property TypeForestGrowingConditions As KindEModel

    <XmlElement("stateForestPlantation")>
    Public Property StateForestPlantation As KindEModel

    <XmlElement("targetSpecies")>
    Public Property TargetSpecies As KindEModel

    <XmlElement("dominantSpecies")>
    Public Property DominantSpecies As KindEModel

    <XmlElement("bonusClass")>
    Public Property BonusClass As KindEModel

    ' --- Запас ---
    <XmlElement("stock")>
    Public Property Stock As StockModel

    ' --- Годы, секции ---
    <XmlElement("year")>
    Public Property Year As String

    <XmlElement("economSection")>
    Public Property EconomSection As String

    <XmlElement("cutAge")>
    Public Property CutAge As String

    ' --- Ярусы и подлесок ---
    <XmlElement("forestPlantings")>
    Public Property ForestPlantings As ForestPlantingsModel

    <XmlElement("nonClosedForestCrops")>
    Public Property NonClosedForestCrops As NonClosedForestCropsModel

    <XmlElement("treeUndergrowth")>
    Public Property TreeUndergrowth As TreeUndergrowthModel

    <XmlElement("forestUndergrowth")>
    Public Property ForestUndergrowth As ForestUndergrowthModel

    ' --- Прочие блоки ---
    <XmlElement("landsNotOccupiedForestPlantations")>
    Public Property LandsNotOccupiedForestPlantations As LandsNotOccupiedForestPlantationsModel

    <XmlElement("deforestation")>
    Public Property Deforestation As DeforestationModel

    <XmlElement("singleTrees")>
    Public Property SingleTrees As SingleTreesModel

    <XmlElement("lostForestPlantations")>
    Public Property LostForestPlantations As LostForestPlantationsModel

    <XmlElement("haymaking")>
    Public Property Haymaking As HaymakingModel

    <XmlElement("gardensVineyardsPlants")>
    Public Property GardensVineyardsPlants As GardensVineyardsPlantsModel

    <XmlElement("swamps")>
    Public Property Swamps As SwampsModel

    <XmlElement("positionForestTaxingAllocation")>
    Public Property PositionForestTaxingAllocation As PositionForestTaxingAllocationModel

    ' --- Разное ---
    <XmlElement("nameObjectForestSeedProduction")>
    Public Property NameObjectForestSeedProduction As KindEModel

    <XmlElement("radionuclideContamination")>
    Public Property RadionuclideContamination As KindEModel

    <XmlElement("events")>
    Public Property Events As EventsModel

    <XmlElement("protectedForests")>
    Public Property ProtectedForests As KindEModel

    <XmlElement("purposeForests")>
    Public Property PurposeForests As KindEModel

    <XmlElement("farmCategory")>
    Public Property FarmCategory As KindEModel

    <XmlElement("relief")>
    Public Property Relief As KindEModel

    <XmlElement("lastManageYear")>
    Public Property LastManageYear As String

    <XmlElement("admRegion")>
    Public Property AdmRegion As String

    <XmlElement("peculiarities")>
    Public Property Peculiarities As String

    <XmlElement("forestCrops")>
    Public Property ForestCrops As ForestCropsModel

    <XmlElement("landsLine")>
    Public Property LandsLine As LandsLineModel

    <XmlElement("tapping")>
    Public Property Tapping As String

    <XmlElement("landscape")>
    Public Property Landscape As String

    <XmlElement("soilCharacteristics")>
    Public Property SoilCharacteristics As SoilCharacteristicsModel

    <XmlElement("schoolForest")>
    Public Property SchoolForest As ForestCropsModel

    <XmlElement("berries")>
    Public Property Berries As BerriesModel

    <XmlElement("nonWood")>
    Public Property NonWood As NonWoodModel

    <XmlElement("selection")>
    Public Property Selection As String

End Class
