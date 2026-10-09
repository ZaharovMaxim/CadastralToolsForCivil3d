Imports System.Xml.Serialization

''' <summary>
''' Таксационное описание лесотаксационного выдела (TaxationDescriptionType).
''' Корневой элемент XML-документа.
''' </summary>
<XmlRoot("taxationDescription",
         Namespace:="http://rosleshoz.gov.ru/fgis-lk/taxationDescription/types/2.0.0")>
Public Class TaxationDescriptionModel

    <XmlElement("header")>
    Public Property Header As HeaderModel

    ' ВНИМАНИЕ: в XSD имя элемента "initialData " с пробелом на конце — это баг схемы.
    ' В реальном XML элемент называется "initialData". Используем корректное имя.
    <XmlElement("initialData")>
    Public Property InitialData As InitialDataModel

    <XmlElement("characteristicsInfo")>
    Public Property CharacteristicsInfo As ForestQuartersModel

    <XmlElement("explanatoryNote")>
    Public Property ExplanatoryNote As String

End Class
