Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Сведения о характеристиках (ForestQuartersType).
''' Контейнер верхнего уровня для списка кварталов.
''' </summary>
Public Class ForestQuartersModel

    <XmlElement("forestQuarter")>
    Public Property ForestQuarters As List(Of ForestQuarterModel)

End Class
