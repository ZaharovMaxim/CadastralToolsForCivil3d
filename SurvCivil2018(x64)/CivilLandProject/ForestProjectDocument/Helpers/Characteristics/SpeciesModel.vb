Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Сведения о древесных/кустарниковых породах (SpeciesType)
''' </summary>
Public Class SpeciesModel

    <XmlElement("speciesInfo")>
    Public Property SpeciesInfo As List(Of SpeciesInfoModel)

End Class
