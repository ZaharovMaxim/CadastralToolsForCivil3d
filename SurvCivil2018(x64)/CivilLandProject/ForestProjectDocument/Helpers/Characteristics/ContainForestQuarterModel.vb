Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Информация о содержимом лесного квартала (ContainForestQuarterType)
''' </summary>
Public Class ContainForestQuarterModel

    <XmlElement("taxationPiece")>
    Public Property TaxationPieces As List(Of TaxationPieceModel)

End Class
