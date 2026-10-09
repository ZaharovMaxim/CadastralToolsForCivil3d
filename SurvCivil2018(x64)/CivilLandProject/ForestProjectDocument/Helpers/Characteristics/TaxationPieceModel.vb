Imports System.Xml.Serialization

''' <summary>
''' Сведения о лесотаксационном выделе (TaxationPieceType)
''' </summary>
Public Class TaxationPieceModel

    <XmlElement("characteristicsTaxationPiece")>
    Public Property CharacteristicsTaxationPiece As CharacteristicsTaxationPieceModel

End Class