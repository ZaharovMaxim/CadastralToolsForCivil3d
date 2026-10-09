Public Class NewTaxationPiece
    Public Property LuForestTaxationNumber As String
    Public Property ObjectGuid As String
    Public Property Area As Decimal?
    Public Property Events As Boolean?
    Public Property PurposeForests As String
    Public Property NameProtectedForests As String
    Public Property RemoveProtectedForests As Boolean?
End Class

Public Class EditTaxationData
    Public Property UseFgislk As Boolean
    Public Property FgislkParentRegistrationNo As String

    Public Property SubforestryCode As String
    Public Property NameTract As String
    Public Property LuForestQuarterNumber As String
    Public Property LuForestTaxationNumber As String

    Public Property Pieces As New List(Of NewTaxationPiece)
End Class

Public Class ComboItem
    Public Property Code As String
    Public Property Description As String

    Public Sub New(code As String, description As String)
        Me.Code = code
        Me.Description = description
    End Sub

    Public Overrides Function ToString() As String
        Return Description
    End Function
End Class
