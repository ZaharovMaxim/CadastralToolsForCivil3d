Imports System.Xml.Linq

Public Class XmlNamespaces
    Public Shared ReadOnly ProvisionChanges As XNamespace = "http://rosleshoz.gov.ru/fgis-lk/provision-changes-boundaries-forest-objects/1.1.0"
    ' === Common / simple / complex ===
    Public Shared ReadOnly St_20 As XNamespace = "http://rosleshoz.gov.ru/fgis-lk/common/simple/types/2.0"
    Public Shared ReadOnly Cmplx As XNamespace = "http://rosleshoz.gov.ru/fgis-lk/common/complex/types/2.0"

    ' === Subject ===
    Public Shared ReadOnly Subj As XNamespace = "http://rosleshoz.gov.ru/fgis-lk/common/subject/types/2.1"

    ' === Extract ===
    Public Shared ReadOnly SfrExtract As XNamespace = "http://rosleshoz.gov.ru/fgis-lk/common/stateForestRegisterExtract/types/2.1"

    ' === Служебные ===
    Public Shared ReadOnly Xsi As XNamespace = "http://www.w3.org/2001/XMLSchema-instance"

    ' Запрещаем создавать экземпляры — это чисто хранилище констант
    Private Sub New()
    End Sub

End Class





