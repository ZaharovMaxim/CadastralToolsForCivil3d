Imports System.Windows.Forms
Imports CadastralTools.Models
Imports CadastralTools.Services
Public Class TaxationDescriptionForm
    ''' <summary>
    ''' Текущая модель документа
    ''' </summary>
    Private _model As TaxationDescriptionModel
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        _model = New TaxationDescriptionModel()
        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    ''' <summary>
    ''' Загружает модель в UI (все вкладки)
    ''' </summary>
    Private Sub LoadModelToUI()
        ' Пока пусто — заполним по мере создания вкладок
    End Sub

    ''' <summary>
    ''' Сохраняет данные из UI в модель
    ''' </summary>
    Private Sub SaveUIToModel()
        ' Пока пусто — заполним по мере создания вкладок
    End Sub

    Private Sub BtnLoadXml_Click(sender As Object, e As EventArgs) Handles BtnLoadXml.Click
        Using dlg As New OpenFileDialog()
            dlg.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
            If dlg.ShowDialog() = DialogResult.OK Then
                Try
                    _model = TaxationDescriptionSerializer.Load(dlg.FileName)
                    LoadModelToUI()
                    MessageBox.Show("Загружено успешно.", "Загрузка",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Ошибка загрузки: " & ex.Message, "Ошибка",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Private Sub BtnSaveXml_Click(sender As Object, e As EventArgs) Handles BtnSaveXml.Click
        SaveUIToModel()

        Using dlg As New SaveFileDialog()
            dlg.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
            If dlg.ShowDialog() = DialogResult.OK Then
                Try
                    TaxationDescriptionSerializer.Save(_model, dlg.FileName)
                    MessageBox.Show("Сохранено успешно.", "Сохранение",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Ошибка сохранения: " & ex.Message, "Ошибка",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles BtnClose.Click
        Close()
    End Sub

End Class
