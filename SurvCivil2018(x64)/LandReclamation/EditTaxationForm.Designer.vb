<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EditTaxationForm
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.grpMode = New System.Windows.Forms.GroupBox()
        Me.rbLu = New System.Windows.Forms.RadioButton()
        Me.rbFgislk = New System.Windows.Forms.RadioButton()
        Me.grpParent = New System.Windows.Forms.GroupBox()
        Me.txtLuTaxation = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtLuQuarter = New System.Windows.Forms.TextBox()
        Me.txtNameTract = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbSubforestry = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lbFgislkParent = New System.Windows.Forms.Label()
        Me.txtFgislkParent = New System.Windows.Forms.TextBox()
        Me.grpPieces = New System.Windows.Forms.GroupBox()
        Me.btnRemovePiece = New System.Windows.Forms.Button()
        Me.btnAddPiece = New System.Windows.Forms.Button()
        Me.dgvPieces = New System.Windows.Forms.DataGridView()
        Me.LuForestTaxationNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ObjectGuid = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Area = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Events = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PurposeForests = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.NameProtectedForests = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.RemoveProtectedForests = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.grpMode.SuspendLayout()
        Me.grpParent.SuspendLayout()
        Me.grpPieces.SuspendLayout()
        CType(Me.dgvPieces, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpMode
        '
        Me.grpMode.Controls.Add(Me.rbLu)
        Me.grpMode.Controls.Add(Me.rbFgislk)
        Me.grpMode.Location = New System.Drawing.Point(6, 13)
        Me.grpMode.Name = "grpMode"
        Me.grpMode.Size = New System.Drawing.Size(376, 76)
        Me.grpMode.TabIndex = 0
        Me.grpMode.TabStop = False
        Me.grpMode.Text = "Способ идентификации"
        '
        'rbLu
        '
        Me.rbLu.AutoSize = True
        Me.rbLu.Location = New System.Drawing.Point(9, 49)
        Me.rbLu.Name = "rbLu"
        Me.rbLu.Size = New System.Drawing.Size(188, 17)
        Me.rbLu.TabIndex = 1
        Me.rbLu.TabStop = True
        Me.rbLu.Text = "По лесоустроительному номеру"
        Me.rbLu.UseVisualStyleBackColor = True
        '
        'rbFgislk
        '
        Me.rbFgislk.AutoSize = True
        Me.rbFgislk.Location = New System.Drawing.Point(11, 23)
        Me.rbFgislk.Name = "rbFgislk"
        Me.rbFgislk.Size = New System.Drawing.Size(181, 17)
        Me.rbFgislk.TabIndex = 0
        Me.rbFgislk.TabStop = True
        Me.rbFgislk.Text = "По учётному номеру ФГИС ЛК"
        Me.rbFgislk.UseVisualStyleBackColor = True
        '
        'grpParent
        '
        Me.grpParent.Controls.Add(Me.txtLuTaxation)
        Me.grpParent.Controls.Add(Me.Label4)
        Me.grpParent.Controls.Add(Me.Label3)
        Me.grpParent.Controls.Add(Me.txtLuQuarter)
        Me.grpParent.Controls.Add(Me.txtNameTract)
        Me.grpParent.Controls.Add(Me.Label2)
        Me.grpParent.Controls.Add(Me.cmbSubforestry)
        Me.grpParent.Controls.Add(Me.Label1)
        Me.grpParent.Controls.Add(Me.lbFgislkParent)
        Me.grpParent.Controls.Add(Me.txtFgislkParent)
        Me.grpParent.Location = New System.Drawing.Point(8, 102)
        Me.grpParent.Name = "grpParent"
        Me.grpParent.Size = New System.Drawing.Size(373, 227)
        Me.grpParent.TabIndex = 1
        Me.grpParent.TabStop = False
        Me.grpParent.Text = "Данные исходного выдела"
        '
        'txtLuTaxation
        '
        Me.txtLuTaxation.Location = New System.Drawing.Point(9, 193)
        Me.txtLuTaxation.Name = "txtLuTaxation"
        Me.txtLuTaxation.Size = New System.Drawing.Size(351, 20)
        Me.txtLuTaxation.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(14, 177)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Номер выдела"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 138)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(91, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Номер квартала"
        '
        'txtLuQuarter
        '
        Me.txtLuQuarter.Location = New System.Drawing.Point(9, 154)
        Me.txtLuQuarter.Name = "txtLuQuarter"
        Me.txtLuQuarter.Size = New System.Drawing.Size(351, 20)
        Me.txtLuQuarter.TabIndex = 6
        '
        'txtNameTract
        '
        Me.txtNameTract.Location = New System.Drawing.Point(9, 115)
        Me.txtNameTract.Name = "txtNameTract"
        Me.txtNameTract.Size = New System.Drawing.Size(352, 20)
        Me.txtNameTract.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 99)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Урочище"
        '
        'cmbSubforestry
        '
        Me.cmbSubforestry.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSubforestry.FormattingEnabled = True
        Me.cmbSubforestry.Location = New System.Drawing.Point(12, 75)
        Me.cmbSubforestry.Name = "cmbSubforestry"
        Me.cmbSubforestry.Size = New System.Drawing.Size(348, 21)
        Me.cmbSubforestry.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(134, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Участковое лесничество"
        '
        'lbFgislkParent
        '
        Me.lbFgislkParent.AutoSize = True
        Me.lbFgislkParent.Location = New System.Drawing.Point(14, 17)
        Me.lbFgislkParent.Name = "lbFgislkParent"
        Me.lbFgislkParent.Size = New System.Drawing.Size(139, 13)
        Me.lbFgislkParent.TabIndex = 1
        Me.lbFgislkParent.Text = "Учётный номер ФГИС ЛК"
        '
        'txtFgislkParent
        '
        Me.txtFgislkParent.Location = New System.Drawing.Point(12, 36)
        Me.txtFgislkParent.Name = "txtFgislkParent"
        Me.txtFgislkParent.Size = New System.Drawing.Size(351, 20)
        Me.txtFgislkParent.TabIndex = 0
        '
        'grpPieces
        '
        Me.grpPieces.Controls.Add(Me.btnRemovePiece)
        Me.grpPieces.Controls.Add(Me.btnAddPiece)
        Me.grpPieces.Controls.Add(Me.dgvPieces)
        Me.grpPieces.Location = New System.Drawing.Point(12, 344)
        Me.grpPieces.Name = "grpPieces"
        Me.grpPieces.Size = New System.Drawing.Size(817, 237)
        Me.grpPieces.TabIndex = 2
        Me.grpPieces.TabStop = False
        Me.grpPieces.Text = "Новые выделы"
        '
        'btnRemovePiece
        '
        Me.btnRemovePiece.Location = New System.Drawing.Point(241, 206)
        Me.btnRemovePiece.Name = "btnRemovePiece"
        Me.btnRemovePiece.Size = New System.Drawing.Size(90, 29)
        Me.btnRemovePiece.TabIndex = 2
        Me.btnRemovePiece.Text = "Удалить"
        Me.btnRemovePiece.UseVisualStyleBackColor = True
        '
        'btnAddPiece
        '
        Me.btnAddPiece.Location = New System.Drawing.Point(18, 209)
        Me.btnAddPiece.Name = "btnAddPiece"
        Me.btnAddPiece.Size = New System.Drawing.Size(101, 27)
        Me.btnAddPiece.TabIndex = 1
        Me.btnAddPiece.Text = "Добавить"
        Me.btnAddPiece.UseVisualStyleBackColor = True
        '
        'dgvPieces
        '
        Me.dgvPieces.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPieces.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.LuForestTaxationNumber, Me.ObjectGuid, Me.Area, Me.Events, Me.PurposeForests, Me.NameProtectedForests, Me.RemoveProtectedForests})
        Me.dgvPieces.Location = New System.Drawing.Point(10, 25)
        Me.dgvPieces.Name = "dgvPieces"
        Me.dgvPieces.Size = New System.Drawing.Size(747, 172)
        Me.dgvPieces.TabIndex = 0
        '
        'LuForestTaxationNumber
        '
        Me.LuForestTaxationNumber.HeaderText = "Лесоустр. номер"
        Me.LuForestTaxationNumber.Name = "LuForestTaxationNumber"
        '
        'ObjectGuid
        '
        Me.ObjectGuid.HeaderText = "objectGuid"
        Me.ObjectGuid.Name = "ObjectGuid"
        '
        'Area
        '
        Me.Area.HeaderText = "Площадь, га"
        Me.Area.Name = "Area"
        '
        'Events
        '
        Me.Events.HeaderText = "events (true/false)"
        Me.Events.Name = "Events"
        '
        'PurposeForests
        '
        Me.PurposeForests.HeaderText = "PurposeForests"
        Me.PurposeForests.Name = "PurposeForests"
        '
        'NameProtectedForests
        '
        Me.NameProtectedForests.HeaderText = "NameProtectedForests"
        Me.NameProtectedForests.Name = "NameProtectedForests"
        '
        'RemoveProtectedForests
        '
        Me.RemoveProtectedForests.HeaderText = "removeProtectedForests (true/false)"
        Me.RemoveProtectedForests.Name = "RemoveProtectedForests"
        '
        'btnGenerate
        '
        Me.btnGenerate.Location = New System.Drawing.Point(266, 595)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(101, 19)
        Me.btnGenerate.TabIndex = 3
        Me.btnGenerate.Text = "Сформировать XML"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(19, 596)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(121, 18)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Отмена"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'EditTaxationForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(841, 623)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnGenerate)
        Me.Controls.Add(Me.grpPieces)
        Me.Controls.Add(Me.grpParent)
        Me.Controls.Add(Me.grpMode)
        Me.Name = "EditTaxationForm"
        Me.Text = "EditTaxationForm"
        Me.grpMode.ResumeLayout(False)
        Me.grpMode.PerformLayout()
        Me.grpParent.ResumeLayout(False)
        Me.grpParent.PerformLayout()
        Me.grpPieces.ResumeLayout(False)
        CType(Me.dgvPieces, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents grpMode As Windows.Forms.GroupBox
    Friend WithEvents rbLu As Windows.Forms.RadioButton
    Friend WithEvents rbFgislk As Windows.Forms.RadioButton
    Friend WithEvents grpParent As Windows.Forms.GroupBox
    Friend WithEvents txtLuTaxation As Windows.Forms.TextBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents txtLuQuarter As Windows.Forms.TextBox
    Friend WithEvents txtNameTract As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents cmbSubforestry As Windows.Forms.ComboBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents lbFgislkParent As Windows.Forms.Label
    Friend WithEvents txtFgislkParent As Windows.Forms.TextBox
    Friend WithEvents grpPieces As Windows.Forms.GroupBox
    Friend WithEvents dgvPieces As Windows.Forms.DataGridView
    Friend WithEvents btnRemovePiece As Windows.Forms.Button
    Friend WithEvents btnAddPiece As Windows.Forms.Button
    Friend WithEvents LuForestTaxationNumber As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ObjectGuid As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Area As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Events As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PurposeForests As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents NameProtectedForests As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents RemoveProtectedForests As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnGenerate As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
End Class
