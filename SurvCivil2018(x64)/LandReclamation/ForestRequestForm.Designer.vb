<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ForestRequestForm
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
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.lblHint = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tabForestPart = New System.Windows.Forms.TabPage()
        Me.tabOKS = New System.Windows.Forms.TabPage()
        Me.tabForestry = New System.Windows.Forms.TabPage()
        Me.tabSpecialZone = New System.Windows.Forms.TabPage()
        Me.dgvForestPart = New System.Windows.Forms.DataGridView()
        Me.dgvOKS = New System.Windows.Forms.DataGridView()
        Me.dgvForestry = New System.Windows.Forms.DataGridView()
        Me.dgvSpecialZone = New System.Windows.Forms.DataGridView()
        Me.TabControl1.SuspendLayout()
        Me.tabForestPart.SuspendLayout()
        Me.tabOKS.SuspendLayout()
        Me.tabForestry.SuspendLayout()
        Me.tabSpecialZone.SuspendLayout()
        CType(Me.dgvForestPart, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvOKS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvForestry, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSpecialZone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(14, 562)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(90, 22)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Отмена"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(226, 562)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(85, 21)
        Me.btnOk.TabIndex = 2
        Me.btnOk.Text = "ОК"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'lblHint
        '
        Me.lblHint.AutoSize = True
        Me.lblHint.Location = New System.Drawing.Point(39, 543)
        Me.lblHint.Name = "lblHint"
        Me.lblHint.Size = New System.Drawing.Size(39, 13)
        Me.lblHint.TabIndex = 3
        Me.lblHint.Text = "Label1"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tabForestPart)
        Me.TabControl1.Controls.Add(Me.tabOKS)
        Me.TabControl1.Controls.Add(Me.tabForestry)
        Me.TabControl1.Controls.Add(Me.tabSpecialZone)
        Me.TabControl1.Location = New System.Drawing.Point(14, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(533, 528)
        Me.TabControl1.TabIndex = 4
        '
        'tabForestPart
        '
        Me.tabForestPart.Controls.Add(Me.dgvForestPart)
        Me.tabForestPart.Location = New System.Drawing.Point(4, 22)
        Me.tabForestPart.Name = "tabForestPart"
        Me.tabForestPart.Padding = New System.Windows.Forms.Padding(3)
        Me.tabForestPart.Size = New System.Drawing.Size(525, 502)
        Me.tabForestPart.TabIndex = 0
        Me.tabForestPart.Text = "Лесной участок"
        Me.tabForestPart.UseVisualStyleBackColor = True
        '
        'tabOKS
        '
        Me.tabOKS.Controls.Add(Me.dgvOKS)
        Me.tabOKS.Location = New System.Drawing.Point(4, 22)
        Me.tabOKS.Name = "tabOKS"
        Me.tabOKS.Padding = New System.Windows.Forms.Padding(3)
        Me.tabOKS.Size = New System.Drawing.Size(525, 502)
        Me.tabOKS.TabIndex = 1
        Me.tabOKS.Text = "ОКС"
        Me.tabOKS.UseVisualStyleBackColor = True
        '
        'tabForestry
        '
        Me.tabForestry.Controls.Add(Me.dgvForestry)
        Me.tabForestry.Location = New System.Drawing.Point(4, 22)
        Me.tabForestry.Name = "tabForestry"
        Me.tabForestry.Size = New System.Drawing.Size(525, 502)
        Me.tabForestry.TabIndex = 2
        Me.tabForestry.Text = "Лесничество"
        Me.tabForestry.UseVisualStyleBackColor = True
        '
        'tabSpecialZone
        '
        Me.tabSpecialZone.Controls.Add(Me.dgvSpecialZone)
        Me.tabSpecialZone.Location = New System.Drawing.Point(4, 22)
        Me.tabSpecialZone.Name = "tabSpecialZone"
        Me.tabSpecialZone.Size = New System.Drawing.Size(525, 502)
        Me.tabSpecialZone.TabIndex = 3
        Me.tabSpecialZone.Text = "ЗОУИТ"
        Me.tabSpecialZone.UseVisualStyleBackColor = True
        '
        'dgvForestPart
        '
        Me.dgvForestPart.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvForestPart.Location = New System.Drawing.Point(8, 8)
        Me.dgvForestPart.Name = "dgvForestPart"
        Me.dgvForestPart.Size = New System.Drawing.Size(516, 493)
        Me.dgvForestPart.TabIndex = 0
        '
        'dgvOKS
        '
        Me.dgvOKS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOKS.Location = New System.Drawing.Point(5, 6)
        Me.dgvOKS.Name = "dgvOKS"
        Me.dgvOKS.Size = New System.Drawing.Size(519, 495)
        Me.dgvOKS.TabIndex = 0
        '
        'dgvForestry
        '
        Me.dgvForestry.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvForestry.Location = New System.Drawing.Point(5, 6)
        Me.dgvForestry.Name = "dgvForestry"
        Me.dgvForestry.Size = New System.Drawing.Size(519, 495)
        Me.dgvForestry.TabIndex = 0
        '
        'dgvSpecialZone
        '
        Me.dgvSpecialZone.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSpecialZone.Location = New System.Drawing.Point(10, 12)
        Me.dgvSpecialZone.Name = "dgvSpecialZone"
        Me.dgvSpecialZone.Size = New System.Drawing.Size(514, 489)
        Me.dgvSpecialZone.TabIndex = 0
        '
        'ForestRequestForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(559, 595)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.lblHint)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Name = "ForestRequestForm"
        Me.Text = "ForestRequestForm"
        Me.TabControl1.ResumeLayout(False)
        Me.tabForestPart.ResumeLayout(False)
        Me.tabOKS.ResumeLayout(False)
        Me.tabForestry.ResumeLayout(False)
        Me.tabSpecialZone.ResumeLayout(False)
        CType(Me.dgvForestPart, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvOKS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvForestry, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSpecialZone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents btnOk As Windows.Forms.Button
    Friend WithEvents lblHint As Windows.Forms.Label
    Friend WithEvents TabControl1 As Windows.Forms.TabControl
    Friend WithEvents tabForestPart As Windows.Forms.TabPage
    Friend WithEvents dgvForestPart As Windows.Forms.DataGridView
    Friend WithEvents tabOKS As Windows.Forms.TabPage
    Friend WithEvents dgvOKS As Windows.Forms.DataGridView
    Friend WithEvents tabForestry As Windows.Forms.TabPage
    Friend WithEvents dgvForestry As Windows.Forms.DataGridView
    Friend WithEvents tabSpecialZone As Windows.Forms.TabPage
    Friend WithEvents dgvSpecialZone As Windows.Forms.DataGridView
End Class
