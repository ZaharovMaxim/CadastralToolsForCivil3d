<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormCreateParcels
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.ComboBox8 = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.ComboBox5 = New System.Windows.Forms.ComboBox()
        Me.ContextMenuParcel = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОбоснованиеРазмераУчасткаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОбоснованиеИзъятияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПрочитатьХарактеристикиИзPSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПрочитатьХарактеристикиИзБлокаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПрочитатьХарактеристикиИзРосреестраToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.ComboBox12 = New System.Windows.Forms.ComboBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.NumericUpDown4 = New System.Windows.Forms.NumericUpDown()
        Me.NumericUpDown2 = New System.Windows.Forms.NumericUpDown()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.CheckBox7 = New System.Windows.Forms.CheckBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.ComboBox11 = New System.Windows.Forms.ComboBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ContextMenuBuild = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ТипСтроенияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ФункциональноеНазначениеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.МатериалСтенToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuParcel.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        CType(Me.NumericUpDown4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuBuild.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(259, 22)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(103, 21)
        Me.Label18.TabIndex = 6
        Me.Label18.Text = "Категория"
        '
        'ComboBox8
        '
        Me.ComboBox8.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox8.FormattingEnabled = True
        Me.ComboBox8.Location = New System.Drawing.Point(6, 19)
        Me.ComboBox8.Name = "ComboBox8"
        Me.ComboBox8.Size = New System.Drawing.Size(247, 21)
        Me.ComboBox8.TabIndex = 1
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(112, 42)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(98, 13)
        Me.Label13.TabIndex = 33
        Me.Label13.Text = "Масштаб чертежа"
        '
        'ComboBox5
        '
        Me.ComboBox5.FormattingEnabled = True
        Me.ComboBox5.Location = New System.Drawing.Point(7, 39)
        Me.ComboBox5.Name = "ComboBox5"
        Me.ComboBox5.Size = New System.Drawing.Size(93, 21)
        Me.ComboBox5.TabIndex = 35
        '
        'ContextMenuParcel
        '
        Me.ContextMenuParcel.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ContextMenuParcel.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.ToolStripMenuItem2, Me.ToolStripMenuItem3, Me.ToolStripMenuItem4, Me.ОбоснованиеРазмераУчасткаToolStripMenuItem, Me.ОбоснованиеИзъятияToolStripMenuItem, Me.ПрочитатьХарактеристикиИзPSToolStripMenuItem, Me.ПрочитатьХарактеристикиИзБлокаToolStripMenuItem, Me.ПрочитатьХарактеристикиИзРосреестраToolStripMenuItem})
        Me.ContextMenuParcel.Name = "ContextMenuStrip1"
        Me.ContextMenuParcel.Size = New System.Drawing.Size(362, 202)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(361, 22)
        Me.ToolStripMenuItem1.Text = "Категория земель"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(361, 22)
        Me.ToolStripMenuItem2.Text = "Разрешенное использование"
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        Me.ToolStripMenuItem3.Size = New System.Drawing.Size(361, 22)
        Me.ToolStripMenuItem3.Text = "Вид права"
        '
        'ToolStripMenuItem4
        '
        Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
        Me.ToolStripMenuItem4.Size = New System.Drawing.Size(361, 22)
        Me.ToolStripMenuItem4.Text = "Способ образования земельного участка"
        '
        'ОбоснованиеРазмераУчасткаToolStripMenuItem
        '
        Me.ОбоснованиеРазмераУчасткаToolStripMenuItem.Name = "ОбоснованиеРазмераУчасткаToolStripMenuItem"
        Me.ОбоснованиеРазмераУчасткаToolStripMenuItem.Size = New System.Drawing.Size(361, 22)
        Me.ОбоснованиеРазмераУчасткаToolStripMenuItem.Text = "Обоснование размера земельного участка"
        '
        'ОбоснованиеИзъятияToolStripMenuItem
        '
        Me.ОбоснованиеИзъятияToolStripMenuItem.Name = "ОбоснованиеИзъятияToolStripMenuItem"
        Me.ОбоснованиеИзъятияToolStripMenuItem.Size = New System.Drawing.Size(361, 22)
        Me.ОбоснованиеИзъятияToolStripMenuItem.Text = "Обоснование изъятия (занятия) земельного участка"
        '
        'ПрочитатьХарактеристикиИзPSToolStripMenuItem
        '
        Me.ПрочитатьХарактеристикиИзPSToolStripMenuItem.Name = "ПрочитатьХарактеристикиИзPSToolStripMenuItem"
        Me.ПрочитатьХарактеристикиИзPSToolStripMenuItem.Size = New System.Drawing.Size(361, 22)
        Me.ПрочитатьХарактеристикиИзPSToolStripMenuItem.Text = "Прочитать характеристики из PS"
        '
        'ПрочитатьХарактеристикиИзБлокаToolStripMenuItem
        '
        Me.ПрочитатьХарактеристикиИзБлокаToolStripMenuItem.Name = "ПрочитатьХарактеристикиИзБлокаToolStripMenuItem"
        Me.ПрочитатьХарактеристикиИзБлокаToolStripMenuItem.Size = New System.Drawing.Size(361, 22)
        Me.ПрочитатьХарактеристикиИзБлокаToolStripMenuItem.Text = "Прочитать характеристики из Блока"
        '
        'ПрочитатьХарактеристикиИзРосреестраToolStripMenuItem
        '
        Me.ПрочитатьХарактеристикиИзРосреестраToolStripMenuItem.Name = "ПрочитатьХарактеристикиИзРосреестраToolStripMenuItem"
        Me.ПрочитатьХарактеристикиИзРосреестраToolStripMenuItem.Size = New System.Drawing.Size(361, 22)
        Me.ПрочитатьХарактеристикиИзРосреестраToolStripMenuItem.Text = "Прочитать характеристики из Росреестра"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DataGridView2)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.ComboBox12)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Controls.Add(Me.ComboBox8)
        Me.GroupBox2.Controls.Add(Me.Label18)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 5)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(368, 489)
        Me.GroupBox2.TabIndex = 47
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Характеристики объекта"
        '
        'DataGridView2
        '
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column14, Me.Column15})
        Me.DataGridView2.ContextMenuStrip = Me.ContextMenuParcel
        Me.DataGridView2.Location = New System.Drawing.Point(5, 175)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(357, 307)
        Me.DataGridView2.TabIndex = 0
        '
        'Column14
        '
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column14.DefaultCellStyle = DataGridViewCellStyle2
        Me.Column14.HeaderText = "Имя"
        Me.Column14.Name = "Column14"
        Me.Column14.Width = 150
        '
        'Column15
        '
        Me.Column15.HeaderText = "Значение"
        Me.Column15.Name = "Column15"
        Me.Column15.Width = 350
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(256, 49)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(103, 27)
        Me.Label14.TabIndex = 68
        Me.Label14.Text = "Создаваемый Объект"
        '
        'ComboBox12
        '
        Me.ComboBox12.FormattingEnabled = True
        Me.ComboBox12.Location = New System.Drawing.Point(5, 49)
        Me.ComboBox12.Name = "ComboBox12"
        Me.ComboBox12.Size = New System.Drawing.Size(247, 21)
        Me.ComboBox12.TabIndex = 67
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ComboBox2)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.NumericUpDown4)
        Me.GroupBox3.Controls.Add(Me.NumericUpDown2)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.CheckBox7)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 75)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(353, 94)
        Me.GroupBox3.TabIndex = 66
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Условное обозначение"
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(6, 19)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(155, 21)
        Me.ComboBox2.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(167, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(186, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Обозначение (префикс) номера"
        '
        'NumericUpDown4
        '
        Me.NumericUpDown4.Location = New System.Drawing.Point(6, 46)
        Me.NumericUpDown4.Maximum = New Decimal(New Integer() {1215752192, 23, 0, 0})
        Me.NumericUpDown4.Name = "NumericUpDown4"
        Me.NumericUpDown4.Size = New System.Drawing.Size(48, 20)
        Me.NumericUpDown4.TabIndex = 21
        Me.NumericUpDown4.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'NumericUpDown2
        '
        Me.NumericUpDown2.Location = New System.Drawing.Point(124, 71)
        Me.NumericUpDown2.Maximum = New Decimal(New Integer() {1410065408, 2, 0, 0})
        Me.NumericUpDown2.Name = "NumericUpDown2"
        Me.NumericUpDown2.Size = New System.Drawing.Size(48, 20)
        Me.NumericUpDown2.TabIndex = 64
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(173, 75)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(144, 13)
        Me.Label8.TabIndex = 65
        Me.Label8.Text = "Следующий номер контура"
        '
        'CheckBox7
        '
        Me.CheckBox7.AutoSize = True
        Me.CheckBox7.Location = New System.Drawing.Point(6, 72)
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.Size = New System.Drawing.Size(112, 17)
        Me.CheckBox7.TabIndex = 61
        Me.CheckBox7.Text = "Многоконтурный"
        Me.CheckBox7.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(60, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(146, 13)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Следующий номер объекта"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.ComboBox11)
        Me.GroupBox1.Controls.Add(Me.ComboBox5)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 501)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(368, 90)
        Me.GroupBox1.TabIndex = 59
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Подпись"
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(302, 16)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(55, 17)
        Me.RadioButton3.TabIndex = 57
        Me.RadioButton3.Text = "Текст"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Checked = True
        Me.RadioButton2.Location = New System.Drawing.Point(128, 16)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(106, 17)
        Me.RadioButton2.TabIndex = 56
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Мультивыноска"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(7, 16)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(89, 17)
        Me.RadioButton1.TabIndex = 55
        Me.RadioButton1.Text = "Не рисовать"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(243, 66)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(120, 13)
        Me.Label12.TabIndex = 54
        Me.Label12.Text = "Стиль мультивыноски"
        '
        'ComboBox11
        '
        Me.ComboBox11.FormattingEnabled = True
        Me.ComboBox11.Location = New System.Drawing.Point(7, 64)
        Me.ComboBox11.Name = "ComboBox11"
        Me.ComboBox11.Size = New System.Drawing.Size(233, 21)
        Me.ComboBox11.TabIndex = 53
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(380, 5)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(600, 600)
        Me.PictureBox1.TabIndex = 44
        Me.PictureBox1.TabStop = False
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(259, 598)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(114, 25)
        Me.Button4.TabIndex = 46
        Me.Button4.Text = "Отмена"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(127, 598)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(117, 25)
        Me.Button3.TabIndex = 45
        Me.Button3.Text = "Выбрать контур"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(6, 598)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(106, 25)
        Me.Button1.TabIndex = 43
        Me.Button1.Text = "Применить"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ToolTip1
        '
        '
        'ContextMenuBuild
        '
        Me.ContextMenuBuild.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ContextMenuBuild.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ТипСтроенияToolStripMenuItem, Me.ФункциональноеНазначениеToolStripMenuItem, Me.МатериалСтенToolStripMenuItem})
        Me.ContextMenuBuild.Name = "ContextMenuBuild"
        Me.ContextMenuBuild.Size = New System.Drawing.Size(237, 70)
        '
        'ТипСтроенияToolStripMenuItem
        '
        Me.ТипСтроенияToolStripMenuItem.Name = "ТипСтроенияToolStripMenuItem"
        Me.ТипСтроенияToolStripMenuItem.Size = New System.Drawing.Size(236, 22)
        Me.ТипСтроенияToolStripMenuItem.Text = "Тип строения"
        '
        'ФункциональноеНазначениеToolStripMenuItem
        '
        Me.ФункциональноеНазначениеToolStripMenuItem.Name = "ФункциональноеНазначениеToolStripMenuItem"
        Me.ФункциональноеНазначениеToolStripMenuItem.Size = New System.Drawing.Size(236, 22)
        Me.ФункциональноеНазначениеToolStripMenuItem.Text = "Функциональное назначение"
        '
        'МатериалСтенToolStripMenuItem
        '
        Me.МатериалСтенToolStripMenuItem.Name = "МатериалСтенToolStripMenuItem"
        Me.МатериалСтенToolStripMenuItem.Size = New System.Drawing.Size(236, 22)
        Me.МатериалСтенToolStripMenuItem.Text = "Материал стен"
        '
        'FormCreateParcels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.SkyBlue
        Me.ClientSize = New System.Drawing.Size(991, 628)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button1)
        Me.Name = "FormCreateParcels"
        Me.Text = "Создать участок"
        Me.ContextMenuParcel.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.NumericUpDown4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuBuild.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label18 As Windows.Forms.Label
    Friend WithEvents ComboBox8 As Windows.Forms.ComboBox
    Friend WithEvents Label13 As Windows.Forms.Label
    Friend WithEvents ComboBox5 As Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents Button4 As Windows.Forms.Button
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents NumericUpDown4 As Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents CheckBox7 As Windows.Forms.CheckBox
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents NumericUpDown2 As Windows.Forms.NumericUpDown
    Friend WithEvents ContextMenuParcel As Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem4 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents Label12 As Windows.Forms.Label
    Friend WithEvents ComboBox11 As Windows.Forms.ComboBox
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents DataGridView2 As Windows.Forms.DataGridView
    Friend WithEvents Label14 As Windows.Forms.Label
    Friend WithEvents ComboBox12 As Windows.Forms.ComboBox
    Friend WithEvents ОбоснованиеРазмераУчасткаToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ОбоснованиеИзъятияToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ПрочитатьХарактеристикиИзPSToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ПрочитатьХарактеристикиИзБлокаToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ПрочитатьХарактеристикиИзРосреестраToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Column14 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column15 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ContextMenuBuild As Windows.Forms.ContextMenuStrip
    Friend WithEvents ТипСтроенияToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ФункциональноеНазначениеToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents МатериалСтенToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents RadioButton3 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As Windows.Forms.RadioButton
End Class
