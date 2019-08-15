<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.btnCreateOSV = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnOpenCurrentDirectory = New System.Windows.Forms.Button()
        Me.cbo_MonthOSV = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnSettings = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnOpenFiles = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnCreateOSV
        '
        Me.btnCreateOSV.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnCreateOSV.Location = New System.Drawing.Point(12, 64)
        Me.btnCreateOSV.Name = "btnCreateOSV"
        Me.btnCreateOSV.Size = New System.Drawing.Size(421, 48)
        Me.btnCreateOSV.TabIndex = 0
        Me.btnCreateOSV.Text = "Сформировать"
        Me.btnCreateOSV.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(59, 37)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(374, 21)
        Me.ComboBox1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Склад:"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(290, 197)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(12, 124)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(193, 17)
        Me.CheckBox1.TabIndex = 4
        Me.CheckBox1.Text = "Сформировать по всем складам"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(89, 197)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnOpenCurrentDirectory
        '
        Me.btnOpenCurrentDirectory.Image = CType(resources.GetObject("btnOpenCurrentDirectory.Image"), System.Drawing.Image)
        Me.btnOpenCurrentDirectory.Location = New System.Drawing.Point(12, 176)
        Me.btnOpenCurrentDirectory.Name = "btnOpenCurrentDirectory"
        Me.btnOpenCurrentDirectory.Size = New System.Drawing.Size(50, 48)
        Me.btnOpenCurrentDirectory.TabIndex = 6
        Me.btnOpenCurrentDirectory.UseVisualStyleBackColor = True
        '
        'cbo_MonthOSV
        '
        Me.cbo_MonthOSV.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbo_MonthOSV.FormattingEnabled = True
        Me.cbo_MonthOSV.Location = New System.Drawing.Point(104, 10)
        Me.cbo_MonthOSV.Name = "cbo_MonthOSV"
        Me.cbo_MonthOSV.Size = New System.Drawing.Size(162, 21)
        Me.cbo_MonthOSV.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(272, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(161, 23)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "sdfasdfasdf asdf asdfas"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnSettings
        '
        Me.btnSettings.Location = New System.Drawing.Point(130, 118)
        Me.btnSettings.Name = "btnSettings"
        Me.btnSettings.Size = New System.Drawing.Size(75, 23)
        Me.btnSettings.TabIndex = 9
        Me.btnSettings.Text = "Настройка"
        Me.btnSettings.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(12, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 23)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Выбор месяца:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnOpenFiles
        '
        Me.btnOpenFiles.Location = New System.Drawing.Point(320, 118)
        Me.btnOpenFiles.Name = "btnOpenFiles"
        Me.btnOpenFiles.Size = New System.Drawing.Size(109, 23)
        Me.btnOpenFiles.TabIndex = 11
        Me.btnOpenFiles.Text = "Открыть файлы"
        Me.btnOpenFiles.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(232, 118)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(82, 23)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "Файлы Excel"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(441, 146)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnOpenFiles)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnSettings)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbo_MonthOSV)
        Me.Controls.Add(Me.btnOpenCurrentDirectory)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.btnCreateOSV)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Программа для сохранения ОСВ в PDF"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCreateOSV As System.Windows.Forms.Button
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents btnOpenCurrentDirectory As System.Windows.Forms.Button
    Friend WithEvents cbo_MonthOSV As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnSettings As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnOpenFiles As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button

End Class
