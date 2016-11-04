<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F_Donde
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Btn_AP = New System.Windows.Forms.Button()
        Me.GB_Persona = New System.Windows.Forms.GroupBox()
        Me.Btn_BP = New System.Windows.Forms.Button()
        Me.Btn_MP = New System.Windows.Forms.Button()
        Me.Btn_EP = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Btn_BS = New System.Windows.Forms.Button()
        Me.Btn_MS = New System.Windows.Forms.Button()
        Me.Btn_ES = New System.Windows.Forms.Button()
        Me.Btn_AS = New System.Windows.Forms.Button()
        Me.DGV_Persona = New System.Windows.Forms.DataGridView()
        Me.DGV_Sitio = New System.Windows.Forms.DataGridView()
        Me.GB_Persona.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DGV_Persona, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGV_Sitio, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Btn_AP
        '
        Me.Btn_AP.Location = New System.Drawing.Point(17, 70)
        Me.Btn_AP.Name = "Btn_AP"
        Me.Btn_AP.Size = New System.Drawing.Size(77, 24)
        Me.Btn_AP.TabIndex = 0
        Me.Btn_AP.Text = "Agregar "
        Me.Btn_AP.UseVisualStyleBackColor = True
        '
        'GB_Persona
        '
        Me.GB_Persona.BackColor = System.Drawing.Color.LightSteelBlue
        Me.GB_Persona.Controls.Add(Me.Btn_BP)
        Me.GB_Persona.Controls.Add(Me.Btn_MP)
        Me.GB_Persona.Controls.Add(Me.Btn_EP)
        Me.GB_Persona.Controls.Add(Me.Btn_AP)
        Me.GB_Persona.Location = New System.Drawing.Point(12, 24)
        Me.GB_Persona.Name = "GB_Persona"
        Me.GB_Persona.Size = New System.Drawing.Size(110, 204)
        Me.GB_Persona.TabIndex = 1
        Me.GB_Persona.TabStop = False
        Me.GB_Persona.Text = "Persona"
        '
        'Btn_BP
        '
        Me.Btn_BP.Location = New System.Drawing.Point(17, 27)
        Me.Btn_BP.Name = "Btn_BP"
        Me.Btn_BP.Size = New System.Drawing.Size(77, 24)
        Me.Btn_BP.TabIndex = 3
        Me.Btn_BP.Text = "Buscar"
        Me.Btn_BP.UseVisualStyleBackColor = True
        '
        'Btn_MP
        '
        Me.Btn_MP.Location = New System.Drawing.Point(17, 116)
        Me.Btn_MP.Name = "Btn_MP"
        Me.Btn_MP.Size = New System.Drawing.Size(77, 24)
        Me.Btn_MP.TabIndex = 2
        Me.Btn_MP.Text = "Modificar"
        Me.Btn_MP.UseVisualStyleBackColor = True
        '
        'Btn_EP
        '
        Me.Btn_EP.Location = New System.Drawing.Point(17, 161)
        Me.Btn_EP.Name = "Btn_EP"
        Me.Btn_EP.Size = New System.Drawing.Size(77, 24)
        Me.Btn_EP.TabIndex = 1
        Me.Btn_EP.Text = "Eliminar "
        Me.Btn_EP.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.GroupBox1.Controls.Add(Me.Btn_BS)
        Me.GroupBox1.Controls.Add(Me.Btn_MS)
        Me.GroupBox1.Controls.Add(Me.Btn_ES)
        Me.GroupBox1.Controls.Add(Me.Btn_AS)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 248)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(110, 204)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Sitio"
        '
        'Btn_BS
        '
        Me.Btn_BS.Location = New System.Drawing.Point(17, 27)
        Me.Btn_BS.Name = "Btn_BS"
        Me.Btn_BS.Size = New System.Drawing.Size(77, 24)
        Me.Btn_BS.TabIndex = 3
        Me.Btn_BS.Text = "Buscar"
        Me.Btn_BS.UseVisualStyleBackColor = True
        '
        'Btn_MS
        '
        Me.Btn_MS.Location = New System.Drawing.Point(17, 116)
        Me.Btn_MS.Name = "Btn_MS"
        Me.Btn_MS.Size = New System.Drawing.Size(77, 24)
        Me.Btn_MS.TabIndex = 2
        Me.Btn_MS.Text = "Modificar"
        Me.Btn_MS.UseVisualStyleBackColor = True
        '
        'Btn_ES
        '
        Me.Btn_ES.Location = New System.Drawing.Point(17, 161)
        Me.Btn_ES.Name = "Btn_ES"
        Me.Btn_ES.Size = New System.Drawing.Size(77, 24)
        Me.Btn_ES.TabIndex = 1
        Me.Btn_ES.Text = "Eliminar "
        Me.Btn_ES.UseVisualStyleBackColor = True
        '
        'Btn_AS
        '
        Me.Btn_AS.Location = New System.Drawing.Point(17, 70)
        Me.Btn_AS.Name = "Btn_AS"
        Me.Btn_AS.Size = New System.Drawing.Size(77, 24)
        Me.Btn_AS.TabIndex = 0
        Me.Btn_AS.Text = "Agregar "
        Me.Btn_AS.UseVisualStyleBackColor = True
        '
        'DGV_Persona
        '
        Me.DGV_Persona.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_Persona.Location = New System.Drawing.Point(128, 24)
        Me.DGV_Persona.Name = "DGV_Persona"
        Me.DGV_Persona.Size = New System.Drawing.Size(439, 204)
        Me.DGV_Persona.TabIndex = 3
        '
        'DGV_Sitio
        '
        Me.DGV_Sitio.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_Sitio.Location = New System.Drawing.Point(128, 248)
        Me.DGV_Sitio.Name = "DGV_Sitio"
        Me.DGV_Sitio.Size = New System.Drawing.Size(439, 204)
        Me.DGV_Sitio.TabIndex = 4
        '
        'F_Donde
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(579, 485)
        Me.Controls.Add(Me.DGV_Sitio)
        Me.Controls.Add(Me.DGV_Persona)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GB_Persona)
        Me.Name = "F_Donde"
        Me.Text = "¿Donde?"
        Me.GB_Persona.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DGV_Persona, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGV_Sitio, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Btn_AP As System.Windows.Forms.Button
    Friend WithEvents GB_Persona As System.Windows.Forms.GroupBox
    Friend WithEvents Btn_BP As System.Windows.Forms.Button
    Friend WithEvents Btn_MP As System.Windows.Forms.Button
    Friend WithEvents Btn_EP As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Btn_BS As System.Windows.Forms.Button
    Friend WithEvents Btn_MS As System.Windows.Forms.Button
    Friend WithEvents Btn_ES As System.Windows.Forms.Button
    Friend WithEvents Btn_AS As System.Windows.Forms.Button
    Friend WithEvents DGV_Persona As System.Windows.Forms.DataGridView
    Friend WithEvents DGV_Sitio As System.Windows.Forms.DataGridView

End Class
