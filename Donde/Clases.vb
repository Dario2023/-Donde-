Imports Oracle.DataAccess.Client ' Sentencia para importar la clase OraleDataAcces



Public Class Persona

    
    Private Nombre, Apellido As String
    Private Dni, Cuil, EstadoCivil As Integer
    Private FechaNacimiento As Date

    Private TxtNombrea, TxtApellidoa, TxtDnia, TxtTipoa, TxtNumeroa, TxtDigitoVerificadora As TextBox
    Private CmbSexoa, CmbEstadoCivila As ComboBox
    Private DtpFNacimientoa As DateTimePicker

    Public Sub AgregarPersona(ByRef Sexop As Byte, ByRef EstadoCivilp As Byte)

        Dim Adaptador As OracleDataAdapter
        Dim PersonaDS As New DataSet
        Dim Registro As DataRow

        Dim InsertCmd As New OracleCommand
        Dim UpdateCmd As New OracleCommand
        Dim DeleteCmd As New OracleCommand


        Registro("APELLIDO") = TxtApellidoa
        Registro("NOMBRE") = TxtNombrea
        Registro("DNI") = CInt(TxtDnia.Text)
        Registro("CUIL") = TxtTipoa.Text + TxtNumeroa.Text + TxtDigitoVerificadora.Text
        Registro("SEXO") = CType(CmbSexoa.SelectedIndex, Sexop)
        Registro("ESTADOCIVIL") = CType(CmbEstadoCivila.SelectedIndex, EstadoCivilp)
        Registro("FECHANACIMIENTO") = DtpFNacimientoa.Value

        If F_Donde.Accion = TipoAccion.Alta Then
            PersonaDS.Tables("persona").Rows.Add(Registro)
        ElseIf F_Donde.Accion = TipoAccion.Baja Then
            PersonaDS.Tables("persona").Rows.Remove(Registro)
        End If

        InsertCmd.CommandText = "Insert Into Persona"
           VALUES (:idpersona,:apellidoynombre,:dni,:cuil,:sexo,:estadocivil,:fechanacimiento)"
        UpdateCmd.CommandText = "Update Persona "
            set Apellido = :apellido,
                Nombre = :nombre,
                Dni = :dni,
                CUIL = :cuil,
                Sexo = :sexo,
                EstadoCivil = :estadocivil,
                FechaNacimiento = :fechanacimiento
            where Id_Persona = :idpersona"

        DeleteCmd.CommandText = "Delete * From Persona Where Id_Persona = :idpersona"

        InsertCmd.Connection = Conexion
        UpdateCmd.Connection = Conexion
        DeleteCmd.Connection = Conexion

        InsertCmd.Parameters.Add(New OracleParameter(":idpersona", OracleDbType.Int32, 0, "ID_PERSONA"))
        InsertCmd.Parameters.Add(New OracleParameter(":apellidoynombre", OracleDbType.Varchar2, 0, "APELLIDOYNOMBRE"))
        InsertCmd.Parameters.Add(New OracleParameter(":sexo", OracleDbType.Byte, 0, "SEXO"))
        InsertCmd.Parameters.Add(New OracleParameter(":dni", OracleDbType.Varchar2, 0, "DNI"))
        InsertCmd.Parameters.Add(New OracleParameter(":cuil", OracleDbType.Varchar2, 0, "CUIL"))
        InsertCmd.Parameters.Add(New OracleParameter(":fechanacimiento", OracleDbType.Date, 0, "FECHANACIMIENTO"))
        InsertCmd.Parameters.Add(New OracleParameter(":estadocivil", OracleDbType.Byte, 0, "ESTADOCIVIL"))


        UpdateCmd.Parameters.Add(New OracleParameter(":apellidoynombre", OracleDbType.Varchar2, 0, "APELLIDOYNOMBRE"))
        UpdateCmd.Parameters.Add(New OracleParameter(":dni", OracleDbType.Varchar2, 8, "DNI"))
        UpdateCmd.Parameters.Add(New OracleParameter(":cuil", OracleDbType.Varchar2, 13, "CUIL"))
        UpdateCmd.Parameters.Add(New OracleParameter(":sexo", OracleDbType.Byte, 0, "SEXO"))
        UpdateCmd.Parameters.Add(New OracleParameter(":estadocivil", OracleDbType.Byte, 0, "ESTADOCIVIL"))
        UpdateCmd.Parameters.Add(New OracleParameter(":fechanacimiento", OracleDbType.Date, 0, "FECHANACIMIENTO"))
        UpdateCmd.Parameters.Add(New OracleParameter(":idpersona", OracleDbType.Int32, 0, "ID_PERSONA"))

        DeleteCmd.Parameters.Add(New OracleParameter(":idpersona", OracleDbType.Int32, 0, "ID_PERSONA"))

        Adaptador.InsertCommand = InsertCmd
        Adaptador.UpdateCommand = UpdateCmd
        Adaptador.DeleteCommand = DeleteCmd
        Try
            Adaptador.Update(PersonaDS, "persona")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        If F_Donde.Accion = TipoAccion.Alta Then
            MessageBox.Show("Los datos se guardaron correctamente.")
        ElseIf F_Donde.Accion = TipoAccion.Modificacion Then
            MessageBox.Show("Los datos se actualizaron correctamente.")
        Else
            MessageBox.Show("El registro se eliminó correctamente.")
        End If
        'Form1.CargarComboPersonas()
        'Me.Close()



    End Sub

End Class

Public Class Sitios


End Class



n