Imports Oracle.DataAccess.Client

Public Class F_CPersona

    Public Enum EstadoCivil As Byte
        Soltero = 0
        Casado = 1
    End Enum

    Public Enum Sexo As Byte
        Femenino = 0
        Masculino = 1
    End Enum

    Private Conexion As New OracleConnection("Data Source=localhost;User Id=Dorian;Password=123456789;")

    Dim Adaptador As OracleDataAdapter
    Dim PersonaDS As New DataSet
    Dim Registro As DataRow

    Private Sub F_CPersona_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Cargo Combo Sexo
        CmbSexo.Items.Insert(Sexo.Femenino, Sexo.Femenino.ToString)
        CmbSexo.Items.Insert(Sexo.Masculino, Sexo.Masculino.ToString)

        'Cargo Combo Estado Civil
        CmbEstadoCivil.Items.Insert(EstadoCivil.Soltero, EstadoCivil.Soltero.ToString)
        CmbEstadoCivil.Items.Insert(EstadoCivil.Casado, EstadoCivil.Casado.ToString)

        Adaptador = New OracleDataAdapter("Select * From Persona Where Id_Persona = " + F_Donde.IdPersona.ToString, Conexion)
        Adaptador.Fill(PersonaDS, "persona")
        If F_Donde.Accion = TipoAccion.Alta Then
            Registro = PersonaDS.Tables("persona").NewRow()
            Registro("ID_PERSONA") = -1
        Else
            Registro = PersonaDS.Tables("persona").Rows.Item(0)
        End If

        TxtApellido.Text = Registro("APELLIDO").ToString
        TxtNombre.Text = Registro("NOMBRE").ToString
        TxtDni.Text = Registro("DNI").ToString
        TxtTipo.Text = Mid$(Registro("CUIL").ToString(), 1, 2)
        TxtNumero.Text = Mid$(Registro("CUIL").ToString(), 3, 8)
        TxtDigitoVerificador.Text = Mid$(Registro("CUIL").ToString(), 11, 1)
        CmbSexo.SelectedIndex = IIf(IsDBNull(Registro("SEXO")), -1, Registro("SEXO"))
        CmbEstadoCivil.SelectedIndex = IIf(IsDBNull(Registro("ESTADOCIVIL")), -1, Registro("ESTADOCIVIL"))
        DtpFNacimiento.Value = IIf(IsDBNull(Registro("FECHANACIMIENTO")), Date.Now, Registro("FECHANACIMIENTO"))

        TxtApellido.Enabled = (F_Donde.Accion <> TipoAccion.Baja)
        TxtApellido.Enabled = (F_Donde.Accion <> TipoAccion.Baja)
        TxtDni.Enabled = (F_Donde.Accion <> TipoAccion.Baja)
        TxtTipo.Enabled = (F_Donde.Accion <> TipoAccion.Baja)
        TxtNumero.Enabled = (F_Donde.Accion <> TipoAccion.Baja)
        TxtDigitoVerificador.Enabled = (F_Donde.Accion <> TipoAccion.Baja)
        DtpFNacimiento.Enabled = (F_Donde.Accion <> TipoAccion.Baja)
        CmbSexo.Enabled = (F_Donde.Accion <> TipoAccion.Baja)
        CmbEstadoCivil.Enabled = (F_Donde.Accion <> TipoAccion.Baja)

    End Sub






End Class
