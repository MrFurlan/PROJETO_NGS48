Public Class ListPisCofins_Reg_1100_1500
    Inherits List(Of PisCofins_Reg_1100_1500)

#Region "Construtor"
    ''' <summary>
    ''' Cria uma instãncia vazia do objeto.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' Busca um registro específico, conforme parâmetros.
    ''' </summary>
    ''' <param name="registro">Se Informado 1100 busca registro = PIS - Registro 1100, senão busca registro COFINS - Registro 1500</param>
    ''' <param name="Mes">Campo Numérico, informando o mês</param>
    ''' <param name="Ano">Campo Numérico, informando o Ano</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal registro As Integer, ByVal Mes As Integer, ByVal Ano As Integer)
        Dim sql As String = "   SELECT Empresa_Id, Per_Apu_Cred, Orig_Cred, Cnpj_Suc, Cod_Cred, Vl_Cred_Apu, Vl_Cred_Ext_Apu, Vl_Tot_Cred_Apu, " & vbCrLf & _
                            "          Vl_Cred_Desc_Pa_Ant, Vl_Cred_Per_Pa_Ant, Vl_Cred_Dcomp_Pa_Ant, Sd_Cred_Disp_Efd," & vbCrLf & _
                            "          Vl_Cred_Desc_Efd, Vl_Cred_Per_Efd, Vl_Cred_Dcomp_Efd, Vl_Cred_Trans, Vl_Cred_Out, Sld_Cred_Fim, Mes, Ano" & vbCrLf & _
                            "     FROM " & IIf(registro = 1100, "PisCofins_Reg_1100", "PisCofins_Reg_1500") & vbCrLf & _
                            "    WHERE Ano = " & Ano & " AND Mes = " & Mes & vbCrLf

        Dim db As New AcessaBanco()
        Dim ds As DataSet = db.ConsultaDataSet(sql, "PisCofinsReg100")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New PisCofins_Reg_1100_1500()
                obj.CodEmpresa = row("Empresa_id").ToString()
                obj.Per_Apu_Cred = row("Per_Apu_Cred")
                obj.Orig_Cred = row("Orig_Cred")
                obj.Cnpj_Suc = row("Cnpj_Suc")
                obj.Cod_Cred = row("Cod_Cred")
                obj.Vl_Cred_Apu = row("Vl_Cred_Apu")
                obj.Vl_Cred_Ext_Apu = row("Vl_Cred_Ext_Apu")
                obj.Vl_Tot_Cred_Apu = row("Vl_Tot_Cred_Apu")
                obj.Vl_Cred_Desc_Pa_Ant = row("Vl_Cred_Desc_Pa_Ant")
                obj.Vl_Cred_Per_Pa_Ant = row("Vl_Cred_Per_Pa_Ant")
                obj.Vl_Cred_Dcomp_Pa_Ant = row("Vl_Cred_Dcomp_Pa_Ant")
                obj.Sd_Cred_Disp_Efd = row("Sd_Cred_Disp_Efd")
                obj.Vl_Cred_Desc_Efd = row("Vl_Cred_Desc_Efd")
                obj.Vl_Cred_Per_Efd = row("Vl_Cred_Per_Efd")
                obj.Vl_Cred_Dcomp_Efd = row("Vl_Cred_Dcomp_Efd")
                obj.Vl_Cred_Trans = row("Vl_Cred_Trans")
                obj.Vl_Cred_Out = row("Vl_Cred_Out")
                obj.Sld_Cred_Fim = row("Sld_Cred_Fim")
                obj.Mes = row("Mes")
                obj.Ano = row("Ano")
                Me.Add(obj)
            Next
        End If
    End Sub

#End Region

End Class
Public Class PisCofins_Reg_1100_1500

#Region "Fields"
    Private _IUD As String
    Private _Encargo As String
    Private _CodEmpresa As String
    Private _EndEmpresa As Integer
    Private _Empresa As Cliente
    Private _Per_Apu_Cred As String
    Private _Orig_Cred As String
    Private _Cnpj_Suc As String
    Private _Cod_Cred As String
    Private _Vl_Cred_Apu As Decimal
    Private _Vl_Cred_Ext_Apu As Decimal
    Private _Vl_Tot_Cred_Apu As Decimal
    Private _Vl_Cred_Desc_Pa_Ant As Decimal
    Private _Vl_Cred_Per_Pa_Ant As Decimal
    Private _Vl_Cred_Dcomp_Pa_Ant As Decimal
    Private _Sd_Cred_Disp_Efd As Decimal
    Private _Vl_Cred_Desc_Efd As Decimal
    Private _Vl_Cred_Per_Efd As Decimal
    Private _Vl_Cred_Dcomp_Efd As Decimal
    Private _Vl_Cred_Trans As Decimal
    Private _Vl_Cred_Out As Decimal
    Private _Sld_Cred_Fim As Decimal
    Private _Mes As Integer
    Private _Ano As Integer
#End Region

#Region "Properties"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodEmpresa() As String
        Get
            Return _CodEmpresa
        End Get
        Set(ByVal value As String)
            _CodEmpresa = value
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing AndAlso Not String.IsNullOrWhiteSpace(_CodEmpresa) Then _Empresa = New Cliente(_CodEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property Per_Apu_Cred() As String
        Get
            Return _Per_Apu_Cred
        End Get
        Set(ByVal value As String)
            _Per_Apu_Cred = value
        End Set
    End Property

    Public Property Orig_Cred() As String
        Get
            Return _Orig_Cred
        End Get
        Set(ByVal value As String)
            _Orig_Cred = value
        End Set
    End Property

    Public Property Cnpj_Suc() As String
        Get
            Return _Cnpj_Suc
        End Get
        Set(ByVal value As String)
            _Cnpj_Suc = value
        End Set
    End Property

    Public Property Cod_Cred() As String
        Get
            Return _Cod_Cred
        End Get
        Set(ByVal value As String)
            _Cod_Cred = value
        End Set
    End Property

    Public Property Vl_Cred_Apu() As Decimal
        Get
            Return _Vl_Cred_Apu
        End Get
        Set(ByVal value As Decimal)
            _Vl_Cred_Apu = value
        End Set
    End Property

    Public Property Vl_Cred_Ext_Apu() As Decimal
        Get
            Return _Vl_Cred_Ext_Apu
        End Get
        Set(ByVal value As Decimal)
            _Vl_Cred_Ext_Apu = value
        End Set
    End Property

    Public Property Vl_Tot_Cred_Apu() As Decimal
        Get
            Return _Vl_Tot_Cred_Apu
        End Get
        Set(ByVal value As Decimal)
            _Vl_Tot_Cred_Apu = value
        End Set
    End Property

    Public Property Vl_Cred_Desc_Pa_Ant() As Decimal
        Get
            Return _Vl_Cred_Desc_Pa_Ant
        End Get
        Set(ByVal value As Decimal)
            _Vl_Cred_Desc_Pa_Ant = value
        End Set
    End Property

    Public Property Vl_Cred_Per_Pa_Ant() As Decimal
        Get
            Return _Vl_Cred_Per_Pa_Ant
        End Get
        Set(ByVal value As Decimal)
            _Vl_Cred_Per_Pa_Ant = value
        End Set
    End Property

    Public Property Vl_Cred_Dcomp_Pa_Ant() As Decimal
        Get
            Return _Vl_Cred_Dcomp_Pa_Ant
        End Get
        Set(ByVal value As Decimal)
            _Vl_Cred_Dcomp_Pa_Ant = value
        End Set
    End Property

    Public Property Sd_Cred_Disp_Efd() As Decimal
        Get
            Return _Sd_Cred_Disp_Efd
        End Get
        Set(ByVal value As Decimal)
            _Sd_Cred_Disp_Efd = value
        End Set
    End Property

    Public Property Vl_Cred_Desc_Efd() As Decimal
        Get
            Return _Vl_Cred_Desc_Efd
        End Get
        Set(ByVal value As Decimal)
            _Vl_Cred_Desc_Efd = value
        End Set
    End Property

    Public Property Vl_Cred_Per_Efd() As Decimal
        Get
            Return _Vl_Cred_Per_Efd
        End Get
        Set(ByVal value As Decimal)
            _Vl_Cred_Per_Efd = value
        End Set
    End Property

    Public Property Vl_Cred_Dcomp_Efd() As Decimal
        Get
            Return _Vl_Cred_Dcomp_Efd
        End Get
        Set(ByVal value As Decimal)
            _Vl_Cred_Dcomp_Efd = value
        End Set
    End Property

    Public Property Vl_Cred_Trans() As Decimal
        Get
            Return _Vl_Cred_Trans
        End Get
        Set(ByVal value As Decimal)
            _Vl_Cred_Trans = value
        End Set
    End Property

    Public Property Vl_Cred_Out() As Decimal
        Get
            Return _Vl_Cred_Out
        End Get
        Set(ByVal value As Decimal)
            _Vl_Cred_Out = value
        End Set
    End Property

    Public Property Sld_Cred_Fim() As Decimal
        Get
            Return _Sld_Cred_Fim
        End Get
        Set(ByVal value As Decimal)
            _Sld_Cred_Fim = value
        End Set
    End Property

    Public Property Mes() As Integer
        Get
            Return _Mes
        End Get
        Set(ByVal value As Integer)
            _Mes = value
        End Set
    End Property

    Public Property Ano() As Integer
        Get
            Return _Ano
        End Get
        Set(ByVal value As Integer)
            _Ano = value
        End Set
    End Property

    Public Property Encargo() As String
        Get
            Return _Encargo
        End Get
        Set(ByVal value As String)
            _Encargo = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList
        SalvarSql(sqls)
        Return Banco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByRef sqls As ArrayList)
        Dim sql As String = String.Empty
        Select Case Me.IUD
            Case "I"
                sql = "INSERT Into PisCofins_Reg_1500 " & vbCrLf & _
                      "       (Empresa_Id, Per_Apu_Cred, Orig_Cred, Cnpj_Suc, Cod_Cred, Vl_Cred_Apu, Vl_Cred_Ext_Apu, Vl_Tot_Cred_Apu, Vl_Cred_Desc_Pa_Ant, Vl_Cred_Per_Pa_Ant, Vl_Cred_Dcomp_Pa_Ant," & vbCrLf & _
                      "        Sd_Cred_Disp_Efd, Vl_Cred_Desc_Efd, Vl_Cred_Per_Efd, Vl_Cred_Dcomp_Efd, Vl_Cred_Trans, Vl_Cred_Out, Sld_Cred_Fim, Mes, Ano)" & vbCrLf & _
                      "VALUES ('" & Me.CodEmpresa & "', '" & Me.Per_Apu_Cred & "', '" & Me.Orig_Cred & "', '" & Me.Cnpj_Suc & "', '" & Me.Cod_Cred & "', " & Str(Me.Vl_Cred_Apu) & "," & _
                      Str(Me.Vl_Cred_Ext_Apu) & ", " & Str(Me.Vl_Tot_Cred_Apu) & ", " & Str(Me.Vl_Cred_Desc_Pa_Ant) & ", " & Str(Me.Vl_Cred_Per_Pa_Ant) & ", " & Str(Me.Vl_Cred_Dcomp_Pa_Ant) & "," & vbCrLf & _
                      Str(Me.Sd_Cred_Disp_Efd) & ", " & Str(Me.Vl_Cred_Desc_Efd) & ", " & Me.mes & ", " & Me.ano & ")"
            Case "D"
                sql = ""
        End Select
        If Not String.IsNullOrWhiteSpace(sql) Then sqls.Add(sql)
    End Sub
#End Region

End Class
