Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListDarDiferencialDeAliquota
    Inherits List(Of DarDiferencialDeAliquota)

#Region "Construtor"
    Public Sub New(Optional ByVal CarregarDados As Boolean = False, Optional ByVal Parametros As Dictionary(Of String, Object) = Nothing)
        If CarregarDados Then

            Dim objBanco As New AcessaBanco()
            Dim objDar As New DarDiferencialDeAliquota

            Dim strSQL As String
            strSQL = "SELECT Dar_Id, Empresa_Id, EndEmpresa_Id, Data, DataReferencia, Cliente, EndCliente, Nota, Serie, CodigoReceita, Valor, SpedInfAdicionaisDeApuracao " & _
                     "  FROM DarDiferencialDeAliquota " & _
                     " WHERE 1=1"

            If Not Parametros Is Nothing Then
                If Parametros.ContainsKey("Ano") AndAlso Not String.IsNullOrEmpty(Parametros("Ano")) Then
                    strSQL &= "   AND YEAR(DataReferencia) = " & CInt(Parametros("Ano"))
                End If

                If Parametros.ContainsKey("Mes") AndAlso Not String.IsNullOrEmpty(Parametros("Mes")) Then
                    strSQL &= "   AND MONTH(DataReferencia) = " & CInt(Parametros("Mes"))
                End If

                If Parametros.ContainsKey("Empresa") AndAlso Not String.IsNullOrEmpty(Parametros("Empresa")) Then
                    Dim s As String = Parametros("Empresa")
                    strSQL &= "   AND Empresa_Id = '" & s.Split("-")(0) & "'" & _
                              "   AND EndEmpresa_Id = " & s.Split("-")(1)
                End If

            End If
            strSQL &= " ORDER BY Empresa_Id, EndEmpresa_Id, Dar_Id, Data"

            Dim dsTxC As DataSet = objBanco.ConsultaDataSet(strSQL, "DarDiferencialDeAliquota")

            For Each drTxC As DataRow In dsTxC.Tables(0).Rows
                objDar = New DarDiferencialDeAliquota()
                objDar.Dar = drTxC("Dar_Id")
                objDar.CodigoEmpresa = drTxC("Empresa_Id")
                objDar.EndEmpresa = drTxC("EndEmpresa_Id")
                objDar.Data = drTxC("Data")
                objDar.DataReferencia = drTxC("DataReferencia")
                objDar.CodigoCliente = drTxC("Cliente")
                objDar.EndCliente = drTxC("EndCliente")
                objDar.Nota = drTxC("Nota")
                objDar.Serie = drTxC("Serie")
                objDar.CodigoReceita = drTxC("CodigoReceita")
                objDar.Valor = drTxC("Valor")
                objDar.CodigoSpedInfAdicionaisDeApuracao = drTxC("SpedInfAdicionaisDeApuracao")
                Me.Add(objDar)
            Next
        End If
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        Return MyBase.Equals(obj)
    End Function

#End Region

    Public Function ListaDeDarDiferencialDeAliquota(CaminhoDaImagem As String) As Object
        Return (From p In Me Order By p.Data, p.Dar Select p.Dar, Empresa = p.Empresa.Reduzido & " - " & p.Empresa.Nome & " / " & p.Empresa.Cidade & " - " & p.Empresa.Estado.Codigo, p.Data, p.DataReferencia, p.Cliente, p.Cliente.Nome, p.Cliente.Endereco, p.Nota, p.Serie, p.CodigoReceita, p.Valor, p.CodigoSpedInfAdicionaisDeApuracao, Imagem = [Lib].Negocio.Conversoes.ConverterImagemEmByte(CaminhoDaImagem))
    End Function

End Class

<Serializable()> _
Public Class DarDiferencialDeAliquota
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pDar As Integer)
        Dim objBanco As New AcessaBanco()
        Dim strSQL As String = "SELECT Dar, Data, DataReferencia, Cliente, EndCliente, Nota, Serie, CodigoReceita, Valor  " & _
                                   "FROM DarDiferencialDeAliquota " & _
                                   "WHERE Dar = " & pDar.ToString()
        Dim dsTxC As DataSet = objBanco.ConsultaDataSet(strSQL, "DarDiferencialDeAliquota")

        For Each drTxC As DataRow In dsTxC.Tables(0).Rows
            Dar = drTxC("Dar")
            Data = drTxC("Data")
            DataReferencia = drTxC("DataReferencia")
            CodigoCliente = drTxC("Cliente")
            EndCliente = drTxC("EndCliente")
            Nota = drTxC("Nota")
            Serie = drTxC("Serie")
            CodigoReceita = drTxC("CodigoReceita")
            Valor = drTxC("Valor")
        Next
    End Sub
#End Region

#Region "Field"
    Private _Cliente As [Lib].Negocio.Cliente
    Private _Empresa As [Lib].Negocio.Cliente
    Private _DataReferencia As DateTime
#End Region

#Region "Property"
    Public Property IUD As String = ""
    Public Property Dar As String
    Public Property DataReferencia() As DateTime
        Get
            Return _DataReferencia
        End Get
        Set(ByVal value As DateTime)
            _DataReferencia = value
        End Set
    End Property
    Public Property CodigoEmpresa As String = ""
    Public Property EndEmpresa As Integer

    Public ReadOnly Property Empresa() As [Lib].Negocio.Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
    End Property

    Public Property Data As Date
    Public Property CodigoCliente As String
    Public Property EndCliente As Integer
    Public ReadOnly Property Cliente() As [Lib].Negocio.Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Length > 0 Then _Cliente = New Cliente(_CodigoCliente, _EndCliente)
            Return _Cliente
        End Get
    End Property
    Public Property Nota As Integer
    Public Property Serie As String = ""
    Public Property CodigoReceita As Integer
    Public Property Valor As Decimal
    Public Property CodigoSpedInfAdicionaisDeApuracao As String = ""

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = "Insert Into DarDiferencialDeAliquota(Dar_Id, Empresa_Id, EndEmpresa_Id, Data, DataReferencia, Cliente, EndCliente, Nota, Serie, CodigoReceita, Valor, SpedInfAdicionaisDeApuracao)" & vbCrLf & _
                      " values('" & Dar & "', '" & CodigoEmpresa & "', " & EndEmpresa & ", '" & Data.ToString("yyyy/MM/dd") & "', '" & DataReferencia.ToString("yyyy/MM/dd") & "', '" & CodigoCliente & "', " & EndCliente & "," & Nota & ", '" & Serie & "'," & CodigoReceita & ", " & Str(Valor) & ",'" & CodigoSpedInfAdicionaisDeApuracao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update DarDiferencialDeAliquota Set  " & _
                      "  Data ='" & Data.ToString("yyyy/MM/dd") & "'" & _
                      "  DataReferencia = '" & DataReferencia.ToString("yyyy/MM/dd") & vbCrLf & _
                      " ,Cliente     ='" & CodigoCliente & "'" & _
                      " ,EndCliente     =" & EndCliente & _
                      " ,Nota     =" & Nota & _
                      " ,Serie     ='" & Serie & "'" & _
                      " ,CodigoReceita     =" & CodigoReceita & _
                      " ,Valor     =" & Str(Valor) & _
                      " ,SpedInfAdicionaisDeApuracao     ='" & CodigoSpedInfAdicionaisDeApuracao & "'" & _
                      " Where Dar_Id = " & Dar & _
                      " AND Empresa_id = '" & CodigoEmpresa & "'" & _
                      " AND EndEmpresa_Id = " & EndEmpresa
                Sqls.Add(Sql)
            Case "D"
                Sql = "Delete DarDiferencialDeAliquota " & vbCrLf & _
                      " Where Dar_Id = " & Dar & _
                      " AND Empresa_id = '" & CodigoEmpresa & "'" & _
                      " AND EndEmpresa_Id = " & EndEmpresa
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class