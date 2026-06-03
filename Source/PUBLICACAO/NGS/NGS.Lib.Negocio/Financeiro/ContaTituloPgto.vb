Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ContaTituloPgto
    Implements IBaseEntity

    Private _Agencia As String
    Private _DigitoAgencia As String
    Private _Conta_Id As String
    Private _DigitoConta_Id As String
    Private _SequenciaDePagamento As Integer
    Private _SequenciaDeRegistroDePagamento As Integer
    Private _TipoConta As String = "C"

    Public Sub New(ByVal Emp As String, ByVal EndEmp As String, ByVal a As String, ByVal da As String, ByVal c As String, ByVal dc As String, ByVal tc As String)
        _Agencia = a
        _DigitoAgencia = da
        _Conta_Id = c
        _DigitoConta_Id = dc
        _TipoConta = tc
        If _SequenciaDePagamento.ToString = 0 Then
            'Numerador dos RemessaBancaria
            Dim intRegistro As Integer = Numerador.PegarNumero(Emp, EndEmp, eTiposNumerador.RemessaBancaria)
            _SequenciaDePagamento = intRegistro

        End If
        If _SequenciaDeRegistroDePagamento.ToString = 0 Then
            'Numerador dos RemessaBancaria
            Dim intRegistro As Integer = Numerador.PegarNumero(Emp, EndEmp, eTiposNumerador.RemessaBancaria)
            _SequenciaDeRegistroDePagamento = intRegistro

        End If
    End Sub

    Public Sub New(ByVal Emp As String, ByVal EndEmp As String, ByVal a As String, ByVal da As String, ByVal c As String, ByVal dc As String, ByVal tc As String, ByVal s As Integer, ByVal sr As Integer)
        _Agencia = a
        _DigitoAgencia = da
        _Conta_Id = c
        _DigitoConta_Id = dc
        _TipoConta = tc
        _SequenciaDePagamento = s
        _SequenciaDeRegistroDePagamento = sr
    End Sub

    Public Property Agencia() As String
        Get
            Return _Agencia
        End Get
        Set(ByVal value As String)
            _Agencia = value
        End Set
    End Property

    Public Property DigitoAgencia() As String
        Get
            Return _DigitoAgencia
        End Get
        Set(ByVal value As String)
            _DigitoAgencia = value
        End Set
    End Property

    Public Property Conta_Id() As String
        Get
            Return _Conta_Id
        End Get
        Set(ByVal value As String)
            _Conta_Id = value
        End Set
    End Property

    Public Property TipoConta() As String
        Get
            Return _TipoConta
        End Get
        Set(ByVal value As String)
            _TipoConta = value
        End Set
    End Property

    Public Property DigitoConta_Id() As String
        Get
            Return _DigitoConta_Id
        End Get
        Set(ByVal value As String)
            _DigitoConta_Id = value
        End Set
    End Property

    Public Property SequenciaDePagamento() As Integer
        Get
            Return _SequenciaDePagamento
        End Get
        Set(ByVal value As Integer)
            _SequenciaDePagamento = value
        End Set
    End Property

    Public Property SequenciaDeRegistroDePagamento() As Integer
        Get
            Return _SequenciaDeRegistroDePagamento
        End Get
        Set(ByVal value As Integer)
            _SequenciaDeRegistroDePagamento = value
        End Set
    End Property

    Public ReadOnly Property Descricao()
        Get
            If Agencia.Length = 0 Then
                Return "Selecione uma Agencia/Conta"
            Else
                Return "Agencia: " & Agencia & "-" & DigitoAgencia & "  Conta: " & Conta_Id & "-" & DigitoConta_Id
            End If

        End Get
    End Property

End Class
