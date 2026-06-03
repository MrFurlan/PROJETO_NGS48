Imports System.Xml.Serialization

<Serializable()> _
Public Class BreadcrumbsNode

    Private _root As BreadcrumbsNode
    <XmlIgnore()> _
    Public Property Root() As BreadcrumbsNode
        Get
            Return _root
        End Get
        Set(ByVal value As BreadcrumbsNode)
            _root = value
        End Set
    End Property

    Private _index As Integer
    Public Property Index() As Integer
        Get
            Return _index
        End Get
        Set(ByVal value As Integer)
            _index = value
        End Set
    End Property

    Private _title As String
    <XmlAttribute("title")> _
    Public Property Title() As String
        Get
            Return _title
        End Get
        Set(ByVal value As String)
            _title = value
        End Set
    End Property

    Private _description As String
    <XmlAttribute("description")> _
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Private _url As String
    <XmlAttribute("url")> _
    Public Property Url() As String
        Get
            Return _url
        End Get
        Set(ByVal value As String)
            _url = value
        End Set
    End Property

    Private _cssclass As String
    <XmlAttribute("class")> _
    Public Property CssClass() As String
        Get
            Return _cssclass
        End Get
        Set(ByVal value As String)
            _cssclass = value
        End Set
    End Property

    Private _nodes As List(Of BreadcrumbsNode)
    <XmlElement("BreadcrumbsNode")> _
    Public Property Nodes() As List(Of BreadcrumbsNode)
        Get
            Return _nodes
        End Get
        Set(ByVal value As List(Of BreadcrumbsNode))
            _nodes = value
        End Set
    End Property

End Class