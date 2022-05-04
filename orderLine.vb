Public Class orderLine

#Region "VARIABLES"

    Dim _iId As Integer
    Dim _sIdClient As String
    Dim _sClient As String
    Dim _iIdArticles As Integer
    Dim _sMonoSKU As String
    Dim _sOrderAX As String
    Dim _OrderAXDate As DateTime
    Dim _sOrderLineAx As String
    Dim _OrderLineAXDate As DateTime
    Dim _iQuantity As Integer
    Dim _bClientBlock As Integer
    Dim _sNotes As String
    Dim _RequestedDate As DateTime
    Dim _DesiredDate As DateTime
    Dim _Iid_states As Integer
    Dim _bIsDeleted As Integer


#End Region

#Region "PROPERTIES"

    Public Property sIdClient As String
        Get
            Return _sIdClient
        End Get
        Set(value As String)
            _sIdClient = value
        End Set
    End Property

    Public Property SClient As String
        Get
            Return _sClient
        End Get
        Set(value As String)
            _sClient = value
        End Set
    End Property

    Public Property SReferenciaAX As String
        Get
            Return _sMonoSKU
        End Get
        Set(value As String)
            _sMonoSKU = value
        End Set
    End Property

    Public Property SOrderAX As String
        Get
            Return _sOrderAX
        End Get
        Set(value As String)
            _sOrderAX = value
        End Set
    End Property

    Public Property OrderAXDate As Date
        Get
            Return _OrderAXDate
        End Get
        Set(value As Date)
            _OrderAXDate = value
        End Set
    End Property

    Public Property SOrderLineAx As String
        Get
            Return _sOrderLineAx
        End Get
        Set(value As String)
            _sOrderLineAx = value
        End Set
    End Property

    Public Property OrderLineAXDate As Date
        Get
            Return _OrderLineAXDate
        End Get
        Set(value As Date)
            _OrderLineAXDate = value
        End Set
    End Property

    Public Property IQuantity As Integer
        Get
            Return _iQuantity
        End Get
        Set(value As Integer)
            _iQuantity = value
        End Set
    End Property

    Public Property BClientBlock As Integer
        Get
            Return _bClientBlock
        End Get
        Set(value As Integer)
            _bClientBlock = value
        End Set
    End Property

    Public Property SNotes As String
        Get
            Return _sNotes
        End Get
        Set(value As String)
            _sNotes = value
        End Set
    End Property

    Public Property RequestedDate As Date
        Get
            Return _RequestedDate
        End Get
        Set(value As Date)
            _RequestedDate = value
        End Set
    End Property

    Public Property DesiredDate As Date
        Get
            Return _DesiredDate
        End Get
        Set(value As Date)
            _DesiredDate = value
        End Set
    End Property

    Public Property IId_states As Integer
        Get
            Return _Iid_states
        End Get
        Set(value As Integer)
            _Iid_states = value
        End Set
    End Property

    Public Property IId As Integer
        Get
            Return _iId
        End Get
        Set(value As Integer)
            _iId = value
        End Set
    End Property

    Public Property IIdArticles As Integer
        Get
            Return _iIdArticles
        End Get
        Set(value As Integer)
            _iIdArticles = value
        End Set
    End Property

    Public Property BIsDeleted As Integer
        Get
            Return _bIsDeleted
        End Get
        Set(value As Integer)
            _bIsDeleted = value
        End Set
    End Property

#End Region

End Class
