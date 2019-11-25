Public Class Product
    ' type in property and hit tab twice to get a template laid out for you as follows

    ' private property as follows:
    Private _IsActive As Boolean
    Public Property IsActive() As Boolean
        ' Get method
        Get
            Return _IsActive
        End Get
        ' Set method
        Set(ByVal value As Boolean)
            _IsActive = value
        End Set
    End Property

    Private _Name As String
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property

    Private _ProductNumber As String
    Public Property ProductNumber() As String
        Get
            Return _ProductNumber
        End Get
        Set(ByVal value As String)
            _ProductNumber = value
        End Set
    End Property
End Class
