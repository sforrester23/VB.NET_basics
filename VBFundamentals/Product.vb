Public Class Product
    ' type in property and hit tab twice to get a template laid out for you as follows

    'private property as follows:
    'Private _IsActive As Boolean
    'public property as follows, use tab to tab through variables
    '(this is the long way of writing a property, which has its benefits and drawbacks
    'Public Property IsActive() As Boolean
    '    ' Get method
    '    Get
    '        Return _IsActive
    '    End Get
    '    ' Set method
    '    Set(ByVal value As Boolean)
    '        _IsActive = value
    '    End Set
    'End Property

    'Private _Name As String
    'Public Property Name() As String
    '    Get
    '        Return _Name
    '    End Get
    '    Set(ByVal value As String)
    '        _Name = value
    '    End Set
    'End Property

    'Private _ProductNumber As String
    'Public Property ProductNumber() As String
    '    Get
    '        Return _ProductNumber
    '    End Get
    '    Set(ByVal value As String)
    '        _ProductNumber = value
    '    End Set
    'End Property

    'Auto-implemented properties:
    Public Property Name As String
    Public Property ProductNumber As String
    Public Property Colour As String
    Public Property StandardCost As Decimal
    Public Property ListPrice As Decimal
    Public Property Size As String
    Public Property Weight As Decimal
    Public Property SellStartDate As DateTime
    Public Property SellEndDate As DateTime

    'Method definition:
    'Sub CalculateSellEndDate(ByVal days As Integer, ByRef sellDate As DateTime) 'ByVal means you can't change the subparameter "days" among this block of code
    '    'Not returning any value: no "return" key word
    '    'ByRef means you can edit the variable outside of the class structure, i.e. if it changed here in this class procedure, it will change it in the calling routine. 
    '    'This can be dangerous, as it means it can be altered from the MainModule, for instance. 
    '    'Setting the Sell end date by adding a given/inputted number of days to the sell start date
    '    SellEndDate = SellStartDate.AddDays(days)
    '    'Set the ByRef parameter
    '    sellDate = SellEndDate
    '    'THE USE OF ByVal IS NOT BEST PRACTISE! A class should only modify its own properties and not affect any calling routines' values. This is why you should use ByRef sparingly. 
    'End Sub

    'Replace the Above COMMENTED OUT method with a function:
    Function CalculateSellEndDate(ByVal days As Integer) As DateTime 'function returns a value, so we need to define a datatype for that value
        SellEndDate = SellStartDate.AddDays(days) 'Calculate the value of SellEndDate, based on the inputted number of days

        Return SellEndDate 'return calculated value

        'Alternative way to return the desired value - set it to the Name of the function:
        'CalculateSellEndDate = SellEndDate

    End Function


End Class
