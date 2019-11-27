Public Class Product
    'Modify Product class to inherit from CommonBase
    'It now has all the properties from CommonBase too!
    Inherits CommonBase
    Sub New()
        'We can intialise default values for each of our properties at the start using this New() syntax
        'Standard assignments only, we want to avoid errors so keep it simple
        StandardCost = 500
        ListPrice = 900
        SellStartDate = DateTime.Now
    End Sub
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
    Public Property ProductID As Integer
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
    'Syntax:
    'Function <Name>(<ByVal/ByRef> <variable_name> As <data_type>) As <data_type_of_output>
    Function CalculateSellEndDate(ByVal days As Integer) As DateTime 'function returns a value, so we need to define a datatype for that value at the end
        SellEndDate = SellStartDate.AddDays(days) 'Calculate the value of SellEndDate, based on the inputted number of days

        Return SellEndDate 'return calculated value

        'Alternative way to return the desired value - set it to the Name of the function:
        'CalculateSellEndDate = SellEndDate
    End Function
    'Define an /OPTIONAL/ function:
    'In the case of making whether you pass a parameter optional
    'Syntax:
    'Function <Name>(///<Optional(Y/N?)>/// <ByVal/ByRef> <variable_name> As <data_type> ///<= VAL?>///) As <data_type_of_output>
    'The parts in /// are the new stuff we've built on to the previous syntaxt type. 
    'The first bit is something you include if you want to make it an optional function
    'The next bit is essentially the default value for the variable if the function is not passed
    '\\\\\\\replace next function with two below it\\\\\\\\\\
    'Function CalculateProfit(Optional ByVal newCost As Decimal = 0) As Decimal
    '    'To make sure that default value is something that you'd never pass, we use the condition to run the code if the variable is not equal to that value
    '    If newCost <> 0 Then
    '        'StandardCost is a property defined further up ^, we assign it in this particular function
    '        StandardCost = newCost
    '    End If

    '    'Profits
    '    Return ListPrice - StandardCost

    'End Function

    'We can use overloading methods for a more object-orientated approach
    'This function below calls the other version of the profit and passes the StandardCost property through it
    'This essentially means we can call the function without parsing an argument, and know it will just use whatever the standard cost is defined as at the time
    Overloads Function CalculateProfit() As Decimal
        Return CalculateProfit(StandardCost)
    End Function
    'this function below is called the same thing as the one above but it is considered different because it passes different variables through it
    Overloads Function CalculateProfit(ByVal newCost As Decimal) As Decimal
        'All properties are intialised to a default value when the class instance is initialised
        'The default value given at the initialisation of a class instance for a decimal is 0
        'This next block checks that the newCost variable given in the argument is not zero, and if it is not, it assigns that new variable to the StandardCost
        If newCost <> 0 Then
            StandardCost = newCost
        End If
        'return the value of the difference between ListPrice and StandardCost, as this is the profit
        Return ListPrice - StandardCost
    End Function

    'We can use shared functions to bypass the requirement of setting property values in a class
    'They just take input and operate based on those inputs
    Shared Function CalculateTheProfit(ByVal cost As Decimal, ByVal price As Decimal) As Decimal
        Return price - cost
    End Function



    Protected Overrides Function GetClassData() As String 'this function works the same as one defined in Customer.vb, with different properties
        Dim sb As New Text.StringBuilder(1024)

        'take string builder and append each of the following lines, which contain information stored in this class
        sb.AppendLine("Product ID: " + ProductID.ToString())
        sb.AppendLine("Product Name: " + Name)
        sb.AppendLine("Product Number: " + ProductNumber)
        'The following are inherited from CommonBase Class
        'They can be removed when using the overidable functions syntax, with the keyword OVERRIDES at the start of this function definition
        'sb.AppendLine("Is Active: " + IsActive.ToString())
        'sb.AppendLine("Modified Date: " + ModifiedDate.ToLongDateString())
        'sb.AppendLine("Created By: " + CreatedBy)
        'They are replaced by the following line, which accesses the function from the inherited class:
        'It is allowed to access it, even if the method is protected using "Protected" keyword, because it is in the inheritance chain
        sb.AppendLine(MyBase.GetClassData())
        'We have managed to eliminate a few lines of code here and saved ourselves some repeating of ourselves. 
        'We have managed to keep ourselves DRY...

        Return sb.ToString()
    End Function

End Class
