Module MainModule
    Dim Name As String = "Mountain Bike"
    Sub Main()
        'Objects()
        Main4()
        Console.ReadKey()
    End Sub
    'Sub Main0()
    '    Dim Name As String = "10 Speed Bike"
    '    ' local variables inside sub "shadow" global variables that are further away 
    '    Console.WriteLine(Name)

    '    Console.ReadKey()
    'End Sub
    'Sub Main1()
    '    Console.WriteLine(Name)
    'End Sub

    'Sub Main2()
    '    If True Then
    '        Dim Name As String = "Tricycle"
    '        Dim ListPrice As Decimal = 59.99
    '    End If

    '    Console.WriteLine(Name)
    '    ' The following does not compile properly, because the if statement above is seen a different block of code.
    '    ' The variable is initialised inside the if statement and expires when the if statements ends
    '    ' Console.WriteLine(ListPrice)

    '    Console.ReadKey()
    'End Sub
    'Sub Main3()
    '    IncrementListPrice()
    '    IncrementListPrice()
    '    IncrementListPrice()

    '    Console.ReadKey()
    'End Sub

    'Sub IncrementListPrice()
    '    Static ListPrice As Decimal = 0

    '    ListPrice = ListPrice + 1

    '    Console.WriteLine(ListPrice)
    'End Sub
    'Sub Objects()
    '    Dim theData As Object
    '    ' This is a rarely the right data type to be using. It is very memory intensive and is slow to access and slow to use, avoid where you can but know it is available.

    '    theData = "10 Speed Bike"
    '    Console.WriteLine(theData)

    '    theData = 999.99
    '    Console.WriteLine(theData)

    '    theData = DateTime.Now
    '    Console.WriteLine(theData)

    '    theData = vbEmpty
    '    Console.WriteLine(theData)

    '    theData = DBNull.Value
    '    Console.WriteLine(theData)
    'End Sub
    'Sub Main4()
    '    Dim isActive As Boolean = ClassConstants.DEFAULT_ACTIVE
    '    Dim Name As String = "10 Speed Bike"
    '    Dim ListPrice As Decimal
    '    Console.WriteLine(isActive)

    '    ListPrice = ClassConstants.DEFAULT_LIST_PRICE
    '    Console.WriteLine(ListPrice)
    '    Console.ReadKey()
    'End Sub
    Sub Maths()
        Dim ListPrice As Decimal = 999.99D

        ListPrice += 1
        Console.WriteLine(List)
    End Sub
End Module
