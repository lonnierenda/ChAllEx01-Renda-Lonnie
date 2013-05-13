'Copyright (c) 2009-2013 Dan Turk

#Region "Class / File Comment Header block"
'Program:               ChAllEx01  
'File:                  ClsBusiness
'Author:                Lonnie Renda
'Description:           Creates the business class.  All the major work is done in this class.  This includes the
'                       adding of inventory, customers, the business, and sales.  It also preforms all the calculations
'                       for any totals.
'Date:                  April 14, 2013
'Tier:                  Business Logic
'Exceptions:            None generated
'Exception-Handling:    This class handles errors that try to create inventory, sale or customer, before a business is
'                       created.  It also assures the item and customer exist before a sale is made.  It also ensures
'                       that duplicate customers and inventory items are not created.  Finally, it makes sure the
'                       quantity of an item is available before it is sold.
'Events:                None generated
'Event-Handling:        None
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
Imports System.IO
#End Region 'Option / Imports

Public Class Business

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************


    '********** Module-level constants
    'Module-level constants
    'customer array variables
    Private mCustomer_ARRAY_SIZE_DEFAULT As Integer = 10
    Private mCustomer_ARRAY_INCREMENT_DEFAULT As Integer = 5

    'inventory array variables
    Private mInventory_ARRAY_SIZE_DEFAULT As Integer = 10
    Private mInventory_ARRAY_INCREMENT_DEFAULT As Integer = 5

    'Sales array variables
    Private mSale_ARRAY_SIZE_DEFAULT As Integer = 10
    Private mSale_ARRAY_Increment_DEFAULT As Integer = 5

    
    '********** Module-level variables
    'business class variables
    Private mBusinessName As String
    Private mBusinessType As String
    Private mBusinessCreationDate As Date

    'customer tracking variables
    Private mCustomers() As Customer
    Private mMaxCustomers As Integer
    Private mNumCustomers As Integer

    'inventory item variables
    Private mInventoryItems() As Inventory
    Private mMaxInventoryItems As Integer
    Private mNumInvenotryItems As Integer
    Private mCostofInventory As Decimal
    Private mTotalInventoryQuantity As Integer

    'sales variables
    Private mSales() As Sale
    Private mMaxSales As Integer
    Private mNumSales As Integer
    Private mCostofSales As Decimal = CDec(0.0)
    Private mTotalSalesAmount As Decimal = CDec(0.0)

    Private mTransactions As String
    Private mNumberofTransactions As Integer = 0
    







#End Region 'Attributes

#Region "Constructors"
    '******************************************************************
    'Constructors
    '******************************************************************

    'No Constructors are currently defined.

    '********** Default constructor
    '             - no parameters

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    'Public Sub New( _
    '        ByVal pName As String, _
    '        ByVal pType As String, _
    '        ByVal pDate As Date
    '              )



    '    MyBase.New()

    '    _BusinessName = pName
    '    _BusinessType = pType
    '    _BusinessCreationDate = pDate

    '    _maxCustomers = mCustomer_ARRAY_SIZE_DEFAULT
    '    ReDim mCustomers(_maxCustomers - 1)
    '    _numCustomers = 0

    '    _maxInventoryItems = mInventory_ARRAY_SIZE_DEFAULT
    '    ReDim mInventoryItems(_maxInventoryItems - 1)
    '    _numInventoryItems = 0






    'End Sub 'New Business

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class


    'Public Sub New(ByVal pBusiness As Business)

    '    'Copy constructor.
    '    'Clone pStudent.

    '    MyBase.New()

    '    _BusinessName = pBusiness.BusinessName
    '    _BusinessType = pBusiness.BusinessType
    '    _BusinessCreationDate = pBusiness.BusinessCreationDate



    'End Sub 'New(pbusiness)

#End Region 'Constructors

#Region "Get/Set Methods"


   
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    'No Get/Set Methods are currently defined.

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public ReadOnly Property BusinessName As String
        Get
            Return _BusinessName
        End Get
    End Property

    Public ReadOnly Property BusinessType As String
        Get
            Return _BusinessType
        End Get
    End Property

    Public ReadOnly Property BusinessCreationDate As Date
        Get
            Return _BusinessCreationDate
        End Get
    End Property

    Public ReadOnly Property numCustomers As Integer
        Get
            Return _numCustomers
        End Get
    End Property

    Public ReadOnly Property numInventory As Integer
        Get
            Return _numInventoryItems
        End Get
    End Property

    Public ReadOnly Property numTransactions As Integer
        Get
            Return _numTransactions
        End Get
    End Property

    Public ReadOnly Property totalInventoryCost As Decimal
        Get
            Return _TotalInventoryCost
        End Get
    End Property

    Public ReadOnly Property totalInventoryQuantity As Integer
        Get
            Return _TotalInventoryQuantity
        End Get
    End Property
    Public ReadOnly Property numSales As Integer
        Get
            Return _numSales
        End Get
    End Property

    Public ReadOnly Property CostofSales As Decimal
        Get
            Return _costofSales
        End Get
    End Property

    Public ReadOnly Property TotalSalesAmount As Decimal
        Get
            Return _totalSalesAmount
        End Get
    End Property

    Public ReadOnly Property ithInventoryItem(ByVal pN As Integer) As Inventory
        Get
            Return _ithInventoryItem(pN)
        End Get
    End Property

    Public ReadOnly Property ithSale(ByVal pN As Integer) As Sale
        Get
            Return _ithSale(pN)
        End Get
    End Property

    Public ReadOnly Property ithCustomer(ByVal pN As Integer) As Customer
        Get
            Return _ithCustomer(pN)
        End Get
    End Property

    
    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _BusinessName As String
        Get
            Return mBusinessName
        End Get
        Set(ByVal pValue As String)
            mBusinessName = pValue
        End Set
    End Property

    Private Property _BusinessType As String
        Get
            Return mBusinessType
        End Get
        Set(ByVal pValue As String)
            mBusinessType = pValue
        End Set
    End Property

    Private Property _BusinessCreationDate As Date
        Get
            Return mBusinessCreationDate
        End Get
        Set(ByVal pValue As Date)
            mBusinessCreationDate = pValue
        End Set
    End Property

    Private Property _numCustomers As Integer
        Get
            Return mNumCustomers
        End Get
        Set(ByVal pValue As Integer)
            mNumCustomers = pValue
        End Set
    End Property

    Private Property _numInventoryItems As Integer
        Get
            Return mNumInvenotryItems
        End Get
        Set(ByVal pValue As Integer)
            mNumInvenotryItems = pValue
        End Set
    End Property

    Private Property _numSales As Integer
        Get
            Return mNumSales
        End Get
        Set(ByVal pValue As Integer)
            mNumSales = pValue
        End Set
    End Property

    Private Property _numTransactions As Integer
        Get
            Return mNumberofTransactions
        End Get
        Set(ByVal pValue As Integer)
            mNumberofTransactions = pValue
        End Set
    End Property

    Private Property _Transactions As String
        Get
            Return mTransactions
        End Get
        Set(ByVal pValue As String)
            mTransactions = pValue
        End Set
    End Property

    Private Property _costofSales As Decimal
        Get
            Return mCostofSales
        End Get
        Set(ByVal pValue As Decimal)
            mCostofSales = pValue
        End Set
    End Property

    Private Property _totalSalesAmount As Decimal
        Get
            Return mTotalSalesAmount
        End Get
        Set(ByVal pValue As Decimal)
            mTotalSalesAmount = pValue
        End Set
    End Property

    Private Property _maxCustomers As Integer
        Get
            Return mMaxCustomers
        End Get
        Set(ByVal pValue As Integer)
            mMaxCustomers = pValue
        End Set
    End Property

    Private Property _maxSales As Integer
        Get
            Return mMaxSales
        End Get
        Set(ByVal pValue As Integer)
            mMaxSales = pValue
        End Set
    End Property

    Private Property _maxInventoryItems As Integer
        Get
            Return mMaxInventoryItems
        End Get
        Set(ByVal pValue As Integer)
            mMaxInventoryItems = pValue
        End Set
    End Property

    Private Property _TotalInventoryCost As Decimal
        Get
            Return mCostofInventory
        End Get
        Set(ByVal pvalue As Decimal)
            mCostofInventory = pvalue
        End Set
    End Property

    Private Property _TotalInventoryQuantity As Integer
        Get
            Return mTotalInventoryQuantity
        End Get
        Set(ByVal pValue As Integer)
            mTotalInventoryQuantity = pValue
        End Set
    End Property

    Private Property _ithCustomer(ByVal pN As Integer) As Customer
        Get
            Return mCustomers(pN)
        End Get
        Set(ByVal pValue As Customer)
            mCustomers(pN) = pValue
        End Set
    End Property

    Private Property _ithInventoryItem(ByVal pN As Integer) As Inventory
        Get
            Return mInventoryItems(pN)
        End Get
        Set(ByVal pValue As Inventory)
            mInventoryItems(pN) = pValue
        End Set
    End Property

    Private Property _ithSale(ByVal pN As Integer) As Sale
        Get
            Return mSales(pN)
        End Get
        Set(ByVal pValue As Sale)
            mSales(pN) = pValue
        End Set
    End Property


#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    'No Behavioral Methods are currently defined.

    '********** Public Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    Public Sub addBusiness(ByVal pBusinessName As String, ByVal pBusinessType As String, ByVal pBusinessDate As Date)
        Try
            _addBusiness(pBusinessName, pBusinessType, pBusinessDate)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub addCustomer(ByVal pCustomer As Customer)

        'add customer to the customer array

        Try
            _addCustomer(pCustomer)
        Catch ex As Exception
            Throw New Business_CustomerAdd_Exception(pCustomer.name)
        End Try

    End Sub 'addCustomer(pCustomer)

    Public Sub addInventoryItem(ByVal pInventory As Inventory)

        'add inventory item to the inventory array

        Try
            _addInventoryItem(pInventory)
        Catch ex As Exception
            Throw New Business_ItemAdd_Exception(pInventory.itemID)
        End Try

    End Sub 'addInventory(pInventory)

    Public Sub addSale(ByVal pSale As Sale)

        'Add Sale to the sale array

        Try
            _addSale(pSale)
        Catch ex As Exception
            Throw New Business_SaleAdd_Exception(pSale.CustomerName, pSale.itemID, pSale.Quantity)
        End Try

    End Sub 'addSale

    Public Function findCustomerByName(ByVal pName As String) As Customer

        'findcustomerByname() searches for pname in the customer array.
        'If found, it returns a reference to this customer;
        'if not found, it returns Nothing.


        Dim foundCustomer As Customer = Nothing
        Dim theCustomer As Customer
        Dim i As Integer

        For i = 0 To _numCustomers - 1
            theCustomer = _ithCustomer(i)
            If theCustomer.name = pName Then
                foundCustomer = theCustomer     'or return _thecustomer
                '                                or exit for
            End If
        Next i

        Return foundCustomer

    End Function 'findcustomerbyname

    Public Function findSaleByCustomerName(ByVal pName As String) As Sale

        'findsalebycustomername searches for pName in the sales array.
        'If found, it returns a reference to this sale;
        'if not found, it returns Nothing.

        Dim foundSale As Sale = Nothing
        Dim theSale As Sale
        Dim i As Integer

        For i = 0 To _numSales - 1
            theSale = _ithSale(i)
            If theSale.CustomerName = pName Then
                foundSale = theSale     'or return _thesale
                '                                or exit for
            End If
        Next i

        Return foundSale

    End Function 'findsaleby customername

    Public Function findInventoryItemByID(ByVal pItemID As String) As Inventory

        'findInventoryByID() searches for pID in the inventory array.
        'If found, it returns a reference to this inventory item;
        'if not found, it returns Nothing.

        Dim foundInventoryItem As Inventory = Nothing
        Dim theInventoryItem As Inventory
        Dim i As Integer

        For i = 0 To _numInventoryItems - 1
            theInventoryItem = _ithInventoryItem(i)
            If theInventoryItem.itemID = pItemID Then
                foundInventoryItem = theInventoryItem     'or return _theitem
                '                                or exit for
            End If
        Next i

        Return foundInventoryItem

    End Function 'findinventoryitembyID

    Public Sub readFromFile(ByVal pFileName As String)


        Try
            _readFromFile(pFileName)
        Catch ex As Exception
            Throw ex
        End Try
        

    End Sub 'readFromFile(pFileName)

    Public Sub writeToFile(ByVal pFileName As String)

        Try
            _writeToFile(pFileName)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub 'writeToFile(pFileName,pAppend)

    Public Overrides Function ToString() As String

        'ToString() creates and formats a string that holds the
        'information about this class.

        Dim tmpStr As String

        tmpStr = _
            "( Business: " _
            & "Name=" & _BusinessName.ToString _
            & ", Type=" & _BusinessType.ToString _
             & ", Date Created=" & _BusinessCreationDate.ToString _
            & " )"

        Return tmpStr

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods
    Private Sub _addBusiness(ByVal pbusinessName As String, ByVal pbusinessType As String, ByVal pbusinessDate As Date)

        If Not BusinessName Is Nothing Then
            MessageBox.Show("Could not add business. The business " & BusinessName & " already exists. Thus, " & pbusinessName & " not added.")
            Exit Sub
        End If
        

        'update variables
        _BusinessName = pbusinessName
        _BusinessType = pbusinessType
        _BusinessCreationDate = pbusinessDate

        'update transactions

        _Transactions += "business," & "create," & pbusinessName & "," & pbusinessType & "," & pbusinessDate.ToString & vbCrLf

        _numTransactions += 1

        'Raise the event to update frmmain

        RaiseEvent Business_Added( _
           Me,
           New Business_Added_EventArgs( _
               pbusinessName, pbusinessType, pbusinessDate _
               ) _
           )




    End Sub '_addBusiness

    Private Sub _addCustomer(ByVal pCustomer As Customer)

        'chekc to see if business exists
        If BusinessName Is Nothing Then
            MessageBox.Show("Sorry a business must exist before you add a customer.")
            Exit Sub
        End If

        'check to see if customer already exists before adding.

        Dim checkname As String
        checkname = pCustomer.name

        If Not findCustomerByName(checkname) Is Nothing Then
            MessageBox.Show("Could not add customer. The customer " & pCustomer.name & " already exists.")
            Exit Sub
        End If
        
        'create an instance of the class
        Dim newCustomer As Customer = New Customer(pCustomer)

        'check size of array and increase if needed.

        If _numCustomers >= _maxCustomers Then
            _maxCustomers += mCustomer_ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mCustomers(_maxCustomers - 1)
        End If

        'update number of customers 
        _ithCustomer(_numCustomers) = newCustomer
        _numCustomers += 1

        'update transactions array

        _Transactions += "customer," & "add," & newCustomer.name & "," & newCustomer.address & "," _
            & newCustomer.city & "," & newCustomer.state & "," & newCustomer.zip.ToString & vbCrLf
        _numTransactions += 1

        'raise event to update frm main
        RaiseEvent Business_CustomerAdded( _
            Me,
            New ClsBusiness_CustomerAdded_EventArgs( _
                newCustomer _
                ) _
            )

    End Sub 'addCustomer

    Private Sub _addInventoryItem(ByVal pInventory As Inventory)

        'check to see if business exists
        If BusinessName Is Nothing Then
            MessageBox.Show("Sorry a business must exist before you can add an inventory item.")
            Exit Sub
        End If

        'check to see if inventory item exists before adding inventory item
        Dim checkID As String
        checkID = pInventory.itemID

        If Not findInventoryItemByID(checkID) Is Nothing Then
            MessageBox.Show("Sorry the item " & checkID & " exists.  Thus you cannot add it again.")
            Exit Sub
        End If


        'create an instance of the class

        Dim newInventory As Inventory = New Inventory(pInventory)

        'check array size and increase if needed

        If _numInventoryItems >= _maxInventoryItems Then
            _maxInventoryItems += mInventory_ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mInventoryItems(_maxInventoryItems - 1)
        End If

        _ithInventoryItem(_numInventoryItems) = newInventory

        'Update total items in inventory

        _numInventoryItems += 1


        'Update total inventory cost and quantity

        mCostofInventory = mCostofInventory + (newInventory.itemCost * newInventory.itemQuantity)

        _TotalInventoryQuantity = _TotalInventoryQuantity + (newInventory.itemQuantity)

        _Transactions += "inventory," & "add," & newInventory.itemID & "," & newInventory.itemDescription & "," _
            & newInventory.itemCost.ToString & "," & newInventory.itemPrice.ToString & "," & newInventory.itemQuantity.ToString & vbCrLf
        _numTransactions += 1

        'Raise event to update frmMain
        RaiseEvent Business_InventoryItemAdded( _
            Me,
            New Business_InventoryItemAdded_EventArgs(newInventory))


    End Sub

    Private Sub _addSale(ByVal pSale As Sale)

        'check to see if business exists
        If BusinessName Is Nothing Then
            MessageBox.Show("Sorry a business must exist before you can add a sale.")
            Exit Sub
        End If

        'declare local variables
        Dim selectedCustomer As Customer
        Dim selectedInventory As Inventory
        Dim name As String

        'Check to see if customer exists before adding sale.
        name = pSale.CustomerName

        If findCustomerByName(name) Is Nothing Then
            MessageBox.Show("Sorry the customer " & name & " does not exist.  Thus you cannot add a sale.")
            Exit Sub
        End If

        

        'get inventory information to see if item exists and enough of  quantity before sale is made.
        If findInventoryItemByID(pSale.itemID) Is Nothing Then
            MessageBox.Show("Sorry the item " & pSale.itemID & " does not exist.  Thus you cannot add a sale.")
            Exit Sub
        End If

        'check to see if there is enough of the quantity
        If pSale.Quantity > findInventoryItemByID(pSale.itemID).itemQuantity Then
            MessageBox.Show("Sorry, there is not enough of " & pSale.itemID & " for this sale.")
            Exit Sub
        End If


        'create instance of the sale

        Dim newSale As Sale = New Sale(pSale)

        'check and increase array size if needed
        If _numSales >= _maxSales Then
            _maxSales += mSale_ARRAY_Increment_DEFAULT
            ReDim Preserve mSales(_maxSales - 1)
        End If

        'update number totals

        _ithSale(_numSales) = newSale
        _numSales += 1

        'get customer info
        'Dim selectedCustomer As Customer
        selectedCustomer = findCustomerByName(newSale.CustomerName)

        'get inventory info
        'Dim selectedInventory As Inventory
        selectedInventory = findInventoryItemByID(newSale.itemID)

        'Update total sales, cost, revenue
        _costofSales += (newSale.Quantity * selectedInventory.itemCost)
        _totalSalesAmount += (newSale.Quantity * selectedInventory.itemPrice)

        'Update total quantity and cost
        _TotalInventoryQuantity = _TotalInventoryQuantity - newSale.Quantity
        _TotalInventoryCost = _TotalInventoryCost - (newSale.Quantity * selectedInventory.itemCost)

        'update individual inventory items quantity
        selectedInventory.itemQuantity = selectedInventory.itemQuantity - newSale.Quantity

        'update transactions and transactions array
        _Transactions += "sale," & "add," & newSale.CustomerName & "," & newSale.itemID & "," _
            & newSale.Quantity.ToString & vbCrLf
        _numTransactions += 1


        'Raise event to update frmmain
        RaiseEvent Business_SaleMade( _
            Me,
            New Business_SaleMade_EventArgs(newSale) _
            )


    End Sub 'SaleMade


    Private Sub _readFromFile(ByVal pFileName As String)

        Dim inputFile As StreamReader
        Dim trxLine As String
        Dim trxParts() As String
        Dim trxType As String
        Dim trxBusinessName As String
        Dim trxBusinessDescription As String
        Dim trxCreationDate As Date
        Dim trxCustomerName As String
        Dim trxCustomerAddress As String
        Dim trxCustomerCity As String
        Dim trxCustomerState As String
        Dim trxCustomerZip As Integer
        Dim trxItemID As String
        Dim trxItemDescription As String
        Dim trxItemCost As Decimal
        Dim trxItemPrice As Decimal
        Dim trxItemQuantity As Integer
        Dim trxSaleName As String
        Dim trxSaleItem As String
        Dim trxSaleQuantity As Integer
        Dim trxAction As String

        Dim i As Integer

        Try
            inputFile = New StreamReader(pFileName)
        Catch ex As Exception
            'Throw ex
            Throw New Business_CouldNotOpenFile_Exception(pFileName)
        End Try 'New Streamreader

        Do While Not inputFile.EndOfStream
            
            trxLine = inputFile.ReadLine

            trxParts = Split(trxLine, ",")
            For i = 0 To trxParts.Length - 1
                trxParts(i) = Trim(trxParts(i))
            Next i

            'parse transactions based on case
            trxType = trxParts(0)
            trxAction = trxParts(1).ToString

            




            Select Case trxType.ToUpper

                Case "BUSINESS"
                    trxBusinessName = trxParts(2).ToString
                    trxBusinessDescription = trxParts(3).ToString
                    Try
                        trxCreationDate = CDate(trxParts(4))

                    Catch ex As Exception
                        Throw ex
                    End Try

                    _addBusiness(trxBusinessName, trxBusinessDescription, trxCreationDate)

                Case "CUSTOMER"
                    trxCustomerName = trxParts(2).ToString
                    trxCustomerAddress = trxParts(3).ToString
                    trxCustomerCity = trxParts(4).ToString
                    trxCustomerState = trxParts(5).ToString
                    Try
                        trxCustomerZip = CInt(trxParts(6))
                    Catch ex As Exception

                    End Try


                    _addCustomer( _
                        New Customer(trxCustomerName, trxCustomerAddress, trxCustomerCity, trxCustomerState, CInt(trxCustomerZip)) _
                        )

                Case "INVENTORY"
                    trxItemID = trxParts(2).ToString
                    trxItemDescription = trxParts(3).ToString
                    Try
                        trxItemCost = CDec(trxParts(4))
                    Catch ex As Exception
                    End Try
                    Try
                        trxItemPrice = CDec(trxParts(5))
                    Catch ex As Exception
                        Throw ex
                    End Try
                    Try
                        trxItemQuantity = CInt(trxParts(6))
                    Catch ex As Exception
                        Throw ex
                    End Try

                    '
                    _addInventoryItem( _
                        New Inventory(trxItemID, trxItemDescription, CDec(trxItemCost), CDec(trxItemPrice), CInt(trxItemQuantity)) _
                        )


                Case "SALE"
                    trxSaleName = trxParts(2).ToString
                    trxSaleItem = trxParts(3).ToString
                    Try
                        trxSaleQuantity = CInt(trxParts(4))

                    Catch ex As Exception

                    End Try

                    _addSale( _
                        New Sale(trxSaleName, trxSaleItem, CInt(trxSaleQuantity)) _
                        )

                Case Else
                    MessageBox.Show("Unknown type." & trxType & " not recognized. Cannot read from file.")

            End Select
            
        Loop

        inputFile.Close()

    End Sub

    Private Sub _writeToFile( _
            ByVal pFileName As String
            )

        Dim outputFile As StreamWriter
        'Dim i As Integer

        Try
            outputFile = New StreamWriter(pFileName)
        Catch ex As Exception
            'Throw ex
            Throw New Business_CouldNotWriteFile_Exception(pFileName)
        End Try 'New Streamreader


        outputFile.WriteLine(_Transactions)


        outputFile.Close()

    End Sub
#End Region 'Behavioral Methods

#Region "Event Procedures"
    '******************************************************************
    'Event Procedures
    '******************************************************************

    'No Event Procedures are currently defined

    '********** User-Interface Event Procedures
    '             - Initiated explicitly by user

    '********** User-Interface Event Procedures
    '             - Initiated automatically by system

    '********** Business Logic Event Procedures
    '             - Initiated as a result of business logic
    '               method(s) running

#End Region 'Event Procedures

#Region "Events"
    '******************************************************************
    'Events
    '******************************************************************
    Public Event Business_Added( _
        ByVal sender As Object, _
        ByVal e As EventArgs)

    Public Event Business_CustomerAdded( _
        ByVal sender As Object, _
        ByVal e As EventArgs)

    Public Event Business_InventoryItemAdded( _
        ByVal sender As Object, _
        ByVal e As EventArgs)

    Public Event Business_SaleMade( _
        ByVal sender As Object, _
        ByVal e As EventArgs)

#End Region 'Events


End Class 'Business

