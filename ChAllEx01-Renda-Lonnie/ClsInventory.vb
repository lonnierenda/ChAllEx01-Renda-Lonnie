'Copyright (c) 2009-2013 Dan Turk

#Region "Class / File Comment Header block"
'Program:               ChAllEx01  
'File:                  ClsInventory
'Author:                Lonnie Renda
'Description:           Creates the Inventory class.
'Date:                  April 14, 2013
'Tier:                  Business Logic
'Exceptions:            None generated
'Exception-Handling:    None
'Events:                None generated
'Event-Handling:        None
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
#End Region 'Option / Imports

Public Class Inventory

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables
    Private mItemID As String
    Private mItemDescription As String
    Private mItemCost As Decimal
    Private mItemSalePrice As Decimal
    Private mItemQuantity As Integer


#End Region 'Attributes

#Region "Constructors"
    'Private _pInventory As Inventory
    '******************************************************************
    'Constructors
    '******************************************************************

    'No Constructors are currently defined.

    '********** Default constructor
    '             - no parameters

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    Public Sub New( _
        ByVal pItemID As String, _
        ByVal pItemDescription As String, _
        ByVal pItemCost As Decimal, _
        ByVal pItemSalePrice As Decimal, _
        ByVal pItemQuantity As Integer)

        'Special constructor.
        'Uses parameters to initialize all attributes.

        MyBase.New()

        _ItemID = pItemID
        _ItemDescription = pItemDescription
        _ItemCost = pItemCost
        _ItemSalePrice = pItemSalePrice
        _ItemQuantity = pItemQuantity

    End Sub 'New(pID,pName,pPrice,pQoH)

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

    Public Sub New(ByVal pInventory As Inventory)

        'Copy constructor.
        'Clones pProduct.

        Me.New( _
            pInventory.itemid, _
            pInventory.itemDescription, _
            pInventory.itemCost, _
            pInventory.itemPrice, _
            pInventory.itemQuantity _
           )

    End Sub 'New(pProduct)

#End Region 'Constructors

#Region "Get/Set Methods"


    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    'No Get/Set Methods are currently defined.

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public ReadOnly Property itemID As String
        Get
            Return _ItemID
        End Get
    End Property


    Public ReadOnly Property itemDescription As String
        Get
            Return _ItemDescription
        End Get
    End Property

    Public ReadOnly Property itemCost As Decimal
        Get
            Return _ItemCost
        End Get
    End Property

    Public ReadOnly Property itemPrice As Decimal
        Get
            Return _ItemSalePrice
        End Get
    End Property

    Public Property itemQuantity As Integer
        Get
            Return _ItemQuantity
        End Get
        Set(ByVal pValue As Integer)
            _ItemQuantity = pValue
        End Set
    End Property




    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _ItemID As String
        Get
            Return mItemID
        End Get
        Set(ByVal pValue As String)
            mItemID = pValue
        End Set
    End Property

    Private Property _ItemDescription As String
        Get
            Return mItemDescription
        End Get
        Set(ByVal pValue As String)
            mItemDescription = pValue
        End Set
    End Property

    Private Property _ItemCost As Decimal
        Get
            Return mItemCost
        End Get
        Set(ByVal pValue As Decimal)
            mItemCost = pValue
        End Set
    End Property

    Private Property _ItemSalePrice As Decimal
        Get
            Return mItemSalePrice
        End Get
        Set(ByVal pValue As Decimal)
            mItemSalePrice = pValue
        End Set
    End Property

    Private Property _ItemQuantity As Integer
        Get
            Return mItemQuantity
        End Get
        Set(ByVal pValue As Integer)
            mItemQuantity = pValue
        End Set
    End Property

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************



    '********** Public Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    


    Public Overrides Function ToString() As String

        'ToString() generates a String version of the object's data.

        Dim _tmpStr As String

        _tmpStr = "( PRODUCT: " _
            & "Item Description=" & _ItemDescription _
            & ", Item Cost=" & _ItemCost.ToString("C") _
            & ", Item Sale Price=" & _ItemSalePrice.ToString("C") _
            & ", Quantity=" & _ItemQuantity.ToString() _
            & " )"

        Return _tmpStr

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods

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

    'No Events are currently defined.

#End Region 'Events


End Class 'Inventory


