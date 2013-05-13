'Copyright (c) 2009-2013 Dan Turk

#Region "Class / File Comment Header block"
'Program:               ChAllEx01  
'File:                  ClsSale
'Author:                Lonnie Renda
'Description:           Creates the Sale class.
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

Public Class Sale

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables

    Private mCustomerName As String
    Private mItemID As String
    Private mQuantity As Integer

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

    Public Sub New( _
           ByVal pCustomerName As String, _
           ByVal pItemID As String, _
           ByVal pQuantity As Integer
                 )



        MyBase.New()

        _customername = pCustomerName
        _itemID = pItemID
        _quantity = pQuantity


    End Sub 'New(pProduct)

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

    Public Sub New(ByVal pSale As Sale)

        'Copy constructor.
        'Clones pProduct.

        Me.New( _
            pSale.CustomerName, _
            pSale.itemID, _
            pSale.Quantity _
            )

    End Sub 'New(pProduct)

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    
    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public ReadOnly Property CustomerName As String
        Get
            Return _CustomerName
        End Get
    End Property

    Public ReadOnly Property itemID As String
        Get
            Return _ItemID
        End Get
    End Property

    Public ReadOnly Property Quantity As Integer
        Get
            Return _Quantity
        End Get
    End Property


    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _CustomerName As String
        Get
            Return mCustomerName
        End Get
        Set(ByVal pValue As String)
            mCustomerName = pValue
        End Set
    End Property

    Private Property _ItemID As String
        Get
            Return mItemID
        End Get
        Set(ByVal pValue As String)
            mItemID = pValue
        End Set
    End Property

    Private Property _Quantity As Integer
        Get
            Return mQuantity
        End Get
        Set(ByVal pValue As Integer)
            mQuantity = pValue
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

    Public Overrides Function ToString() As String

        'ToString() generates a String version of the object's data.

        Dim _tmpStr As String

        _tmpStr = "( SALE: " _
            & "Customer=" & _CustomerName _
            & ", Item ID=" & _ItemID _
            & ", Quantity=" & _Quantity.ToString _
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

End Class 'Sale

