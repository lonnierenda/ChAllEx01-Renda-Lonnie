'Copyright (c) 2009-2013 Dan Turk

#Region "Class / File Comment Header block"
'Program:            ChAllEx01
'File:               ClsBusiness_CustomerAdd_Exception.vb
'Author:             Lonnie Renda
'Description:        Throws a customer exception when customer cannot be adde.
'Date:               2013 May 10
'                      - Created.
'Tier:               Error
'Exceptions:         Handles errors when trying to add a customer.
'Exception-Handling: none
'Events:             none
'Event-Handling:     Handles customer add errors..
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
#End Region 'Option / Imports

Public Class Business_SaleAdd_Exception
    Inherits System.Exception

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

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

    '********** Default constructor
    '             - no parameters

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    Public Sub New( _
            ByVal pCustomerName As String, ByVal pItemID As String, ByVal pQuantity As Integer _
            )

        'Special constructor - Create object from pStudent attribute.

        MyBase.New()

        _customername = pCustomerName
        _itemID = pItemID
        _quantity = pQuantity
    End Sub 'New(pFileName)

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public ReadOnly Property customername As String
        Get
            Return _customername
        End Get
    End Property

    Public ReadOnly Property ItemID As String
        Get
            Return _ItemID
        End Get
    End Property

    Public ReadOnly Property Quantity As Integer
        Get
            Return _quantity
        End Get
    End Property

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _customername As String
        Get
            Return mCustomerName
        End Get
        Set(ByVal pValue As String)
            mCustomerName = pValue
        End Set
    End Property

    Private Property _itemID As String
        Get
            Return mItemID
        End Get
        Set(ByVal pValue As String)
            mItemID = pValue
        End Set
    End Property

    Private Property _quantity As Integer
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

    '********** Public Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    Public Overrides Function ToString() As String

        Return _toString()

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods

    Private Function _toString() As String

        Dim tmpStr As String

        tmpStr = _
            "( Business SALE_ADD_EXCEPTION: " _
            & "Customer Name=" & _customername _
            & "Item ID=" & _itemID _
            & "Quantity=" & _quantity _
            & " )"

        Return tmpStr

    End Function '_toString()

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

End Class 'Business_CouldNotWriteFile_Exception

