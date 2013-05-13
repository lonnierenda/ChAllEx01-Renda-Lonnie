'Copyright (c) 2009-2013 Dan Turk

#Region "Class / File Comment Header block"
'Program:               ChAllEx01  
'File:                  ClsCustomer
'Author:                Lonnie Renda
'Description:           Creates the customer class.
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

Public Class Customer

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables

    Private mCustomerName As String
    Private mCustomerAddress As String
    Private mCustomerCity As String
    Private mCustomerState As String
    Private mCustomerZip As Integer




#End Region 'Attributes

#Region "Constructors"
    'Private _pCustomer As Customer
    '******************************************************************
    'Constructors
    '******************************************************************



    '********** Default constructor
    '             - no parameters

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    Public Sub New( _
            ByVal pName As String, _
            ByVal pAddress As String, _
            ByVal pCity As String, _
            ByVal pState As String, _
            ByVal pZip As Integer
                  )



        MyBase.New()

        _name = pName
        _address = pAddress
        _city = pCity
        _state = pState
        _zip = pZip



    End Sub 'New

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

    Public Sub New(ByVal pCustomer As Customer)

        'Copy constructor.
        'Clones pProduct.

        Me.New( _
            pCustomer.name, _
            pCustomer.address, _
            pCustomer.city, _
            pCustomer.state, _
            pCustomer.zip)

    End Sub 'New(pProduct)

#End Region 'Constructors

#Region "Get/Set Methods"

    
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    'No Get/Set Methods are currently defined.

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public ReadOnly Property name As String
        Get
            Return _name
        End Get
    End Property

    Public ReadOnly Property address As String
        Get
            Return _address
        End Get
    End Property

    Public ReadOnly Property city As String
        Get
            Return _city
        End Get
    End Property

    Public ReadOnly Property state As String
        Get
            Return _state
        End Get
    End Property

    Public ReadOnly Property zip As Integer
        Get
            Return _zip
        End Get
    End Property

    

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _name As String
        Get
            Return mCustomerName
        End Get
        Set(ByVal pValue As String)
            mCustomerName = pValue
        End Set
    End Property

    Private Property _address As String
        Get
            Return mCustomerAddress
        End Get
        Set(ByVal pValue As String)
            mCustomerAddress = pValue
        End Set
    End Property

    Private Property _city As String
        Get
            Return mCustomerCity
        End Get
        Set(ByVal pValue As String)
            mCustomerCity = pValue
        End Set
    End Property

    Private Property _state As String
        Get
            Return mCustomerState
        End Get
        Set(ByVal pValue As String)
            mCustomerState = pValue
        End Set
    End Property

    Private Property _zip As Integer
        Get
            Return mCustomerZip
        End Get
        Set(ByVal pValue As Integer)
            mCustomerZip = pValue
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

    '********** Private Non-Shared Behavioral Methods

    Public Overrides Function ToString() As String

        'ToString() creates and formats a string that holds the
        'information about this class.

        Dim tmpStr As String

        tmpStr = _
            "( Customer: " _
            & "Name=" & _name _
            & ", Address=" & _address _
            & ", City=" & _city _
            & ", State=" & _state _
            & ", Zip=" & _zip.ToString _
            & " )"

        Return tmpStr

    End Function 'ToString()



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

End Class 'Customer


