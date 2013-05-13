'Copyright (c) 2009-2013 Dan Turk

#Region "Class / File Comment Header block"
'Program:               ChAllEx01  
'File:                  Business_InventroyItemAdded_EventArgs
'Author:                Lonnie Renda
'Description:           Allows for the passing of this particular instance to the frmmain after a item was added.
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

Public Class Business_InventoryItemAdded_EventArgs
    Inherits System.EventArgs

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    '********** Module-level variables

    Private mInventory As Inventory


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
            ByVal pInventory As Inventory _
            )

        'Special constructor - Create object from pInventory attribute.

        MyBase.New()

        _inventory = pInventory

    End Sub 'New(pInventory)

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public ReadOnly Property inventory As Inventory
        Get
            Return _inventory
        End Get
    End Property

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _inventory As Inventory
        Get
            Return mInventory
        End Get
        Set(ByVal pValue As Inventory)
            mInventory = pValue
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
            "( Inventory_ADDED EVENT_ARGS: " _
            & "Business=" & _inventory.ToString _
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

End Class 'Business_BusinessAdded_EventArgs