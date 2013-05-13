'Copyright (c) 2009-2013 Dan Turk

#Region "Class / File Comment Header block"
'Program:                      
'File:              Business_SaleMade_EventArgs.vb     
'Author:            Lonnie Renda            
'Description:       allows for the sending of the instance of the class to the frm main
'                   
'                   
'                   .
'Date:              2013-April-12               
'Tier:              User-Interface     
'Exceptions:        None     
'Exception-Handling:None 
'Events:            None 
'Event-Handling:    Only standard User_Interface event-handling 
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
#End Region 'Option / Imports

Public Class Business_SaleMade_EventArgs
    Inherits System.EventArgs

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables

    Private mSale As Sale

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
           ByVal pSale As Sale _
           )

        'Special constructor - Create object from pSale attribute.

        MyBase.New()

        _sale = pSale

    End Sub 'New(pSale)

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    'No Get/Set Methods are currently defined.

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public ReadOnly Property sale As Sale
        Get
            Return _sale
        End Get
    End Property

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _sale As Sale
        Get
            Return mSale
        End Get
        Set(ByVal pValue As Sale)
            mSale = pValue
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
            "( SaleMade_ADDED EVENT_ARGS: " _
            & "Business=" & _sale.ToString _
            & " )"

        Return tmpStr

    End Function '_toString()



    'No user interface defined.



#End Region 'Behavioral Methods

#Region "Event Procedures"
    '******************************************************************
    'Event Procedures
    '******************************************************************



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



End Class 'Business_SaleMade_EventArgs
