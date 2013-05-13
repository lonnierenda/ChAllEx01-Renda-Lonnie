﻿'Copyright (c) 2009-2013 Dan Turk

#Region "Class / File Comment Header block"
'Program:            ChAllEx01
'File:               ClsBusiness_CouldNotWriteFile_Exception.vb
'Author:             Lonnie Renda
'Description:        Throws a customer exception when file name is not correct.
'Date:               2013 May 10
'                      - Created.
'Tier:               Error
'Exceptions:         Handles writing of file error based on name
'Exception-Handling: none
'Events:             none
'Event-Handling:     Handles writing of a file button click when there are errors.
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
#End Region 'Option / Imports

Public Class Business_CouldNotWriteFile_Exception
    Inherits System.Exception

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    '********** Module-level variables

    Private mFileName As String

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
            ByVal pFileName As String _
            )

        'Special constructor - Create object from pStudent attribute.

        MyBase.New()

        _filename = pFileName

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

    Public ReadOnly Property filename As String
        Get
            Return _filename
        End Get
    End Property

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _filename As String
        Get
            Return mFileName
        End Get
        Set(ByVal pValue As String)
            mFileName = pValue
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
            "( Business COULD_NOT_WRITE_FILE EXCEPTION: " _
            & "Filename=" & _filename _
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
