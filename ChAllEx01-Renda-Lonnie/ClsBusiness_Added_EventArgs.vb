Public Class Business_Added_EventArgs
    Inherits System.EventArgs

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables
    Private mBusinessName As String
    Private mBusinessType As String
    Private mBusinessDate As Date



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
            ByVal pBusinessName As String, _
            ByVal pBusinessType As String, _
            ByVal pBusinessDate As Date)

        'creates the new the EventArgs object with the Automobile rented and the distance driven.

        MyBase.New()

        'call the private get/set methods to set the values.

        _businessname = pBusinessName
        _businesstype = pBusinessType
        _businessdate = pBusinessDate

    End Sub 'New()

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

    Public ReadOnly Property businessname As String
        Get
            Return _businessname
        End Get
    End Property

    Public ReadOnly Property businesstype As String
        Get
            Return _businesstype
        End Get
    End Property

    Public ReadOnly Property businessdate As Date
        Get
            Return _businessdate
        End Get
    End Property
    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _businessname As String
        Get
            Return mBusinessName
        End Get
        Set(ByVal pValue As String)
            mBusinessName = pValue
        End Set
    End Property

    Private Property _businesstype As String
        Get
            Return mBusinessType
        End Get
        Set(ByVal pValue As String)
            mBusinessType = pValue
        End Set
    End Property

    Private Property _businessdate As Date
        Get
            Return mBusinessDate
        End Get
        Set(ByVal pValue As Date)
            mBusinessDate = pValue
        End Set
    End Property
#End Region 'Get/Set Methods

End Class
