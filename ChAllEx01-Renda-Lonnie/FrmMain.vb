'Copyright (c) 2009-2013 Dan Turk

#Region "Class / File Comment Header block"
'Program:           ChAllEx01           
'File:              FrmMain.vb     
'Author:            Lonnie Renda            
'Description:       Simulate the creation of a program to add a business, add customers to the business
'                   add inventory items to sell, and the cost to buy those items for the business
'                   a way to sell the items to the customer and record all transactions.
'                   
'Date:              2013-April-12               
'Tier:              User-Interface     
'Exceptions:        None     
'Exception-Handling: for inappropirate input
'Events:            Sale_Made, Inventory_ItemAdded, Customer_Added
'Event-Handling:    Only standard User_Interface event-handling 
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
#End Region 'Option / Imports

Public Class FrmMain

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables

    'Class Variables
    Private WithEvents mTheBusiness As Business
    

    Private mFileNameInput As String
    Private mFileNameOutput As String

  


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

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************


    '********** Public Shared Behavioral Methods
    

    '********** Public Non-Shared Behavioral Methods

    '********** Private Non-Shared Behavioral Methods

    Private Sub _initializeBusinessLogic()

        mTheBusiness = New Business



    End Sub '_initializeBusinessLogic()


    Private Sub _initializeUserInterface()

        'Clear display lables

        lbltotalTrxgrpBusinessSummarytabBusinesstbcMainFrmMain.Text = ""
        lblItemTotalgrpInventorySummarytabBusinesstbcMainFrmMain.Text = ""
        lblTotalQuantitygrpInventorySummarytabBusinesstbcMainFrmMain.Text = ""
        lblTotalCostgrpInventorySummarytabBusinesstbcMainFrmMain.Text = ""
        lblSalesNumbergrpSalesSummarytabBusinesstbcMainFrmMain.Text = ""
        lblSalesCostgrpSalesSummarytabBusinesstbcMainFrmMain.Text = ""
        lblSalesValuegrpSalesSummarytabBusinesstbcMainFrmMain.Text = ""
        lblCustomerNamegrpCustomerInformationtabCustomerstbcMainFrmMain.Text = ""
        lblCustomerAddressgrpCustomerInformationtabCustomerstbcMainFrmMain.Text = ""
        lblCustomerCitygrpCustomerInformationtabCustomerstbcMainFrmMain.Text = ""
        lblCustomerStategrpCustomerInformationtabCustomerstbcMainFrmMain.Text = ""
        lblCustomerZipgrpCustomerInformationtabCustomerstbcMainFrmMain.Text = ""
        lblInventoryTotalItemsNumbergrpInventoryDetailstabInventorytbcMainFrmMain.Text = ""
        lblInventoryCostTotalItemsgrpInventoryDetailstabInventorytbcMainFrmMain.Text = ""
        lblNamegrpCustomerDetailstabSalestbcMainFrmMain.Text = ""
        lblStreetAddressgrpCustomerDetailstabSalestbcMainFrmMain.Text = ""
        lblCitygrpCustomerDetailstabSalestbcMainFrmMain.Text = ""
        lblStategrpCustomerDetailstabSalestbcMainFrmMain.Text = ""
        lblZipgrpCustomerDetailstabSalestbcMainFrmMain.Text = ""
        lblItemIDgrpInventoryDetailstabSalestbcMainFrmMain.Text = ""
        lblItemDescriptiongrpInventoryDetailstabSalestbcMainFrmMain.Text = ""
        lblQuantitygrpInventoryDetailstabSalestbcMainFrmMain.Text = ""
        lblCostgrpInventoryDetailstabSalestbcMainFrmMain.Text = ""
        lblSalegrpInventoryDetailstabSalestbcMainFrmMain.Text = ""

        
        'test data
        mTheBusiness.addBusiness("Tom's Hardware Supply", "Hardware Supply Company", #2/8/2003#)
        mTheBusiness.addCustomer(New Customer("Lonnie TD", "18 N. County Street", "Waukegan", "IL", 60085))
        mTheBusiness.addCustomer(New Customer("Allison TD", "702 Juneway Ave", "Deerfield", "IL", 60015))
        mTheBusiness.addCustomer(New Customer("Hal's TD", "356 Beech Drive", "Wheeling", "CO", 60090))
        mTheBusiness.addInventoryItem(New Inventory("adsd4", "Screws", CDec(0.3), CDec(0.5), 1000))
        mTheBusiness.addInventoryItem(New Inventory("bsd4", "Hammer", CDec(1.3), CDec(3.5), 50))
        mTheBusiness.addInventoryItem(New Inventory("D1234", "Drill", CDec(20.3), CDec(30.5), 13))
        mTheBusiness.addSale(New Sale("Lonnie TD", "adsd4", 400))
        mTheBusiness.addSale(New Sale("Allison TD", "bsd4", 8))
        mTheBusiness.addSale(New Sale("Lonnie TD", "D1234", 12))
        mTheBusiness.addSale(New Sale("Ethan TD", "adsd4", 12))
        mTheBusiness.addSale(New Sale("Lonnie TD", "D1234", 12))
        mTheBusiness.addCustomer(New Customer("Hal's TD", "356 Beech Drive", "Wheeling", "CO", 60090))


        'set default buttons
        Me.CancelButton = btnExitFrmMain
        Me.AcceptButton = btnCreategrpCreateBusinesstabBusinesstbcMainFrmMain

        'disable listboxes:

        'customer tab
        lstQuantityOrderedgrpCustomerTransactionstabCustomerstbcMainFrmMain.Enabled = False
        'inventory tab
        lstInventoryItemgrpInventoryDetailstabInventorytbcMainFrmMain.Enabled = False
        lstInventoryQuantitygrpInventoryDetailstabInventorytbcMainFrmMain.Enabled = False
        lstInventoryCostgrpInventoryDetailstabInventorytbcMainFrmMain.Enabled = False
        lstInventoryPricegrpInventoryDetailstabInventorytbcMainFrmMain.Enabled = False
        lstInventoryCarryCostgrpInventoryDetailstabInventorytbcMainFrmMain.Enabled = False
        'sales tab
        lstInventoryIDgrpSalesHistorytabSalestbcMainFrmMain.Enabled = False
        lstQuantitygrpSalesHistorytabSalestbcMainFrmMain.Enabled = False

    End Sub 'initializeUserInterface()

    Private Sub _BusinessExistsCheck()

        'checks to see if a business already exists for setting up the user interface for future input of the business

        If Not mTheBusiness Is Nothing Then

            btnCreategrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
            txtBusinessNamegrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
            txtBusinessTypegrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
            dtpBusinessCreationDategrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
            lblBusinessNamegrpBusinessSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.BusinessName.ToString

        End If

    End Sub

#End Region 'Behavioral Methods

#Region "Event Procedures"
    '******************************************************************
    'Event Procedures
    '******************************************************************



    '********** User-Interface Event Procedures
    '             - Initiated explicitly by user

    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _initializeBusinessLogic()
        _initializeUserInterface()

    End Sub 'FrmMain_Load

    'Creation buttons
    Private Sub btnCreategrpCreateBusinesstabBusinesstbcMainFrmMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreategrpCreateBusinesstabBusinesstbcMainFrmMain.Click
        _BusinessExistsCheck()
        Dim BusinessName As String
        Dim BusinessType As String
        Dim BusinessCreationDate As Date

        BusinessName = txtBusinessNamegrpCreateBusinesstabBusinesstbcMainFrmMain.Text
        BusinessType = txtBusinessTypegrpCreateBusinesstabBusinesstbcMainFrmMain.Text

        'Check to see if a business already exists.  If so, no others can be created.

        If Not mTheBusiness.BusinessName Is Nothing Then
            MessageBox.Show("Sorry a business already exists.  You can only have one business.")
            btnCreategrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
            txtBusinessNamegrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
            txtBusinessTypegrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
            dtpBusinessCreationDategrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
            lblBusinessNamegrpBusinessSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.BusinessName.ToString


        End If


        If BusinessName = "" Then
            MessageBox.Show("Please enter a name for the business.")
            txtBusinessNamegrpCreateBusinesstabBusinesstbcMainFrmMain.Focus()
            txtBusinessNamegrpCreateBusinesstabBusinesstbcMainFrmMain.SelectAll()
            Exit Sub
        End If

        If BusinessType = "" Then
            MessageBox.Show("Please enter a type for the business.")
            txtBusinessTypegrpCreateBusinesstabBusinesstbcMainFrmMain.Focus()
            txtBusinessTypegrpCreateBusinesstabBusinesstbcMainFrmMain.SelectAll()
            Exit Sub
        End If

        BusinessName = txtBusinessNamegrpCreateBusinesstabBusinesstbcMainFrmMain.Text
        BusinessType = txtBusinessTypegrpCreateBusinesstabBusinesstbcMainFrmMain.Text
        BusinessCreationDate = dtpBusinessCreationDategrpCreateBusinesstabBusinesstbcMainFrmMain.Value


        'Pass to business to update mthebusiness variables
        mTheBusiness.addBusiness(BusinessName,BusinessName,BusinessCreationDate)

        'set tab so another business cannot be entered.






    End Sub

    Private Sub btnAddgrpAddCustomertabCustomertbcMainFrmMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddgrpAddCustomertabCustomertbcMainFrmMain.Click
        'Customer tab variables
        Dim CustomerName As String
        Dim CustomerAddress As String
        Dim CustomerCity As String
        Dim CustomerState As String
        Dim CustomerZip As Integer


        'set variables based on input in text boxes
        CustomerName = txtNamegrpAddCustomertabCustomertbcMainFrmMain.Text
        CustomerAddress = txtStreetAddressgrpAddCustomertabCustomertbcMainFrmMain.Text
        CustomerCity = txtCitygrpAddCustomertabCustomertbcMainFrmMain.Text
        CustomerState = txtStategrpAddCustomertabCustomertbcMainFrmMain.Text

        'Check for input errors
        If CustomerName = "" Then
            MessageBox.Show("Please enter a name for the customer.")
            txtNamegrpAddCustomertabCustomertbcMainFrmMain.Focus()
            txtNamegrpAddCustomertabCustomertbcMainFrmMain.SelectAll()
            Exit Sub
        End If

        If CustomerAddress = "" Then
            MessageBox.Show("Please enter a street address for the customer.")
            txtStreetAddressgrpAddCustomertabCustomertbcMainFrmMain.Focus()
            txtStreetAddressgrpAddCustomertabCustomertbcMainFrmMain.SelectAll()
            Exit Sub
        End If

        If CustomerCity = "" Then
            MessageBox.Show("Please enter a city for the customer address.")
            txtCitygrpAddCustomertabCustomertbcMainFrmMain.Focus()
            txtCitygrpAddCustomertabCustomertbcMainFrmMain.SelectAll()
            Exit Sub
        End If

        If CustomerState = "" Then
            MessageBox.Show("Please enter a state for the customer address.")
            txtStategrpAddCustomertabCustomertbcMainFrmMain.Focus()
            txtStategrpAddCustomertabCustomertbcMainFrmMain.SelectAll()
            Exit Sub
        End If
        Try

            CustomerZip = Integer.Parse(txtZipgrpAddCustomertabCustomertbcMainFrmMain.Text)

        Catch ex As Exception
            MessageBox.Show("Please enter a five-digit zip code.")
            txtZipgrpAddCustomertabCustomertbcMainFrmMain.SelectAll()
            txtZipgrpAddCustomertabCustomertbcMainFrmMain.Focus()
            Exit Sub
            'displays the error, selects all and reset the focus.

        End Try

        'make sure customer does not exist
        Dim i As Integer
        Dim currentCustomer As Customer
        Dim name As String
        name = txtNamegrpAddCustomertabCustomertbcMainFrmMain.Text



        For i = 0 To mTheBusiness.numCustomers - 1

            currentCustomer = mTheBusiness.ithCustomer(i)
            If name = currentCustomer.name Then
                MessageBox.Show("Sorry that customer already exists.  You cannot add them.")
                Exit Sub

            End If
        Next i

        'Create instance of class




        mTheBusiness.addCustomer(New Customer(CustomerName, CustomerAddress, CustomerCity, CustomerState, CustomerZip))
        

        'get ready for next input   
        txtNamegrpAddCustomertabCustomertbcMainFrmMain.Clear()
        txtStreetAddressgrpAddCustomertabCustomertbcMainFrmMain.Clear()
        txtCitygrpAddCustomertabCustomertbcMainFrmMain.Clear()
        txtStategrpAddCustomertabCustomertbcMainFrmMain.Clear()
        txtZipgrpAddCustomertabCustomertbcMainFrmMain.Clear()

        txtNamegrpAddCustomertabCustomertbcMainFrmMain.Focus()


    End Sub 'btnAddgrpAddCustomertabCustomertbcMainFrmMain


    Private Sub btnAddItemgrpAddInventorytabInventorytbcMainFrmMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddItemgrpAddInventorytabInventorytbcMainFrmMain.Click

        'Inventory tab variables
        Dim ItemID As String
        Dim ItemDescription As String
        Dim ItemCost As Decimal
        Dim ItemPrice As Decimal
        Dim ItemQuantity As Integer

        'set variables from textboxes
        ItemID = txtItemIDgrpAddItemtabInventorytbcMainFrmMain.Text
        ItemDescription = txtItemDescriptiongrpAddItemtabInventorytbcMainFrmMain.Text
        ItemCost = CDec(txtItemCostgrpAddItemtabInventorytbcMainFrmMain.Text)
        ItemPrice = CDec(txtItemPricegrpAddItemtabInventorytbcMainFrmMain.Text)
        ItemQuantity = CInt(txtItemQuantitygrpAddItemtabInventorytbcMainFrmMain.Text)

        'errorchecking
        If ItemID = "" Then
            MessageBox.Show("Please enter an ID for inventory item.")
            txtItemIDgrpAddItemtabInventorytbcMainFrmMain.Focus()
            txtItemIDgrpAddItemtabInventorytbcMainFrmMain.SelectAll()
            Exit Sub
        End If
        If ItemDescription = "" Then

            MessageBox.Show("Please enter a name for inventory item.")
            txtItemDescriptiongrpAddItemtabInventorytbcMainFrmMain.Focus()
            txtItemDescriptiongrpAddItemtabInventorytbcMainFrmMain.SelectAll()
            Exit Sub

        End If

        Try
            ItemCost = Decimal.Parse(txtItemCostgrpAddItemtabInventorytbcMainFrmMain.Text)
        Catch ex As Exception
            MessageBox.Show( _
                "Please enter a proper cost amount (eg 18.75).  No '$' needed.")

            txtItemCostgrpAddItemtabInventorytbcMainFrmMain.Focus()
            txtItemCostgrpAddItemtabInventorytbcMainFrmMain.SelectAll()
            Exit Sub


        End Try

        Try
            ItemPrice = Decimal.Parse(txtItemPricegrpAddItemtabInventorytbcMainFrmMain.Text)
        Catch ex As Exception
            MessageBox.Show( _
                "Please enter a proper price amount (eg 18.75).  No '$' needed.")

            txtItemPricegrpAddItemtabInventorytbcMainFrmMain.Focus()
            txtItemPricegrpAddItemtabInventorytbcMainFrmMain.SelectAll()
            Exit Sub


        End Try

        Try

            ItemQuantity = Integer.Parse(txtItemQuantitygrpAddItemtabInventorytbcMainFrmMain.Text)

        Catch ex As Exception
            MessageBox.Show("Please enter a whole number (eg. 132).")
            txtItemQuantitygrpAddItemtabInventorytbcMainFrmMain.SelectAll()
            txtItemQuantitygrpAddItemtabInventorytbcMainFrmMain.Focus()
            Exit Sub
            'displays the error, selects all and reset the focus.

        End Try

        'check to see if item exists

        Dim i As Integer
        Dim currentInventoryItem As Inventory
        Dim ID As String
        ID = txtItemIDgrpAddItemtabInventorytbcMainFrmMain.Text



        For i = 0 To mTheBusiness.numInventory - 1

            currentInventoryItem = mTheBusiness.ithInventoryItem(i)
            If ID = currentInventoryItem.itemID Then
                MessageBox.Show("Sorry that item already exists.  You cannot add it.")
                Exit Sub

            End If
        Next i

        'Create instance of class

        mTheBusiness.addInventoryItem(New Inventory(ItemID, ItemDescription, ItemCost, ItemPrice, ItemQuantity))

        'get ready for next input
        txtItemIDgrpAddItemtabInventorytbcMainFrmMain.Clear()
        txtItemDescriptiongrpAddItemtabInventorytbcMainFrmMain.Clear()
        txtItemCostgrpAddItemtabInventorytbcMainFrmMain.Clear()
        txtItemPricegrpAddItemtabInventorytbcMainFrmMain.Clear()
        txtItemQuantitygrpAddItemtabInventorytbcMainFrmMain.Clear()

        txtItemIDgrpAddItemtabInventorytbcMainFrmMain.Focus()

    End Sub 'btnAddItemgrpAddInventorytabInventorytbcMainFrmMain

    Private Sub btnConfirmSalegrpNewSaletabSalestbcMainFrmMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConfirmSalegrpNewSaletabSalestbcMainFrmMain.Click

        'local variables for input
        Dim maxQuantity As Integer
        Dim Quantity As Integer
        Dim CustomerName As String
        Dim ItemID As String

        'set variables

        'check to makes sure a quantity was entered.
        Quantity = CInt(nudSaleQuantityNewSaletabSalestbcMainFrmMain.Value)
        maxQuantity = CInt(nudSaleQuantityNewSaletabSalestbcMainFrmMain.Maximum)

        If Quantity <= 0 Then
            MessageBox.Show("Please enter an amount greater than zero.")
            nudSaleQuantityNewSaletabSalestbcMainFrmMain.Focus()
            Exit Sub
        End If
        If Quantity > maxQuantity Then
            MessageBox.Show("Please enter an amount not greater than the quantity in inventory.")
            nudSaleQuantityNewSaletabSalestbcMainFrmMain.Focus()
        End If

        'check to makes sure they choose a customer
        If lstCustomergrpNewSaletabSalestbcMainFrmMain.SelectedIndex < 0 Then
            MessageBox.Show("Please choose a cusomter.")
            lstCustomergrpNewSaletabSalestbcMainFrmMain.Focus()
            Exit Sub
        End If

        'check to see that they choose and inventory item
        If lstInventorygrpNewSaletabSalestbcMainFrmMain.SelectedIndex < 0 Then
            MessageBox.Show("Please choose an item.")
            lstInventorygrpNewSaletabSalestbcMainFrmMain.Focus()
            Exit Sub
        End If

        'move in text from listboxes to variables
        CustomerName = lstCustomergrpNewSaletabSalestbcMainFrmMain.Text
        ItemID = lstInventorygrpNewSaletabSalestbcMainFrmMain.Text

        'Create instance of class

        mTheBusiness.addSale(New Sale(CustomerName, ItemID, Quantity))

        'get ready for next input

        lstInventorygrpNewSaletabSalestbcMainFrmMain.SelectedIndex = -1
        lstCustomergrpNewSaletabSalestbcMainFrmMain.SelectedIndex = -1
        nudSaleQuantityNewSaletabSalestbcMainFrmMain.Value = 1
        lstCustomergrpNewSaletabSalestbcMainFrmMain.Focus()

        lblNamegrpCustomerDetailstabSalestbcMainFrmMain.Text = ""
        lblStreetAddressgrpCustomerDetailstabSalestbcMainFrmMain.Text = ""
        lblCitygrpCustomerDetailstabSalestbcMainFrmMain.Text = ""
        lblStategrpCustomerDetailstabSalestbcMainFrmMain.Text = ""
        lblZipgrpCustomerDetailstabSalestbcMainFrmMain.Text = ""
        lblItemIDgrpInventoryDetailstabSalestbcMainFrmMain.Text = ""
        lblItemDescriptiongrpInventoryDetailstabSalestbcMainFrmMain.Text = ""
        lblQuantitygrpInventoryDetailstabSalestbcMainFrmMain.Text = ""
        lblCostgrpInventoryDetailstabSalestbcMainFrmMain.Text = ""
        lblSalegrpInventoryDetailstabSalestbcMainFrmMain.Text = ""


    End Sub ' btnConfirmSalegrpNewSaletabSalestbcMainFrmMain

    Private Sub _btnReadfromFile_Click( _
            ByVal sender As System.Object, _
            ByVal e As System.EventArgs) _
        Handles _
            btnReadgrpReadWritetabTransactiontbcMainFrmMain.Click


        mFileNameInput = txtlblFilegrpReadWritetabTransactiontbcMainFrmMain.Text
        Try
            mTheBusiness.readFromFile(mFileNameInput)
        Catch ex As Exception
            MessageBox.Show("Error Reading from file: " & ex.ToString)
        
        End Try

    End Sub '_btnReadFromFile_Click(sender,e)

    Private Sub _btnWriteToFile_Click( _
            ByVal sender As System.Object, _
            ByVal e As System.EventArgs) _
        Handles _
            btnWritegrpReadWritetabTransactiontbcMainFrmMain.Click


        mFileNameOutput = txtlblFilegrpReadWritetabTransactiontbcMainFrmMain.Text
        Try
            mTheBusiness.writeToFile(mFileNameOutput)
        Catch ex As Exception
            MessageBox.Show("Error writing to file: " & ex.ToString)
        End Try

    End Sub '_btnReadFromFile_Click(sender,e)
    'Exit Buttons
    Private Sub _btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExitFrmMain.Click


        'btnExit_Click() is the event procedure
        'that fires when the user activates the
        'btnExit Button. It closes the form.


        Me.Close()

    End Sub 'btnExit_Click

    'Customer tab

    Private Sub lstCustomergrpCustomerSummarytabCustomertbcMainFrmMain_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstCustomergrpCustomerSummarytabCustomertbcMainFrmMain.SelectedIndexChanged


        'handles updates labels and list boxes on customer tab based on index selection.

        Dim selectedIndex = lstCustomergrpCustomerSummarytabCustomertbcMainFrmMain.SelectedIndex
        Dim selectedName As String
        Dim selectedCustomer As Customer


        selectedName = lstCustomergrpCustomerSummarytabCustomertbcMainFrmMain.SelectedItem.ToString

        'Updating customer information based on index selection
        If selectedIndex >= 0 Then

            selectedCustomer = mTheBusiness.findCustomerByName(selectedName)
            If selectedCustomer Is Nothing Then
                lblCustomerNamegrpCustomerInformationtabCustomerstbcMainFrmMain.Text = _
                    "Selected customer not found."
            Else '_selectedCustomer found
                lblCustomerNamegrpCustomerInformationtabCustomerstbcMainFrmMain.Text = _
                    selectedCustomer.name.ToString
                lblCustomerAddressgrpCustomerInformationtabCustomerstbcMainFrmMain.Text = _
                    selectedCustomer.address.ToString
                lblCustomerCitygrpCustomerInformationtabCustomerstbcMainFrmMain.Text = _
                    selectedCustomer.city.ToString
                lblCustomerStategrpCustomerInformationtabCustomerstbcMainFrmMain.Text = _
                    selectedCustomer.state.ToString
                lblCustomerZipgrpCustomerInformationtabCustomerstbcMainFrmMain.Text = _
                    selectedCustomer.zip.ToString
            End If '_selectedcustomer found
        Else 'selectedIndex < 0
            lblCustomerNamegrpCustomerInformationtabCustomerstbcMainFrmMain.Text = _
                "No customer selected."
        End If 'selectedIndex < 0

        'updating transaction information based on index selection


        Dim i As Integer
        lstItemNumbergrpCustomerTransactionstabCustomerstbcMainFrmMain.Items.Clear()
        lstQuantityOrderedgrpCustomerTransactionstabCustomerstbcMainFrmMain.Items.Clear()
        Dim currentSaleItem As Sale
        Dim name As String
        name = lstCustomergrpCustomerSummarytabCustomertbcMainFrmMain.Text
        lstItemNumbergrpCustomerTransactionstabCustomerstbcMainFrmMain.Items.Clear()
        lstQuantityOrderedgrpCustomerTransactionstabCustomerstbcMainFrmMain.Items.Clear()


        For i = 0 To mTheBusiness.numSales - 1

            currentSaleItem = mTheBusiness.ithSale(i)
            If name = currentSaleItem.CustomerName Then
                lstItemNumbergrpCustomerTransactionstabCustomerstbcMainFrmMain.Items.Add(currentSaleItem.itemID).ToString()
                lstQuantityOrderedgrpCustomerTransactionstabCustomerstbcMainFrmMain.Items.Add(currentSaleItem.Quantity).ToString()

            End If
        Next i




    End Sub 'lstCustomergrpCustomerSummarytabCustomertbcMainFrmMain_SelectedIndexChanged

    'Inventory tab
    Private Sub _lstInventoryIDgrpInventoryDetailstabInventorytbcMainFrmMain_SelectedIndexChanged( _
            ByVal sender As System.Object, _
            ByVal e As System.EventArgs) _
        Handles lstInventoryIDgrpInventoryDetailstabInventorytbcMainFrmMain.SelectedIndexChanged

        'updates highlight when index changes

        Dim selectedIndex As Integer =
           lstInventoryIDgrpInventoryDetailstabInventorytbcMainFrmMain.SelectedIndex

        lstInventoryItemgrpInventoryDetailstabInventorytbcMainFrmMain.SelectedIndex = selectedIndex
        lstInventoryCostgrpInventoryDetailstabInventorytbcMainFrmMain.SelectedIndex = selectedIndex
        lstInventoryPricegrpInventoryDetailstabInventorytbcMainFrmMain.SelectedIndex = selectedIndex
        lstInventoryQuantitygrpInventoryDetailstabInventorytbcMainFrmMain.SelectedIndex = selectedIndex
        lstInventoryCarryCostgrpInventoryDetailstabInventorytbcMainFrmMain.SelectedIndex = selectedIndex



    End Sub '_lstcusomtergrpcustomersummarytabcustomertbcmainfrmmain_SelectedIndexChanged(sender,e)
    'Sales tab
    Private Sub _lstInventorygrpNewSaletabSalestbcMainFrmMain_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstInventorygrpNewSaletabSalestbcMainFrmMain.SelectedIndexChanged


        'changes highlights based on index change

        Dim selectedIndex = lstInventorygrpNewSaletabSalestbcMainFrmMain.SelectedIndex
        Dim selectedID As String
        Dim selectedInventory As Inventory

        If selectedIndex >= 0 Then
            selectedID = lstInventorygrpNewSaletabSalestbcMainFrmMain.SelectedItem.ToString
            selectedInventory = mTheBusiness.findInventoryItemByID(selectedID)
            If selectedInventory Is Nothing Then
                lblItemIDgrpInventoryDetailstabSalestbcMainFrmMain.Text = _
                    "Selected Item not found."
            Else '_selectedcustomer found
                lblItemIDgrpInventoryDetailstabSalestbcMainFrmMain.Text = _
                    selectedInventory.itemID.ToString
                lblItemDescriptiongrpInventoryDetailstabSalestbcMainFrmMain.Text = _
                    selectedInventory.itemDescription.ToString
                lblQuantitygrpInventoryDetailstabSalestbcMainFrmMain.Text = _
                    selectedInventory.itemQuantity.ToString
                lblCostgrpInventoryDetailstabSalestbcMainFrmMain.Text = _
                    selectedInventory.itemCost.ToString("c")
                lblSalegrpInventoryDetailstabSalestbcMainFrmMain.Text = _
                    selectedInventory.itemPrice.ToString("c")


                nudSaleQuantityNewSaletabSalestbcMainFrmMain.Value = 1
                nudSaleQuantityNewSaletabSalestbcMainFrmMain.Minimum = 1
                nudSaleQuantityNewSaletabSalestbcMainFrmMain.Maximum = selectedInventory.itemQuantity


            End If '_selectedcustomer found
        Else '_selectedIndex < 0
            lblCustomerNamegrpCustomerInformationtabCustomerstbcMainFrmMain.Text = _
                "No Item selected."
        End If '_selectedIndex < 0

    End Sub 'lstInventorygrpNewSaletabSalestbcMainFrmMain_SelectedIndexChanged

    Private Sub _lstCustomergrpNewSaletabSalestbcMainFrmMain_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstCustomergrpNewSaletabSalestbcMainFrmMain.SelectedIndexChanged


        'adds customer info based on index change

        Dim _selectedIndex = lstCustomergrpNewSaletabSalestbcMainFrmMain.SelectedIndex
        Dim _selectedName As String
        Dim _selectedCustomer As Customer

        If _selectedIndex >= 0 Then
            _selectedName = lstCustomergrpNewSaletabSalestbcMainFrmMain.SelectedItem.ToString
            _selectedCustomer = mTheBusiness.findCustomerByName(_selectedName)
            If _selectedCustomer Is Nothing Then
                lblNamegrpCustomerDetailstabSalestbcMainFrmMain.Text = _
                    "Selected customer not found."
            Else '_selectedProduct found
                lblNamegrpCustomerDetailstabSalestbcMainFrmMain.Text = _
                    _selectedCustomer.name.ToString
                lblStreetAddressgrpCustomerDetailstabSalestbcMainFrmMain.Text = _
                    _selectedCustomer.address.ToString
                lblCitygrpCustomerDetailstabSalestbcMainFrmMain.Text = _
                    _selectedCustomer.city.ToString
                lblStategrpCustomerDetailstabSalestbcMainFrmMain.Text = _
                    _selectedCustomer.state.ToString
                lblZipgrpCustomerDetailstabSalestbcMainFrmMain.Text = _
                    _selectedCustomer.zip.ToString
            End If '_selectedProduct found
        Else '_selectedIndex < 0
            lblCustomerNamegrpCustomerInformationtabCustomerstbcMainFrmMain.Text = _
                "No customer selected."
        End If '_selectedIndex < 0

    End Sub 'lstCustomergrpNewSaletabSalestbcMainFrmMain_SelectedIndexChanged

    Private Sub _lstCustomergrpSalesHistorytabSalestbcMainFrmMain_SelectedIndexChanged( _
           ByVal sender As System.Object, _
           ByVal e As System.EventArgs) _
       Handles lstCustomergrpSalesHistorytabSalestbcMainFrmMain.SelectedIndexChanged

        'changes highlights based on index change

        Dim selectedIndex As Integer =
           lstCustomergrpSalesHistorytabSalestbcMainFrmMain.SelectedIndex

        lstInventoryIDgrpSalesHistorytabSalestbcMainFrmMain.SelectedIndex = selectedIndex
        lstQuantitygrpSalesHistorytabSalestbcMainFrmMain.SelectedIndex = selectedIndex



    End Sub '_lstCustomergrpSalesHistorytabSalestbcMainFrmMain_SelectedIndexChanged(sender,e)
    '********** User-Interface Event Procedures
    '             - Initiated automatically by system



    '********** Business Logic Event Procedures
    '             - Initiated as a result of business logic
    '               method(s) running

    Private Sub _Business_CustomerAdded( _
            ByVal sender As Object, _
            ByVal e As EventArgs) _
        Handles _
            mTheBusiness.Business_CustomerAdded



        Dim theCustomerEventArgs As ClsBusiness_CustomerAdded_EventArgs = _
            CType(e, ClsBusiness_CustomerAdded_EventArgs)



        lblCustomerTotalsgrpBusinessSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.numCustomers.ToString

        lstCustomergrpCustomerSummarytabCustomertbcMainFrmMain.Items.Add( _
           theCustomerEventArgs.customer.name.ToString)


        lstCustomergrpNewSaletabSalestbcMainFrmMain.Items.Add( _
           theCustomerEventArgs.customer.name.ToString)

        'update label
        lbltotalTrxgrpBusinessSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.numTransactions.ToString

        'write transaction to screen
        txtTrxTabTransactionsTbcMain.Text &= "Business_CustomerAdded: " _
            & "sender=" & sender.ToString _
            & ", e=" & e.ToString _
            & vbCrLf & vbCrLf

    End Sub '_Business_CustomerAdded(sender,e)



    Private Sub _Business_InventoryItemAdded( _
               ByVal sender As Object, _
               ByVal e As EventArgs) _
           Handles _
               mTheBusiness.Business_InventoryItemAdded

        Dim theInventoryItemEventArgs As Business_InventoryItemAdded_EventArgs = _
            CType(e, Business_InventoryItemAdded_EventArgs)

        'Add amount of inventory items to busines tab
        lblItemTotalgrpInventorySummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.numInventory.ToString

        lblTotalCostgrpInventorySummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.totalInventoryCost.ToString("C")

        lblTotalQuantitygrpInventorySummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.totalInventoryQuantity.ToString

        'Add amount of inventory items to inventory tab
        lblInventoryTotalItemsNumbergrpInventoryDetailstabInventorytbcMainFrmMain.Text = mTheBusiness.numInventory.ToString

        lblInventoryCostTotalItemsgrpInventoryDetailstabInventorytbcMainFrmMain.Text = mTheBusiness.totalInventoryCost.ToString("C")

        lblInventoryTotalItemsNumbergrpInventoryDetailstabInventorytbcMainFrmMain.Text = mTheBusiness.totalInventoryQuantity.ToString

        'Add info to Inventory tab

        lstInventoryIDgrpInventoryDetailstabInventorytbcMainFrmMain.Items.Add( _
           theInventoryItemEventArgs.inventory.itemID.ToString)

        lstInventoryItemgrpInventoryDetailstabInventorytbcMainFrmMain.Items.Add( _
           theInventoryItemEventArgs.inventory.itemDescription.ToString)

        lstInventoryCostgrpInventoryDetailstabInventorytbcMainFrmMain.Items.Add( _
           theInventoryItemEventArgs.inventory.itemCost.ToString("C"))

        lstInventoryPricegrpInventoryDetailstabInventorytbcMainFrmMain.Items.Add( _
           theInventoryItemEventArgs.inventory.itemPrice.ToString("C"))

        lstInventoryQuantitygrpInventoryDetailstabInventorytbcMainFrmMain.Items.Add( _
           theInventoryItemEventArgs.inventory.itemQuantity.ToString)

        Dim carrycost As Decimal
        carrycost = theInventoryItemEventArgs.inventory.itemQuantity * theInventoryItemEventArgs.inventory.itemCost

        lstInventoryCarryCostgrpInventoryDetailstabInventorytbcMainFrmMain.Items.Add(carrycost.ToString("C"))


        'lstInventoryCarryCostgrpInventoryDetailstabInventorytbcMainFrmMain.Items.Add = ( _
        '   theInventoryItemEventArgs.inventory.itemCost * theInventoryItemEventArgs.inventory.itemQuantity).ToString


        'Add info to sales tab
        lstInventorygrpNewSaletabSalestbcMainFrmMain.Items.Add( _
            theInventoryItemEventArgs.inventory.itemID.ToString)

        'update label
        lbltotalTrxgrpBusinessSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.numTransactions.ToString

        'write to transactions screen
        txtTrxTabTransactionsTbcMain.Text &= "Business_InventoryItemAdded: " _
            & "sender=" & sender.ToString _
            & ", e=" & e.ToString _
            & vbCrLf & vbCrLf

    End Sub '_Business_CustomerAdded(sender,e)

    Private Sub _Business_SaleMade( _
              ByVal sender As Object, _
              ByVal e As EventArgs) _
          Handles _
              mTheBusiness.Business_SaleMade

        Dim theSaleMadeEventArgs As Business_SaleMade_EventArgs = _
            CType(e, Business_SaleMade_EventArgs)

        'Updates all fields after a sales is made on all tabs

        'update Sales tab
        lstCustomergrpSalesHistorytabSalestbcMainFrmMain.Items.Add( _
        theSaleMadeEventArgs.sale.CustomerName.ToString)

        lstInventoryIDgrpSalesHistorytabSalestbcMainFrmMain.Items.Add( _
        theSaleMadeEventArgs.sale.itemID.ToString)

        lstQuantitygrpSalesHistorytabSalestbcMainFrmMain.Items.Add( _
            theSaleMadeEventArgs.sale.Quantity.ToString)

        lblTotalSalesTransactionsNumbergrpSalesHistorytabSalestbcMainFrmMain.Text = mTheBusiness.numSales.ToString
        lblSalesValueAmountgrpSalesHistorytabSalestbcMainFrmMain.Text = mTheBusiness.TotalSalesAmount.ToString("c")

        'update Business tab--sales

        lblSalesNumbergrpSalesSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.numSales.ToString
        lblSalesCostgrpSalesSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.CostofSales.ToString("c")
        lblSalesValuegrpSalesSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.TotalSalesAmount.ToString("c")

        'update Business tab-- inventory

        lblTotalQuantitygrpInventorySummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.totalInventoryQuantity.ToString
        lblTotalCostgrpInventorySummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.totalInventoryCost.ToString("C")

        'update inventory tab--totals

        lblInventoryTotalItemsNumbergrpInventoryDetailstabInventorytbcMainFrmMain.Text = mTheBusiness.totalInventoryQuantity.ToString
        lblInventoryCostTotalItemsgrpInventoryDetailstabInventorytbcMainFrmMain.Text = mTheBusiness.totalInventoryCost.ToString("C")

        'update inventory tab-inventory details
        Dim i As Integer
        lstInventoryQuantitygrpInventoryDetailstabInventorytbcMainFrmMain.Items.Clear()
        lstInventoryCarryCostgrpInventoryDetailstabInventorytbcMainFrmMain.Items.Clear()
        Dim currentinventoryitem As Inventory
        For i = 0 To mTheBusiness.numInventory - 1

            currentinventoryitem = mTheBusiness.ithInventoryItem(i)
            lstInventoryQuantitygrpInventoryDetailstabInventorytbcMainFrmMain.Items.Add(currentinventoryitem.itemQuantity.ToString)
            lstInventoryCarryCostgrpInventoryDetailstabInventorytbcMainFrmMain.Items.Add((currentinventoryitem.itemQuantity * currentinventoryitem.itemCost).ToString("C"))
        Next i

        'update label
        lbltotalTrxgrpBusinessSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.numTransactions.ToString

        'write to transactions screen
        txtTrxTabTransactionsTbcMain.Text &= "Business_SaleMade: " _
            & "sender=" & sender.ToString _
            & ", e=" & e.ToString _
            & vbCrLf & vbCrLf
    End Sub '_Business_SaleMade(sender,e)

    Private Sub _Business_Added( _
            ByVal sender As Object, _
            ByVal e As EventArgs) _
        Handles _
            mTheBusiness.Business_Added



        Dim theBusinessEventArgs As Business_Added_EventArgs = _
            CType(e, Business_Added_EventArgs)


        'create instance of the class

        'upate info and stop new business input
        btnCreategrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
        txtBusinessNamegrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
        txtBusinessTypegrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
        dtpBusinessCreationDategrpCreateBusinesstabBusinesstbcMainFrmMain.Enabled = False
        lblBusinessNamegrpBusinessSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.BusinessName.ToString

        'update label
        lbltotalTrxgrpBusinessSummarytabBusinesstbcMainFrmMain.Text = mTheBusiness.numTransactions.ToString

        'write transaction to screen
        txtTrxTabTransactionsTbcMain.Text &= "Business_Added: " _
            & "sender=" & sender.ToString _
            & ", e=" & e.ToString _
            & vbCrLf & vbCrLf

    End Sub '_Business_Added(sender,e)

#End Region 'Event Procedures

#Region "Events"
    '******************************************************************
    'Events
    '******************************************************************

    'No Events are currently defined.

#End Region 'Events






End Class 'FrmMain


