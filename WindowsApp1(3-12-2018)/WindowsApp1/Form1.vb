Imports System.Data.SqlClient

Public Class Form1
    Dim dateAndTime As Date
    Dim BeratMasuk As Double
    Dim BeratKeluar As Double
    Dim Code As String
    Dim TotalBerat As Double
    Dim LorryNo As String
    Dim Product As String
    Dim Bags As Integer
    Dim Basah As Double
    Dim Kotor As Double
    Dim Rosak As Double
    Dim Lain As Double
    Dim Muda As Double
    Dim Price As Double
    Dim Amount As Double
    Dim NetWeight As Double
    Dim TotalDisc As Double
    Dim WaktuMasuk As String
    Dim WaktuKeluar As String
    Dim CustomerFound As Boolean
    Dim ComServer



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtLorryNo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtLorryNo.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If txtLorryNo.Text IsNot ("") Then
                LorryNo = txtLorryNo.Text
                txtCode.Select()
            Else
                MessageBox.Show("Please enter the lorry number to proceed.", "Invalid Lorry No", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtCode_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCode.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If txtCode.Text IsNot ("") Then
                Code = txtCode.Text
                GetCustomerData()
            Else
                MessageBox.Show("Supplier not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtMasukKg_KeyDown(sender As Object, e As KeyEventArgs) Handles txtMasukKg.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtMasukKg.Text) Then
                lblTimeIn.Text = TimeString
                WaktuMasuk = lblTimeIn.Text
                BeratMasuk = txtMasukKg.Text
                txtKeluarKg.Select()
            Else
                MessageBox.Show("Please enter a valid KG!", "Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtMasukKg.Text = ""
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtKeluarKg_KeyDown(sender As Object, e As KeyEventArgs) Handles txtKeluarKg.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If txtKeluarKg.Text.Trim.Length > 0 Then
                If IsNumeric(txtKeluarKg.Text) Then
                    BeratKeluar = txtKeluarKg.Text
                    If BeratMasuk >= BeratKeluar Then
                        lblTimeOut.Text = TimeString
                        WaktuKeluar = lblTimeOut.Text
                        TotalBerat = BeratMasuk - BeratKeluar
                        txtTotal.Text = TotalBerat
                        txtGrossWT.Text = TotalBerat
                        txtProduct.Select()
                    Else
                        MessageBox.Show("Please enter a valid KG!", "Invalid KG", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                        txtKeluarKg.Text = ""
                        txtGrossWT.Text = ""
                        txtTotal.Text = ""
                    End If
                Else
                    MessageBox.Show("Please enter a valid KG!", "Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    txtKeluarKg.Text = ""
                    txtGrossWT.Text = ""
                    txtTotal.Text = ""
                End If
            Else
                MessageBox.Show("Please enter Keluar KG!", "Blank", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtKeluarKg.Text = ""
                txtGrossWT.Text = ""
                txtTotal.Text = ""
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtProduct_KeyDown(sender As Object, e As KeyEventArgs) Handles txtProduct.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If txtProduct.Text IsNot ("") Then
                Product = txtProduct.Text
                GetInfo()
                txtBags.Select()
            Else
                MessageBox.Show("Product code is empty! Please enter a valid product code.", "Invalid Product", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtProduct.Select()
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtBags_KeyDown(sender As Object, e As KeyEventArgs) Handles txtBags.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtBags.Text) Then
                Bags = txtBags.Text
                txtBasah.Select()
            ElseIf txtBags.Text = ("") Then
                txtBags.Text = "0"
                Bags = "0"
                txtBasah.Select()
            Else
                MessageBox.Show("Invalid Number! Please try again.", "Invalid number!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtBags.Text = ""
                txtBags.Select()
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtBasah_KeyDown(sender As Object, e As KeyEventArgs) Handles txtBasah.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtBasah.Text) Then
                Basah = txtBasah.Text
                txtKotor.Select()
            ElseIf txtBasah.Text = ("") Then
                txtBasah.Text = "0"
                Basah = "0"
                txtKotor.Select()
            Else
                MessageBox.Show("Invalid percentage! Please try again.", "Invalid number!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtBasah.Text = ""
                txtBasah.Select()
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtKotor_KeyDown(sender As Object, e As KeyEventArgs) Handles txtKotor.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtKotor.Text) Then
                Kotor = txtKotor.Text
                txtMuda.Select()
            ElseIf txtKotor.Text = ("") Then
                txtKotor.Text = "0"
                Kotor = "0"
                txtMuda.Select()
            Else
                MessageBox.Show("Invalid percentage! Please try again.", "Invalid number!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtKotor.Text = ""
                txtKotor.Select()
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtMuda_KeyDown(sender As Object, e As KeyEventArgs) Handles txtMuda.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtMuda.Text) Then
                Muda = txtMuda.Text
                txtRosak.Select()
            ElseIf txtMuda.Text = ("") Then
                txtMuda.Text = "0"
                Muda = "0"
                txtRosak.Select()
            Else
                MessageBox.Show("Invalid percentage! Please try again.", "Invalid number!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtMuda.Text = ""
                txtMuda.Select()
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtRosak_KeyDown(sender As Object, e As KeyEventArgs) Handles txtRosak.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtRosak.Text) Then
                Rosak = txtRosak.Text
                txtLain.Select()
            ElseIf txtRosak.Text = ("") Then
                txtRosak.Text = "0"
                Rosak = "0"
                txtLain.Select()
            Else
                MessageBox.Show("Invalid percentage! Please try again.", "Invalid number!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtRosak.Text = ""
                txtRosak.Select()
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtLain_KeyDown(sender As Object, e As KeyEventArgs) Handles txtLain.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtLain.Text) Then
                Lain = txtLain.Text
                TotalDisc = (Basah + Kotor + Muda + Rosak + Lain)
                txtPotongan.Text = TotalDisc
                txtPrice.Select()
            ElseIf txtLain.Text = ("") Then
                txtLain.Text = "0"
                Lain = "0"
                TotalDisc = (Basah + Kotor + Muda + Rosak + Lain)
                txtPotongan.Text = TotalDisc
                txtPrice.Select()
            Else
                MessageBox.Show("Invalid percentage! Please try again.", "Invalid number!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtLain.Text = ""
                txtLain.Select()
            End If
        End If
    End Sub
#Disable Warning IDE1006 ' Naming Styles
    Private Sub txtPrice_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPrice.KeyDown
#Enable Warning IDE1006 ' Naming Styles
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtPrice.Text) Then
                Price = txtPrice.Text
                If TotalDisc < 100 Then
                    Amount = ((((100 - TotalDisc) / 100) * TotalBerat) * Price)
                    NetWeight = ((100 - TotalDisc) / 100) * TotalBerat
                    txtNetWeight.Text = NetWeight
                    txtAmount.Text = Amount
                    txtAmount.Select()
                Else
                    MessageBox.Show("Invalid percentage! Please try again.", "Invalid number!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    txtKotor.Text = ""
                    txtBasah.Text = ""
                    txtMuda.Text = ""
                    txtRosak.Text = ""
                    txtLain.Text = ""
                    txtBasah.Select()
                End If
            ElseIf txtPrice.Text = ("") Then
                txtPrice.Text = "0"
                Price = "0"
                If TotalDisc < 100 Then
                    Amount = ((((100 - TotalDisc) / 100) * TotalBerat) * Price)
                    NetWeight = ((100 - TotalDisc) / 100) * TotalBerat
                    txtNetWeight.Text = NetWeight
                    txtAmount.Text = Amount
                    txtAmount.Select()
                Else
                    MessageBox.Show("Invalid percentage! Please try again.", "Invalid number!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    txtKotor.Text = ""
                    txtBasah.Text = ""
                    txtMuda.Text = ""
                    txtRosak.Text = ""
                    txtLain.Text = ""
                    txtBasah.Select()
                End If
            Else
                MessageBox.Show("Invalid price! Please try again.", "Invalid number!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtPrice.Text = ""
                txtPrice.Select()
            End If
        End If
    End Sub
    Function CreateSQLAccServer()
        CreateSQLAccServer = CreateObject("SQLACC.BizApp")
    End Function

    Function GetCustomerData()
        Dim ComServer, RptObject, lDataSet, lDataSet2, lDateFrom, lDateTo
        Dim FoundCode As String
        'Step 1: Create Com Server object
        ComServer = CreateSQLAccServer() 'Create Com Server
        If Not ComServer.IsLogin Then 'if user hasn't logon to SQL application
            ComServer.Login("ADMIN", "hhc9696hhc", "C:\eStream\SQLAccounting\Share\Default.DCF", "ACC-0027.FDB")
        End If
        'Step 2: Find and Create the Report Objects
        RptObject = ComServer.RptObjects.Find("Supplier.RO")

        'Step 3: Spool parameters
        RptObject.Params.Find("AllAgent").Value = True
        RptObject.Params.Find("AllArea").Value = True
        RptObject.Params.Find("AllCompany").Value = False
        RptObject.Params.Find("AllCompanyCategory").Value = True
        RptObject.Params.Find("AllCurrency").Value = True
        RptObject.Params.Find("AllTerms").Value = True
        RptObject.Params.Find("SelectDate").Value = True
        RptObject.Params.Find("PrintActive").Value = True
        RptObject.Params.Find("PrintInactive").Value = False
        RptObject.Params.Find("PrintPending").Value = False
        RptObject.Params.Find("PrintProspect").Value = False
        RptObject.Params.Find("PrintSuspend").Value = False
        lDateFrom = CDate("January 1, 2000")
        lDateTo = Date.Now()

        RptObject.Params.Find("DateFrom").Value = lDateFrom
        RptObject.Params.Find("DateTo").Value = lDateTo
        RptObject.Params.Find("CompanyData").Value = Code

        'Step 4: Perform Report calculation 
        RptObject.CalculateReport()
        lDataSet = RptObject.DataSets.Find("cdsMain")
        lDataSet2 = RptObject.DataSets.Find("cdsBranch")
        'MsgBox("Count " & lDataSet.RecordCount)
        FoundCode = lDataSet.FindField("Code").AsString

        If FoundCode = Code Then
            'Step 5 Retrieve the output 
            txtMasukKg.Select()
            lDataSet.First
            While (Not lDataSet.eof)
                MsgBox(lDataSet.FindField("Code").AsString & " " & lDataSet.FindField("CompanyName").AsString)
                txtName.Text = lDataSet.FindField("CompanyName").AsString
                txtAdd1.Text = lDataSet.FindField("Address1").AsString
                txtAdd2.Text = lDataSet.FindField("Address2").AsString
                txtAdd3.Text = lDataSet.FindField("Address3").AsString
                txtAgentID.Text = lDataSet.FindField("Agent").AsString
                txtState.Text = lDataSet.FindField("Area").AsString
                lDataSet2.First
                While (Not lDataSet2.eof)
                    lDataSet2.Next
                End While
                lDataSet.Next
            End While
        Else
            MessageBox.Show("Supplier not found.", "Invalid supplier!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCode.Select()
            txtCode.Text = ""
        End If


#Disable Warning BC42105 ' Function doesn't return a value on all code paths
    End Function
#Enable Warning BC42105 ' Function doesn't return a value on all code paths


    Function InsertData()
        Dim ComServer, BizObject, lDate

        'Step 1: Create Com Server object
        ComServer = CreateSQLAccServer() 'Create Com Server
        If Not ComServer.IsLogin Then 'if user hasn't logon to SQL application
            ComServer.Login("ADMIN", "hhc9696hhc", "C:\estream\SQLAccounting\Share\Default.DCF", "ACC-0027.FDB")
        End If

        'Step 2: Find and Create the Biz Objects
        BizObject = ComServer.BizObjects.Find("PH_PI")

        Dim lMainDataSet As Object
        'Step 3: Set Dataset
        lMainDataSet = BizObject.DataSets.Find("MainDataSet")    'lMainDataSet contains master data
        Dim lDetailDataSet As Object
        lDetailDataSet = BizObject.DataSets.Find("cdsDocDetail") 'lDetailDataSet contains detail data  

        'Step 4 : Insert Data - Master
        lDate = Date.Now
        BizObject.[New]()
        lMainDataSet.FindField("DocKey").value = -1
        lMainDataSet.FindField("DocNo").AsString = "<<New>>"
        lMainDataSet.FindField("DocDate").value = lDate
        lMainDataSet.FindField("PostDate").value = lDate
        lMainDataSet.FindField("Code").value = Code
        lMainDataSet.FindField("Description").value = "Purchase"
        lMainDataSet.FindField("UDF_NOLORI").value = LorryNo
        lMainDataSet.FindField("UDF_WaktuMasuk").value = WaktuMasuk
        lMainDataSet.FindField("UDF_WaktuKeluar").value = WaktuKeluar
        lMainDataSet.FindField("UDF_Berat1").value = BeratMasuk
        lMainDataSet.FindField("UDF_Berat2").value = BeratKeluar
        lMainDataSet.FindField("UDF_BeratBersih").value = TotalBerat

        'Step 5: Insert Data - Detail
        'For Tax Inclusive = True with override Tax Amount
        lDetailDataSet.Append
        lDetailDataSet.FindField("DtlKey").value = -1
        lDetailDataSet.FindField("DocKey").value = -1
        lDetailDataSet.FindField("ItemCode").value = Product
        lDetailDataSet.FindField("Account").value = "610-000"
        'lDetailDataSet.FindField("Description").value = "Purchase Item A"
        lDetailDataSet.FindField("Qty").value = NetWeight
        lDetailDataSet.FindField("Tax").value = ""
        lDetailDataSet.FindField("TaxInclusive").value = 0
        lDetailDataSet.FindField("UnitPrice").value = Price
        lDetailDataSet.FindField("Amount").value = Amount
        lDetailDataSet.FindField("TaxAmt").value = 0
        lDetailDataSet.FindField("UDF_Basah").value = Basah
        lDetailDataSet.FindField("UDF_Kotor").value = Kotor
        lDetailDataSet.FindField("UDF_Muda").value = Muda
        lDetailDataSet.FindField("UDF_Rosak").value = Rosak
        lDetailDataSet.FindField("UDF_LainLain").value = Lain
        lDetailDataSet.FindField("UDF_Qty1").value = TotalBerat
        lDetailDataSet.FindField("UDF_Disc").value = TotalDisc

        lDetailDataSet.DisableControls
        lDetailDataSet.FindField("TaxInclusive").value = 1
        lDetailDataSet.EnableControls

        lDetailDataSet.FindField("Changed").value = "F"
        lDetailDataSet.Post
        'Step 6: Save Document
        BizObject.Save
        BizObject.Close
        'ComServer.Logout
        MsgBox("Done")
#Disable Warning BC42105 ' Function doesn't return a value on all code paths
    End Function
#Enable Warning BC42105 ' Function doesn't return a value on all code paths

    Function GetInfo()
        Dim ComServer, RptObject, lDataSet1, lDataSet2, lDateFrom, lDateTo
        'Step 1 Create Com Server object
        ComServer = CreateSQLAccServer() 'Create Com Server
        If Not ComServer.IsLogin Then 'if user hasn't logon to SQL application
            ComServer.Login("ADMIN", "hhc9696hhc", "C:\estream\SQLAccounting\Share\Default.DCF", "ACC-0027.FDB")
        End If

        'Step 2 Find and Create the Report Objects
        RptObject = ComServer.RptObjects.Find("Stock.Item.RO")

        'Step 3 Spool parameters
        RptObject.Params.Find("AllItem").AsBoolean = False
        RptObject.Params.Find("AllStockGroup").AsBoolean = True
        RptObject.Params.Find("AllCustomerPriceTag").AsBoolean = True
        RptObject.Params.Find("AllSupplierPriceTag").AsBoolean = True

        lDateFrom = CDate("January 01, 2000")
        lDateTo = Date.Now

        RptObject.Params.Find("HasAltStockItem").AsBoolean = False
        RptObject.Params.Find("HasBarcode").AsBoolean = False
        RptObject.Params.Find("HasBOM").AsBoolean = False
        RptObject.Params.Find("HasCategory").AsBoolean = False
        RptObject.Params.Find("HasCustomerItem").AsBoolean = False
        RptObject.Params.Find("HasOpeningBalance").AsBoolean = False
        RptObject.Params.Find("HasPurchasePrice").AsBoolean = False
        RptObject.Params.Find("HasSellingPrice").AsBoolean = False
        RptObject.Params.Find("HasSupplierItem").AsBoolean = False
        RptObject.Params.Find("PrintActive").AsBoolean = True
        RptObject.Params.Find("PrintInActive").AsBoolean = True
        RptObject.Params.Find("PrintNonStockControl").AsBoolean = True
        RptObject.Params.Find("PrintStockControl").AsBoolean = True
        RptObject.Params.Find("SelectCategory").AsBoolean = False
        RptObject.Params.Find("SelectDate").AsBoolean = False
        RptObject.Params.Find("SortBy").AsString = "Code"
        RptObject.Params.Find("ItemData").AsBlob = Product

        'Step 4 Perform Report calculation 
        RptObject.CalculateReport()
        lDataSet1 = RptObject.DataSets.Find("cdsMain")
        lDataSet2 = RptObject.DataSets.Find("cdsUOM") ' To link Master Data use Code

        'Step 5 Retrieve the output 
        lDataSet1.First
        While (Not lDataSet1.eof)
            MsgBox("Code " & lDataSet1.FindField("Code").AsString)
            'MsgBox("Price " & lDataSet1.FindField("RefCost").AsString)
            txtPrice.Text = lDataSet1.FindField("RefCost").AsString
            'MsgBox("Description " & lDataSet1.FindField("Description").AsString)
            'MsgBox("Balance Qty " & lDataSet1.FindField("BalSQty").AsString)
            lDataSet1.Next
        End While

        'MsgBox("Count " & lDataSet2.RecordCount)
        'lDataSet2.First
        'While (Not lDataSet2.eof)
        'MsgBox("Code " & lDataSet2.FindField("Code").AsString)
        'MsgBox("UOM " & lDataSet2.FindField("UOM").AsString)
        'MsgBox("Rate " & lDataSet2.FindField("Rate").AsString)
        'MsgBox("RefPrice " & lDataSet2.FindField("RefPrice").AsString)
        'lDataSet2.Next
        ' End While
#Disable Warning BC42105 ' Function doesn't return a value on all code paths
    End Function
#Enable Warning BC42105 ' Function doesn't return a value on all code paths

#Disable Warning IDE1006 ' Naming Styles
    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
#Enable Warning IDE1006 ' Naming Styles
        If txtAmount.Text IsNot ("") Then
            InsertData()
        Else MessageBox.Show("Some info are missing. Please check again", "Posting Fail!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        lblTimeNow.Text = Date.Now.ToString("hh:mm:ss")
        lblDate.Text = Date.Now.ToString("dd-MM-yy")
    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
#Enable Warning IDE1006 ' Naming Styles
        txtLorryNo.Text = ""
        LorryNo = ""
        txtCode.Text = ""
        Code = ""
        txtMasukKg.Text = ""
        BeratMasuk = 0
        txtKeluarKg.Text = ""
        BeratKeluar = 0
        txtProduct.Text = ""
        Product = ""
        txtBags.Text = ""
        Bags = 0
        txtGrossWT.Text = ""
        TotalBerat = 0
        txtBasah.Text = ""
        Basah = 0
        txtKotor.Text = ""
        Kotor = 0
        txtMuda.Text = ""
        Muda = 0
        txtLain.Text = ""
        Lain = 0
        txtRosak.Text = ""
        Rosak = 0
        txtPrice.Text = ""
        Price = 0
        txtPotongan.Text = ""
        TotalDisc = 0
        txtNetWeight.Text = ""
        NetWeight = 0
        txtName.Text = ""
        txtSubsidi.Text = ""
        txtAdd1.Text = ""
        txtAdd2.Text = ""
        txtAdd3.Text = ""
        txtAgentID.Text = ""
        txtState.Text = ""
        txtPesawahID.Text = ""
        txtTotal.Text = ""
        WaktuKeluar = ""
        WaktuMasuk = ""
        lblTimeIn.Text = ""
        lblTimeOut.Text = ""
        txtAmount.Text = ""
    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Private Sub btnHold1_Click(sender As Object, e As EventArgs) Handles btnHold1.Click
#Enable Warning IDE1006 ' Naming Styles
        If Form2.Visible = False Then
            Form2.Show()
        Else
            Form2.Hide()
        End If
    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Private Sub btnHold2_Click(sender As Object, e As EventArgs) Handles btnHold2.Click
#Enable Warning IDE1006 ' Naming Styles
        If Form3.Visible = False Then
            Form3.Show()
        Else
            Form3.Hide()
        End If
    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Private Sub btnHold3_Click(sender As Object, e As EventArgs) Handles btnHold3.Click
#Enable Warning IDE1006 ' Naming Styles
        If Form4.Visible = False Then
            Form4.Show()
        Else
            Form4.Hide()
        End If
    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Private Sub btnHold4_Click(sender As Object, e As EventArgs) Handles btnHold4.Click
#Enable Warning IDE1006 ' Naming Styles
        If Form5.Visible = False Then
            Form5.Show()
        Else
            Form5.Hide()
        End If
    End Sub

#Disable Warning IDE1006 ' Naming Styles
    Private Sub btnHold5_Click(sender As Object, e As EventArgs) Handles btnHold5.Click
#Enable Warning IDE1006 ' Naming Styles
        If Form6.Visible = False Then
            Form6.Show()
        Else
            Form6.Hide()
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If MessageBox.Show("Are you sure you want to exit?", "Close Form", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            e.Cancel = True
        End If
    End Sub
End Class
