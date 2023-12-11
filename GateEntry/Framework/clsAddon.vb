Imports System.IO
Imports SAPbouiCOM.Framework
Imports SAPbobsCOM
Public Class clsAddOn
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Dim oProgBarx As SAPbouiCOM.ProgressBar
    Public objGenFunc As Mukesh.SBOLib.GeneralFunctions
    Public objUIXml As Mukesh.SBOLib.UIXML
    Public ZB_row As Integer = 0
    Public SOMenuID As String = "0"
   
    Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
    Dim ret As Long
    Dim str As String
    Dim objForm As SAPbouiCOM.Form
    Dim MenuCount As Integer = 0
    Public objInward As clsGateInward
    Public objOutward As clsGateOutward
    Public objItemDetails As clsItemDetails
    Public objGRN As clsGRN
    Public HANA As Boolean = True
    ' Public HANA As Boolean = False

    Public HWKEY() As String = New String() {"L1653539483", "H1397589148"}
    Private Sub CheckLicense()

    End Sub
    Function isValidLicense() As Boolean
        Try
            objApplication.Menus.Item("257").Activate()
            Dim CrrHWKEY As String = objApplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
            objApplication.Forms.ActiveForm.Close()

            For i As Integer = 0 To HWKEY.Length - 1
                If HWKEY(i).Trim = CrrHWKEY.Trim Then
                    Return True
                End If
            Next
            MsgBox("Add-on installation failed due to license mismatch", MsgBoxStyle.OkOnly, "License Management")
            Return False
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return True
    End Function
    Public Sub Intialize()
        Dim objSBOConnector As New Mukesh.SBOLib.SBOConnector
        objApplication = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        objCompany = objSBOConnector.GetCompany(objApplication)
        Try
            createTables()
            createUDOs()
            createObjects()
            loadMenu()
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.ToString)
            End
        End Try
        If isValidLicense() Then
            objApplication.SetStatusBarMessage("Addon connected successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Else
            objApplication.SetStatusBarMessage("Failed To Connect, Please Check The License Configuration", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            objCompany.Disconnect()
            objApplication = Nothing
            objCompany = Nothing
            End
        End If
    End Sub

    Public Sub Intialize(ByVal args() As String)
        Try
            Dim oapplication As Application
            If (args.Length < 1) Then oapplication = New Application Else oapplication = New Application(args(0))
            objapplication = Application.SBO_Application
            If isValidLicense() Then
                objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objCompany = Application.SBO_Application.Company.GetDICompany()
                Try
                    createObjects()
                    createTables()
                    createUDOs()
                    loadMenu()
                Catch ex As Exception
                    objAddOn.objApplication.MessageBox(ex.ToString)
                    End
                End Try
                objApplication.StatusBar.SetText("Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                oapplication.Run()
            Else
                objapplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            'System.Windows.Forms.Application.Run()
        Catch ex As Exception
            objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub createUDOs()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        Dim ct1(1) As String
        'ct1(0) = "" -----------Need to check -------------------------
        'objUDFEngine.createUDO("MIVHTYPE", "MIVHTYPE", "VehicleType", ct1, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, False)
        ct1(0) = "MIGTOT1"
        objUDFEngine.createUDO("MIGTOT", "MIGTOT", "GTOutward", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)
        ct1(0) = "MIGTIN1"
        objUDFEngine.createUDO("MIGTIN", "MIGTIN", "GTInward", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)
     
    End Sub
    Private Sub createObjects()
        objGenFunc = New Mukesh.SBOLib.GeneralFunctions(objCompany)
        objUIXml = New Mukesh.SBOLib.UIXML(objApplication)
        objOutward = New clsGateOutward
        objInward = New clsGateInward
        objGRN = New clsGRN
        objItemDetails = New clsItemDetails
    End Sub
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            Select Case pVal.FormTypeEx
                Case clsGateOutward.Formtype
                    objOutward.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsGateInward.Formtype
                    objInward.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsItemDetails.Formtype
                    objItemDetails.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsGRN.formtype
                    objGRN.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
    End Sub
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent
        If BusinessObjectInfo.BeforeAction Then
           
        Else
            Try

                Select Case BusinessObjectInfo.FormTypeEx
                    Case clsGateOutward.Formtype
                        objOutward.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    Case clsGateInward.Formtype
                        objInward.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                    Case clsGRN.formtype
                        objGRN.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    Case "179", "180", "721", "940", "141"
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            CloseGateEntry(BusinessObjectInfo.FormTypeEx)
                        End If
                End Select

            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End Try
        End If
    End Sub
    Public Shared Function ErrorHandler(ByVal p_ex As Exception, ByVal objApplication As SAPbouiCOM.Application) As Boolean
        Dim sMsg As String = Nothing
        If p_ex.Message = "Form - already exists [66000-11]" Then
            Return True
            Exit Function  'ignore error
        End If
        Return False
    End Function
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        If pVal.BeforeAction Then
            Select Case pVal.MenuUID
               
            End Select
        Else
            Try
                Select Case pVal.MenuUID

                    Case clsGateOutward.Formtype
                        objOutward.LoadScreen()

                    Case clsGateInward.Formtype
                        objInward.LoadScreen()

                    Case "1282" ', "1290", "1288", "1289", "1291"
                        If objAddOn.objApplication.Forms.ActiveForm.UniqueID.Contains(clsGateOutward.Formtype) Then
                            objOutward.MenuEvent(pVal, BubbleEvent)
                        ElseIf objAddOn.objApplication.Forms.ActiveForm.UniqueID.Contains(clsGateInward.Formtype) Then
                            objInward.MenuEvent(pVal, BubbleEvent)
                        End If
                    Case "ditem"
                    Case "3073"
                        '   objItemMaster.LoadForm(objApplication.Forms.ActiveForm.UniqueID)

                End Select
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub
    Private Sub loadMenu()
        If objApplication.Menus.Item("43520").SubMenus.Exists("GT") Then Return
        MenuCount = objApplication.Menus.Item("43520").SubMenus.Count

        CreateMenu(Windows.Forms.Application.StartupPath + "\jc1.png", MenuCount + 1, "Gate Entry Module", SAPbouiCOM.BoMenuType.mt_POPUP, "GT", objApplication.Menus.Item("43520"))
        CreateMenu("", 1, "Gate Entry Outward", SAPbouiCOM.BoMenuType.mt_STRING, clsGateOutward.Formtype, objApplication.Menus.Item("GT"))
        CreateMenu("", 2, "Gate Entry Inward", SAPbouiCOM.BoMenuType.mt_STRING, clsGateInward.Formtype, objApplication.Menus.Item("GT"))

    End Sub
    Private Function CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.AddEx(oMenuPackage)
        Catch ex As Exception
            objApplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
        Return ParentMenu.SubMenus.Item(UniqueID)
    End Function
    Private Sub createTables()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        objAddOn.objApplication.SetStatusBarMessage("Creating Tables Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
        ' WriteSMSLog("0")
        'Gate Entry outward
        objUDFEngine.CreateTable("MIGTOT", "GTEntry Outward", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddAlphaField("@MIGTOT", "loc", "Location", 15)
        objUDFEngine.AddAlphaField("@MIGTOT", "sname", "Security Name", 15)
        objUDFEngine.AddAlphaField("@MIGTOT", "type", "Type", 15)
        objUDFEngine.AddAlphaField("@MIGTOT", "partyid", "Party Id", 15)
        objUDFEngine.AddAlphaField("@MIGTOT", "partynm", "Party Name", 100)
        objUDFEngine.AddNumericField("@MIGTOT", "nopack", "No of Packages", 10)
        objUDFEngine.AddDateField("@MIGTOT", "outtime", "Out Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objUDFEngine.AddAlphaField("@MIGTOT", "lrno", "LRNo", 15)
        objUDFEngine.AddDateField("@MIGTOT", "docdate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        ' objUDFEngine.AddAlphaField("@MIGTOT", "vehtype", "Vehicle Type", 15)
        objUDFEngine.AddAlphaField("@MIGTOT", "vehno", "Vehicle No", 15)
        '  objUDFEngine.AddAlphaField("@MIGTOT", "vehname", "Vehicle Name", 50)
        objUDFEngine.AddAlphaField("@MIGTOT", "transnm", "Transporter Name", 15)
        objUDFEngine.AddDateField("@MIGTOT", "lrdate", "LR Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIGTOT", "geno", "Gate Entry No", 25)
        objUDFEngine.AddAlphaField("@MIGTOT", "wcno", "Weight Challan No", 25)
        objUDFEngine.AddDateField("@MIGTOT", "cutdate", "Cutoff Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objUDFEngine.CreateTable("MIGTOT1", "GT Outward Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@MIGTOT1", "basenum", "Base Document Number", 15)
        objUDFEngine.AddAlphaField("@MIGTOT1", "baseline", "Base Line", 15)
        objUDFEngine.AddAlphaField("@MIGTOT1", "basentry", "Base Entry", 15)
        objUDFEngine.AddAlphaField("@MIGTOT1", "basetype", "Base Type", 15)
        objUDFEngine.AddAlphaField("@MIGTOT1", "itemcode", "Item/Service Code", 50)
        objUDFEngine.AddAlphaField("@MIGTOT1", "itemdesc", "Item/Service Description", 100)
        'objUDFEngine.AddAlphaField("@MIGTOT1", "itemdet", "Item/Service Details", 254)
        '-------------------- Need to add text field --------------------------------------------------
        objUDFEngine.AddFloatField("@MIGTOT1", "qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIGTOT1", "unitpric", "Unit Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@MIGTOT1", "linetot", "Line Total", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddFloatField("@MIGTOT1", "orderqty", "Order Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIGTOT1", "pendqty", "Pending Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@MIGTOT1", "remarks", "Remarks ", 254)
        objUDFEngine.AddAlphaField("@MIGTOT1", "uom", "UoM ", 10)
        objUDFEngine.AddAlphaMemoField("@MIGTOT1", "itemdet1", "ItemDetails", 4)

        'Gate Entry Inward

        objUDFEngine.CreateTable("MIGTIN", "GTEntry Inward", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddAlphaField("@MIGTIN", "loc", "Location", 15)
        objUDFEngine.AddAlphaField("@MIGTIN", "type", "Type", 15)
        objUDFEngine.AddAlphaField("@MIGTIN", "partyid", "Party Id", 25)
        objUDFEngine.AddAlphaField("@MIGTIN", "partynm", "Party Name", 100)
        objUDFEngine.AddAlphaField("@MIGTIN", "supdcno", "Supplier DC No", 15)
        objUDFEngine.AddAlphaField("@MIGTIN", "supinvno", "Supplier InvNo", 15)
        objUDFEngine.AddAlphaField("@MIGTIN", "mdvtcprd", "ModVatCopy Received", 5)
        objUDFEngine.AddNumericField("@MIGTIN", "nopack", "No of Packages", 10)
        objUDFEngine.AddDateField("@MIGTIN", "intime", "Out Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objUDFEngine.AddAlphaField("@MIGTIN", "lrno", "LR No", 15)
        objUDFEngine.AddDateField("@MIGTIN", "lrdate", "LR Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIGTIN", "sname", "Security Name", 15)
        objUDFEngine.AddDateField("@MIGTIN", "docdate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddDateField("@MIGTIN", "supdcdt", "Supplier DC Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddDateField("@MIGTIN", "supinvdt", "Supplier InvDate", SAPbobsCOM.BoFldSubTypes.st_None)
        ' objUDFEngine.AddAlphaField("@MIGTIN", "vehtype", "Vehicle Type", 15)
        objUDFEngine.AddAlphaField("@MIGTIN", "vehno", "Vehicle No", 15)
        '  objUDFEngine.AddAlphaField("@MIGTIN", "vehname", "Vehicle Name", 50)
        objUDFEngine.AddAlphaField("@MIGTIN", "transnm", "Transporter Name", 50)
        objUDFEngine.AddAlphaField("@MIGTIN", "copyto", "Copy To", 15)
        objUDFEngine.AddAlphaField("@MIGTIN", "geno", "Gate Entry No", 25)
        objUDFEngine.AddAlphaField("@MIGTIN", "wcno", "Weight Challan No", 25)
        objUDFEngine.AddDateField("@MIGTIN", "cutdate", "Cutoff Date", SAPbobsCOM.BoFldSubTypes.st_None)


        objUDFEngine.CreateTable("MIGTIN1", "GT Inward Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@MIGTIN1", "basetype", "Base Type", 15)
        objUDFEngine.AddAlphaField("@MIGTIN1", "basenum", "Base Document Number", 15)
        objUDFEngine.AddAlphaField("@MIGTIN1", "baseline", "Base Line", 15)
        objUDFEngine.AddAlphaField("@MIGTIN1", "basentry", "Base Entry", 15)
        objUDFEngine.AddAlphaField("@MIGTIN1", "itemcode", "Item/Service Code", 15)
        objUDFEngine.AddAlphaField("@MIGTIN1", "itemdesc", "Item/Service Description", 15)
        '  objUDFEngine.AddAlphaField("@MIGTIN1", "itemdet", "Item/Service Details", 15)
        objUDFEngine.AddFloatField("@MIGTIN1", "qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIGTIN1", "unitpric", "Unit Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@MIGTIN1", "linetot", "Line Total", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddFloatField("@MIGTIN1", "orderqty", "Order Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIGTIN1", "pendqty", "Pending Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@MIGTIN1", "remarks", "Remarks ", 100)
        objUDFEngine.AddAlphaField("@MIGTIN1", "uom", "UoM ", 10)
        objUDFEngine.AddAlphaMemoField("@MIGTIN1", "itemdet1", "ItemDetails", 4)
        ' vehicle type master
        objUDFEngine.CreateTable("MIVHTYPE", "Vehicle Type", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        ' Marketting documents
        objUDFEngine.AddAlphaField("INV!", "getype", "GateEntry Type", 15)
        objUDFEngine.AddAlphaField("INV1", "geentry", "GateEntry DocEntry", 15)
        objUDFEngine.AddAlphaField("INV1", "gedocno", "GateEntry DocNum", 15)
        objUDFEngine.AddAlphaField("OWTR", "VENDORCODE", "Vendor Code", 15)
        objUDFEngine.AddAlphaField("OWTR", "VENDORNAME", "Vendor Name", 50)
        objUDFEngine.AddAlphaField("OPDN", "gever", "GE Verification", 15)

        '*******************  Table ******************* START********************************* END
    End Sub
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        If eventInfo.BeforeAction Then
        Else
            If eventInfo.FormUID.Contains("MIREJDET") And (eventInfo.ItemUID = "13") And eventInfo.Row > 0 Then

                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try

                    If objAddOn.objApplication.Menus.Exists("ditem") Then
                        objAddOn.objApplication.Menus.RemoveEx("ditem")
                    End If
                Catch ex As Exception

                End Try
                Try

                    oMenuItem = objAddOn.objApplication.Menus.Item("1280").SubMenus.Item("ditem")
                    ZB_row = eventInfo.Row
                Catch ex As Exception
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "ditem"
                    oCreationPackage.String = "Delete Row"
                    oCreationPackage.Enabled = True

                    oMenuItem = objAddOn.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    oMenus.AddEx(oCreationPackage)
                    ZB_row = eventInfo.Row
                End Try
                If eventInfo.ItemUID <> "13" Then
                    '   Dim oMenuItem As SAPbouiCOM.MenuItem
                    '  Dim oMenus As SAPbouiCOM.Menus
                    Try
                        objAddOn.objApplication.Menus.RemoveEx("ditem")
                    Catch ex As Exception
                        ' MessageBox.Show(ex.Message)
                    End Try
                End If
            End If
            End If
    End Sub

    Private Sub objApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            Try
                ' objUIXml.LoadMenuXML("RemoveMenu.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded)
                If objCompany.Connected Then objCompany.Disconnect()
                objCompany = Nothing
                objApplication = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objApplication)
                GC.Collect()
            Catch ex As Exception
            End Try
            End
        End If
    End Sub

    Private Sub applyFilter()
        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        oFilters = New SAPbouiCOM.EventFilters
        'Item Master Data 
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)


    End Sub
    Public Sub WriteSMSLog(ByVal Str As String)
        Dim fs As FileStream
        Dim chatlog As String = Windows.Forms.Application.StartupPath & "\Log_" & Today.ToString("yyyyMMdd") & ".txt"
        If File.Exists(chatlog) Then
        Else
            fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
            fs.Close()
        End If
        ' Dim objReader As New System.IO.StreamReader(chatlog)
        Dim sdate As String
        sdate = Now
        'objReader.Close()
        If System.IO.File.Exists(chatlog) = True Then
            Dim objWriter As New System.IO.StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
            objWriter.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            ' MsgBox("Failed to send message!")
        End If
    End Sub
    Private Sub CloseGateEntry(ByVal FormUID As String)
        Dim strSQL As String
        Dim objRS As SAPbobsCOM.Recordset
        Dim oForm As SAPbouiCOM.Form
        oForm = objAddOn.objApplication.Forms.GetForm(FormUID, 1)
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim MatrixNo As String
        Select Case FormUID
            Case "179" 'ARCreditMemo
                MatrixNo = "38"
            Case "180" 'SalesReturn
                MatrixNo = "38"
            Case "721" ' Goods Receipt
                MatrixNo = "13"
            Case "940" 'Stock transfer
                MatrixNo = "23"
            Case "141" ' APInvoice
                MatrixNo = "39"
        End Select
        oMatrix = oForm.Items.Item(MatrixNo).Specific
        Dim GEDocEntry As String = Trim(oMatrix.Columns.Item("U_geentry").Cells.Item(1).Specific.value)
        If objAddOn.HANA Then
            strSQL = "UPDATE ""@MIGTIN"" set ""Status""='C' where ""DocEntry"" ='" & GEDocEntry & "'"
        Else

            strSQL = "Update [@MIGTIN] set Status='C' where DocEntry='" & GEDocEntry & "'"
        End If
        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRS.DoQuery(strSQL)
        objRS = Nothing
    End Sub
    Private Sub addJobCardReporttype()
        'Dim rptTypeService As SAPbobsCOM.ReportTypesService
        'Dim newType As SAPbobsCOM.ReportType
        'Dim newtypeParam As SAPbobsCOM.ReportTypeParams
        'Dim newReportParam As SAPbobsCOM.ReportLayoutParams
        'Dim ReportExists As Boolean = False
        'Try


        '    Dim newtypesParam As SAPbobsCOM.ReportTypesParams
        '    rptTypeService = objAddOn.objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
        '    newtypesParam = rptTypeService.GetReportTypeList

        '    Dim i As Integer
        '    For i = 0 To newtypesParam.Count - 1
        '        If newtypesParam.Item(i).TypeName = clsJobCard.FormType And newtypesParam.Item(i).MenuID = clsJobCard.FormType Then
        '            ReportExists = True
        '            Exit For
        '        End If
        '    Next i

        '    If Not ReportExists Then
        '        rptTypeService = objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
        '        newType = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)


        '        newType.TypeName = clsJobCard.FormType
        '        newType.AddonName = "JC2Addon"
        '        newType.AddonFormType = clsJobCard.FormType
        '        newType.MenuID = clsJobCard.FormType
        '        newtypeParam = rptTypeService.AddReportType(newType)

        '        Dim rptService As SAPbobsCOM.ReportLayoutsService
        '        Dim newReport As SAPbobsCOM.ReportLayout
        '        rptService = objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
        '        newReport = rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout)
        '        newReport.Author = objCompany.UserName
        '        newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal
        '        newReport.Name = clsJobCard.FormType
        '        newReport.TypeCode = newtypeParam.TypeCode

        '        newReportParam = rptService.AddReportLayout(newReport)

        '        newType = rptTypeService.GetReportType(newtypeParam)
        '        newType.DefaultReportLayout = newReportParam.LayoutCode
        '        rptTypeService.UpdateReportType(newType)

        '        Dim oBlobParams As SAPbobsCOM.BlobParams
        '        oBlobParams = objCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
        '        oBlobParams.Table = "RDOC"
        '        oBlobParams.Field = "Template"
        '        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment
        '        oKeySegment = oBlobParams.BlobTableKeySegments.Add
        '        oKeySegment.Name = "DocCode"
        '        oKeySegment.Value = newReportParam.LayoutCode

        '        Dim oFile As FileStream
        '        oFile = New FileStream(Application.StartupPath + "\JobCard.rpt", FileMode.Open)
        '        Dim fileSize As Integer
        '        fileSize = oFile.Length
        '        Dim buf(fileSize) As Byte
        '        oFile.Read(buf, 0, fileSize)
        '        oFile.Dispose()

        '        Dim oBlob As SAPbobsCOM.Blob
        '        oBlob = objCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob)
        '        oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)
        '        objCompany.GetCompanyService.SetBlob(oBlobParams, oBlob)
        '    End If
        'Catch ex As Exception
        '    objApplication.MessageBox(ex.ToString)
        'End Try

    End Sub

    Private Sub objApplication_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles objApplication.LayoutKeyEvent

        ''BubbleEvent = True
        'If eventInfo.BeforeAction = True Then
        '    If eventInfo.FormUID.Contains(clsJobCard.FormType) Then
        '        objJobCard.LayoutKeyEvent(eventInfo, BubbleEvent)
        '    End If
        'End If
    End Sub
End Class


