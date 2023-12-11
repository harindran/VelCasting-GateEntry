Public Class clsGateInward
    Public Const Formtype As String = "MIGTIN"
    Dim objForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim Header As SAPbouiCOM.DBDataSource
    Dim objRS As SAPbobsCOM.Recordset
    Dim strSQL As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objHeader As SAPbouiCOM.DBDataSource
    Dim objLine As SAPbouiCOM.DBDataSource
    Public Sub LoadScreen()
        objForm = objAddOn.objUIXml.LoadScreenXML("Inward.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
        objHeader = objForm.DataSources.DBDataSources.Item("@MIGTIN")
        objLine = objForm.DataSources.DBDataSources.Item("@MIGTIN1")
        InitForm(objForm.UniqueID)
        objForm.Visible = True
    End Sub
    Private Sub NeedToBeDone()
        'close status, form shpuld be frozen	
        'GRN can have the GE number link button	
        'Copy to & copy From button should disabled when status close
        'Next document number is not loaded
    End Sub
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try


            If pval.BeforeAction Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        If pval.ItemUID = "36" And pval.ColUID = "9" Then ' line total calculation
                            If Not QtyValidation(FormUID, pval.Row) Then
                                BubbleEvent = False
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = objAddOn.objApplication.Forms.Item(FormUID)
                        If pval.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If Not Validate(FormUID) Then
                                BubbleEvent = False
                            End If
                        End If
                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pval.ItemUID = "37" Then ' copy From
                            CopyFrom(FormUID)
                        ElseIf pval.ItemUID = "38" Then
                            CopyTo(FormUID)
                        ElseIf pval.ItemUID = "1" And pval.ActionSuccess And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            InitForm(FormUID)

                        End If

                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If pval.ItemUID = "36" And (pval.ColUID = "9" Or pval.ColUID = "11") Then ' line total calculation
                            LineTotalCalc(FormUID, pval.Row)

                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pval.ItemUID = "10" Then 'Partyid
                            ChooseFromListBP(FormUID, pval)
                        End If
                End Select

            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        If BusinessObjectInfo.BeforeAction Then
            Try

                If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                    If Validate(BusinessObjectInfo.FormUID) Then
                    Else
                        BubbleEvent = False
                    End If
                End If

            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End Try
        Else
            'Try
            '    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.ActionSuccess Then
            '        objForm = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
            '        If objForm.Items.Item("25").Specific.selected.value = "C" Then


            '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            '        Else
            '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            '        End If
            '    End If
            'Catch ex As Exception

            'End Try
        End If
    End Sub
    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        Try
            Select Case pval.MenuUID
                Case "1282"
                    If pVal.BeforeAction = False Then InitForm(objAddOn.objApplication.Forms.ActiveForm.UniqueID)
                    'Case "1289"
                    '    If pVal.BeforeAction = False Then Me.UpdateMode()
                    'Case "1293"
                    'Case "1281"
                    '    If pVal.BeforeAction = False Then

                    '   End If
            End Select
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Menu Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Public Sub InitForm(ByVal FormUID As String)
        LoadType(FormUID)
        LoadSeries(FormUID)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objMatrix.Columns.Item("9").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        objMatrix.Columns.Item("12").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    End Sub
    Public Sub LoadSeries(ByVal FormUID As String)

        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        '---------------- Load locations ------------
        objCombo = objForm.Items.Item("4").Specific
        If objCombo.ValidValues.Count = 0 Then
            If objAddOn.HANA Then
                strSQL = "select ""Code"", ""Location"" from OLCT"
            Else
                strSQL = "select Code, Location from OLCT"
            End If

           
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(strSQL)
            While Not objRS.EoF
                objCombo.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(1).Value)
                objRS.MoveNext()
            End While
            objRS = Nothing
        End If
        '----------------Load series --------------
        objCombo = objForm.Items.Item("20").Specific
        objCombo.ValidValues.LoadSeries(Formtype, SAPbouiCOM.BoSeriesMode.sf_Add)
        If objCombo.ValidValues.Count > 0 Then objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        objHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetDocNum(Formtype, CInt(objForm.Items.Item("20").Specific.Selected.value)))
        objForm.Items.Item("23").Specific.String = "A" ' current date
        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            objCombo = objForm.Items.Item("8").Specific
            objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        End If
        '------------ Load Security-------------
        objCombo = objForm.Items.Item("6").Specific
        If objCombo.ValidValues.Count = 0 Then
            If objAddOn.HANA Then
                strSQL = "SELECT T0.""empID"", T0.""firstName"" || ' ' || T0.""lastName"" as ""empName"", T1.""Name"" FROM OHEM T0 INNER JOIN OUDP T1 ON T0.""dept"" = T1.""Code"" WHERE T1.""Name"" ='Security' ;"
            Else
                strSQL = "SELECT T0.[empID], T0.[firstName] + ' ' + T0.[lastName] as empName, T1.[Name] FROM OHEM T0  INNER JOIN OUDP T1 ON T0.[dept] = T1.[Code] WHERE T1.[Name] ='Security'"
            End If

            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(strSQL)
            While Not objRS.EoF
                objCombo.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(1).Value)
                objRS.MoveNext()
            End While
            objRS = Nothing
        End If
    End Sub

    Private Sub LoadType(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objCombo = objForm.Items.Item("8").Specific
        If objCombo.ValidValues.Count = 0 Then

            objCombo.ValidValues.Add("PO", "Purchase Order") '--
            objCombo.ValidValues.Add("SR", "Sales Return with Invoice") '--
            objCombo.ValidValues.Add("DR", "Customer Delivery Return") '--
            objCombo.ValidValues.Add("GR", "GRN")
            objCombo.ValidValues.Add("SC", "Sales Credit Memo")
            objCombo.ValidValues.Add("WI", "Sales Return without Invoice")
            objCombo.ValidValues.Add("DC", "Returnable DC")
            objCombo.ValidValues.Add("JR", "JobOrder Repair")
            objCombo.ValidValues.Add("SP", "Scrap Receipt")
            objCombo.ValidValues.Add("WM", "Without Process Material")
            objCombo.ValidValues.Add("RW", "Job Order Rework ")
            objCombo.ValidValues.Add("ST", "Stock Transfer")
            objCombo.ValidValues.Add("SO", "Service Order")
            objCombo.ValidValues.Add("JO", "Job Order Regular")
            objCombo.ValidValues.Add("JW", "Job Rework")
            objCombo.ValidValues.Add("CP", "Cash Purchase")
            objCombo.ValidValues.Add("MI", "Material Inward")
            objCombo.ValidValues.Add("HR", "Service Invoice HR")
            'objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        End If

    End Sub
    Private Sub ChooseFromListBP(ByVal FormUID As String, ByRef pval As SAPbouiCOM.ItemEvent)
        Dim CFLEvent As SAPbouiCOM.ChooseFromListEvent
        CFLEvent = pval
        Dim datatable As SAPbouiCOM.DataTable
        If CFLEvent.ChooseFromListUID = "BP_CFL" Then
            datatable = CFLEvent.SelectedObjects()
            Try
                objHeader.SetValue("U_partyid", 0, datatable.GetValue("CardCode", 0))
                objHeader.SetValue("U_partynm", 0, datatable.GetValue("CardName", 0))
            Catch ex As Exception

            End Try

        End If
    End Sub
      Private Sub LineTotalCalc(ByVal FormUID As String, ByVal RowID As Integer)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objLine = objForm.DataSources.DBDataSources.Item("@MIGTIN1")
        'objMatrix.GetLineData(RowID)
        Dim linetotal As Double
        linetotal = CDbl(objMatrix.Columns.Item("9").Cells.Item(RowID).Specific.value) * CDbl(objLine.GetValue("U_unitpric", RowID - 1))

        objMatrix.Columns.Item("12").Cells.Item(RowID).Specific.value = CStr(linetotal)
        'objLine.SetValue("U_linetot", RowID - 1, linetotal)
        ' MsgBox(CStr(objLine.GetValue("U_linetot", RowID - 1)))
        'objMatrix.SetLineData(RowID)
        objForm.Update()
        objForm.Refresh()
    End Sub
    Private Function Validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Try


            If Trim(objForm.Items.Item("4").Specific.Value) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please select Location name")
                Return False
            ElseIf Trim(objForm.Items.Item("6").Specific.Value) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please select Security name")
                Return False
            ElseIf Trim(objForm.Items.Item("10").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Party details")
                Return False
                'ElseIf Trim(objForm.Items.Item("14").Specific.String) = "" Then
                '    objAddOn.objApplication.SetStatusBarMessage("Please fill up No of packages")
                '    Return False
            ElseIf Trim(objForm.Items.Item("16").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up In time")
                Return False
            ElseIf Trim(objForm.Items.Item("23").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Date")
                Return False

                'ElseIf Trim(objForm.Items.Item("18").Specific.String) = "" Then
                '    objAddOn.objApplication.SetStatusBarMessage("Please fill up LR Number")
                '    Return False

            ElseIf Trim(objForm.Items.Item("27").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Gate Entry Number")
                Return False

            ElseIf Trim(objForm.Items.Item("31").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Vehicle Number")
                Return False

            ElseIf Trim(objForm.Items.Item("33").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Transporter Name")
                Return False
            End If
            objMatrix = objForm.Items.Item("36").Specific
            If objMatrix.RowCount = 0 Then
                objAddOn.objApplication.SetStatusBarMessage("Minimum one Line Item is Required.. ")
                Return False
            Else
                If objMatrix.Columns.Item("4").Cells.Item(1).Specific.value = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Minimum one Line Item is Required.. ")
                    Return False
                End If
            End If
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage("Please check the mandatory fields")
            Return False
        End Try
        Return True
    End Function
    Private Function QtyValidation(ByVal FormUID As String, ByVal RowID As Integer) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objCombo = objForm.Items.Item("8").Specific
        Select Case (UCase(Trim(objCombo.Value)))
            Case "DC", "JR", "SR", "WM", "RW", "ST", "SO", "CP", "JO", "JW", "MI", "HR"
                Return True
        End Select



        objMatrix = objForm.Items.Item("36").Specific
        objLine = objForm.DataSources.DBDataSources.Item("@MIGTIN1")
        objMatrix.GetLineData(RowID)
        If CDbl(objMatrix.Columns.Item("9").Cells.Item(RowID).Specific.value) > CDbl(objLine.GetValue("U_pendqty", 0)) Then
            objAddOn.objApplication.SetStatusBarMessage("Quantity exceeds pending quantity", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If
        Return True
    End Function
    Private Sub OpenColumns(ByVal FormUID As String)
        Select Case (UCase(Trim(objCombo.Value)))
            Case "DC", "JR", "SR", "WM", "RW", "ST", "SO", "CP", "JO", "JW", "MI", "HR"
                objMatrix.Columns.Item("11").Editable = True
            Case Else
                objMatrix.Columns.Item("11").Editable = False
        End Select

    End Sub

    Private Sub CopyTo(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        If Not objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            objAddOn.objApplication.MessageBox("Form should be in OK mode")
            Exit Sub
        End If
        '  objCombo = objForm.Items.Item("49").Specific
        objCombo = objForm.Items.Item("8").Specific
        If Trim(objCombo.Value) = "" Then Exit Sub
        Select Case Trim(objCombo.Value)
            Case "GR"
                objAddOn.objApplication.MessageBox("Please generate GRPO document with copy from PO and Do GE Verification")
            Case "CP"
                CopyToGRN(FormUID)
            Case "SC"
                CopyToARCreditMemo(FormUID) ' 179,180,721,940,141
            Case "WI" ' 
                CopyToSalesReturn(FormUID) ' 180
            Case "MI"
                CopyToGoodsReceipt(FormUID) '721
            Case "DC", "JR", "SP", "WM", "RW", "ST", "SO", "JO", "JW"
                CopyToStockTransfer(FormUID) '940
            Case "HR"
                CopyToAPInvoice(FormUID) '141
        End Select
    End Sub
    Private Sub CopyFrom(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objHeader = objForm.DataSources.DBDataSources.Item("@MIGTIN")
        objCombo = objForm.Items.Item("8").Specific
        If Trim(objCombo.Value) = "" Then Exit Sub
        ' If objForm.Items.Item("10").Specific.string <> "" Then
        OpenColumns(FormUID)
        objAddOn.objItemDetails.LoadScreen(Formtype, objCombo.Value, objForm.Items.Item("10").Specific.string, objHeader.GetValue("U_cutdate", 0))
        ' Else
        'objAddOn.objApplication.MessageBox("Please select Party id")
        'End If

    End Sub
    Private Sub CopyToStockTransfer(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("3080")
        'Matrix 23; form 940
        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("940", 1)
        CopyToForm.Items.Item("3").Specific.String = objHeader.GetValue("U_partyid", 0)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("23").Specific

        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            ' CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
            CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
            CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
        Next

    End Sub
    Private Sub CopyToARCreditMemo(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("2085")

        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("179", 1)
        CopyToForm.Items.Item("4").Specific.String = objHeader.GetValue("U_partyid", 0)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("38").Specific

        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            ' CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
            CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
            CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
        Next

    End Sub
    Private Sub CopyToGRN(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        Try

       
        objAddOn.objApplication.ActivateMenuItem("2306")

        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("143", 1)
        CopyToForm.Items.Item("4").Specific.String = objHeader.GetValue("U_partyid", 0)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("38").Specific
            CopyToForm.Items.Item("TVer").Specific.String = objHeader.GetValue("DocEntry", 0)
        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            ' CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
                CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
                CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
                CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
                'Dim objGRN As SAPbobsCOM.Documents
                'objGRN = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                'objGRN.Lines.BaseType = objMatrix.Columns.Item("1A").Cells.Item(i).Specific.String
                'objGRN.Lines.BaseEntry = objMatrix.Columns.Item("2").Cells.Item(i).Specific.String
                'objGRN.Lines.BaseLine = objMatrix.Columns.Item("3").Cells.Item(i).Specific.String

        Next
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub CopyToSalesReturn(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("2052")

        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("180", 1)
        CopyToForm.Items.Item("4").Specific.String = objHeader.GetValue("U_partyid", 0)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("38").Specific

        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            'CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
            CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
            CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
        Next

    End Sub
    Private Sub CopyToGoodsReceipt(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("3078")

        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("721", 1)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("13").Specific

        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            'CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
            CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
            CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
        Next

    End Sub
    Private Sub CopyToAPInvoice(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("2308")

        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("141", 1)
        CopyToForm.Items.Item("4").Specific.String = objHeader.GetValue("U_partyid", 0)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("39").Specific

        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            'CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("11").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
            CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
            CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
        Next

    End Sub
    Private Sub viewMode(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)

    End Sub
    ''Private Sub updatePO(ByVal FormUID As String)

    ''    objForm = objAddOn.objApplication.Forms.Item(FormUID)
    ''    objCombo = objForm.Items.Item("20").Specific
    ''    If objCombo.Value = "PO" Then
    ''        Dim objPO As SAPbobsCOM.Documents
    ''        objPO = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
    ''        For i As Integer = 1 To objMatrix.RowCount
    ''            Dim POEntry As Integer = objHeader.GetValue("DocEntry", 0)
    ''            If objPO.GetByKey(POEntry) Then
    ''                objPO.Lines.SetCurrentLine(CInt(objMatrix.Columns.Item("0").Cells.Item(i).Specific.String))
    ''                objPO.Lines.UserFields.Fields.Item("U_getype").Value = objHeader.GetValue("U_type", 0)
    ''                objPO.Lines.UserFields.Fields.Item("U_gedocno").Value = objHeader.GetValue("DocNum", 0)
    ''                objPO.Lines.UserFields.Fields.Item("U_geentry").Value = objHeader.GetValue("DocEntry", 0)
    ''            End If
    ''        Next

    ''    End If
    ''End Sub
End Class
