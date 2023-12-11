Public Class clsGateOutward
    Public Const Formtype As String = "MIGTOT"
    Dim objForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim Header As SAPbouiCOM.DBDataSource
    Dim objRS As SAPbobsCOM.Recordset
    Dim strSQL As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objHeader As SAPbouiCOM.DBDataSource
    Dim objLine As SAPbouiCOM.DBDataSource
    Public Sub LoadScreen()
        objForm = objAddOn.objUIXml.LoadScreenXML("Outward.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)

        objHeader = objForm.DataSources.DBDataSources.Item("@MIGTOT")
        objLine = objForm.DataSources.DBDataSources.Item("@MIGTOT1")
        InitForm(objForm.UniqueID)
        objForm.Visible = True
    End Sub
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try


            If pval.BeforeAction Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        If pval.ItemUID = "36" And pval.ColUID = "9" Then ' line total calculation
                            If Not QtyValidation(FormUID, pval.Row) Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        objForm = objAddOn.objApplication.Forms.Item(FormUID)
                        If pval.ItemUID = "1" Then
                            If pval.BeforeAction = True And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                If Validate(FormUID) = False Then
                                    '     System.Media.SystemSounds.Asterisk.Play()
                                    BubbleEvent = False
                                    ' objAddOn.objApplication.SetStatusBarMessage("ItemEvent")
                                    Exit Sub
                                End If
                            End If
                        End If
                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pval.ItemUID = "37" Then ' copy From
                            CopyFrom(FormUID)
                        ElseIf pval.ItemUID = "38" Then
                            CopyToStockTransfer(FormUID)
                        ElseIf pval.ItemUID = "1" And pval.ActionSuccess And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            InitForm(FormUID)
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If pval.ItemUID = "36" And pval.ColUID = "9" Then ' line total calculation
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
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        If BusinessObjectInfo.BeforeAction Then
            Try

                If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                    If Validate(BusinessObjectInfo.FormUID) Then
                    Else
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If

            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End Try
        Else
            If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                InitForm(BusinessObjectInfo.FormUID)
            End If
        End If
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
        Dim StrDocNum = objForm.BusinessObject.GetNextSerialNumber(Trim(objForm.Items.Item("20").Specific.Selected.value), objForm.BusinessObject.Type)
        objForm.DataSources.DBDataSources.Item("@MIGTOT").SetValue("DocNum", 0, StrDocNum)
        objForm.Items.Item("23").Specific.String = "A" ' current date
        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            objCombo = objForm.Items.Item("8").Specific
            objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        End If
     
        '------------ Load Security Name-------------
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
            objCombo.ValidValues.Add("SI", "Sales Invoice")
            objCombo.ValidValues.Add("SR", "Supplier Return")
            objCombo.ValidValues.Add("JO", "Job order DC")
            objCombo.ValidValues.Add("SO", "Service Order DC")
            objCombo.ValidValues.Add("RT", "Returnable DC")
            objCombo.ValidValues.Add("NR", "Non-Returnable DC")
            objCombo.ValidValues.Add("RW", "Rework DC")
            objCombo.ValidValues.Add("RJ", "Rejection DC")
            objCombo.ValidValues.Add("SC", "Supplier Credit Memo")
            objCombo.ValidValues.Add("ST", "Stock Transfer")
            objCombo.ValidValues.Add("IU", "InterUnit DC")
            objCombo.ValidValues.Add("JW", "Job Work")
            objCombo.ValidValues.Add("MO", "Material Outward")
            '   objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
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
    Private Sub LineTotalCalc0(ByVal FormUID As String, ByVal RowID As Integer)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objLine = objForm.DataSources.DBDataSources.Item("@MIGTOT1")
        objMatrix.GetLineData(RowID)
        Dim linetotal As Double

        linetotal = CDbl(objMatrix.Columns.Item("9").Cells.Item(RowID).Specific.value) * CDbl(objLine.GetValue("U_unitpric", RowID - 1))

        objLine.SetValue("U_linetot", RowID - 1, linetotal)
        ' MsgBox(CStr(objLine.GetValue("U_linetot", RowID - 1)))
        objMatrix.SetLineData(RowID)
        objForm.Update()

    End Sub
    Private Sub LineTotalCalc(ByVal FormUID As String, ByVal RowID As Integer)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objLine = objForm.DataSources.DBDataSources.Item("@MIGTOT1")
        'objMatrix.GetLineData(RowID)
        Dim linetotal As Double
        linetotal = CDbl(objMatrix.Columns.Item("9").Cells.Item(RowID).Specific.value) * CDbl(objMatrix.Columns.Item("11").Cells.Item(RowID).Specific.value)

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
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Location name")
                Return False
            ElseIf Trim(objForm.Items.Item("6").Specific.Value) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Security name")
                Return False
            ElseIf Trim(objForm.Items.Item("10").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Party details")
                Return False
                'ElseIf Trim(objForm.Items.Item("14").Specific.String) = "" Then
                '    objAddOn.objApplication.SetStatusBarMessage("Please fill up No of packages")
                '    Return False
            ElseIf Trim(objForm.Items.Item("16").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Out time")
                Return False
                'ElseIf Trim(objForm.Items.Item("18").Specific.String) = "" Then
                '    objAddOn.objApplication.SetStatusBarMessage("Please fill up LR Number")
                '    Return False

            ElseIf Trim(objForm.Items.Item("23").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Date")
                Return False
            ElseIf Trim(objForm.Items.Item("27").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Gate Entry No")
                Return False

                'ElseIf Trim(objForm.Items.Item("29").Specific.String) = "" Then
                '    objAddOn.objApplication.SetStatusBarMessage("Please fill up Vehicle Name")
                '    Return False

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
                If objMatrix.Columns.Item("1").Cells.Item(1).Specific.value = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Minimum one Line Item is Required.. ")
                    Return False
                End If
            End If
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage("Please check mandatory fields")
            Return False
        End Try
        Return True
    End Function
    Private Function QtyValidation(ByVal FormUID As String, ByVal RowID As Integer) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objLine = objForm.DataSources.DBDataSources.Item("@MIGTOT1")
        objMatrix.GetLineData(RowID)
        If CDbl(objMatrix.Columns.Item("9").Cells.Item(RowID).Specific.value) > CDbl(objLine.GetValue("U_pendqty", 0)) Then
            objAddOn.objApplication.SetStatusBarMessage("Quantity exceeds pending quantity", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If
        Return True
    End Function
    Private Sub CopyFrom(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objHeader = objForm.DataSources.DBDataSources.Item("@MIGTOT")
        objCombo = objForm.Items.Item("8").Specific
        If objForm.Items.Item("10").Specific.string <> "" Then
            objAddOn.objItemDetails.LoadScreen(Formtype, objCombo.Value, objForm.Items.Item("10").Specific.string, objHeader.GetValue("U_cutdate", 0))
        Else
            objAddOn.objApplication.MessageBox("Please select Party id")
        End If

    End Sub
   
    Private Sub ARInvoiceCopyFrom(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        

    End Sub
    Private Sub GoodsReturnCopyFrom(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)

    End Sub
    Private Sub CopyToStockTransfer(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("3080")
        'Matrix 23; form 940
        Dim StockTransferForm As SAPbouiCOM.Form
        StockTransferForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("940", 1)
        Dim STMatrix As SAPbouiCOM.Matrix
        STMatrix = StockTransferForm.Items.Item("23").Specific

        For i As Integer = 1 To objMatrix.RowCount
            STMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            STMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            STMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
        Next

    End Sub
    Private Sub DeliveryCopyFrom(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)

    End Sub
    Private Sub APCreditMemoCopyFrom(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)

    End Sub
    Private Sub GoodsIssueCopyFrom(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)

    End Sub

End Class


