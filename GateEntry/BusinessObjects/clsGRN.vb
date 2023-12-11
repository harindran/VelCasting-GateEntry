Public Class clsGRN
    Public Const formtype As String = "143"
    Dim objForm As SAPbouiCOM.Form
    Dim strSQL As String
    Dim objRS As SAPbobsCOM.Recordset
    Dim GEDocEntry As String
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try


            If pval.BeforeAction Then
                objForm = objAddOn.objApplication.Forms.Item(FormUID)
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        If pval.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                            If Not GEVerfication(FormUID) Then
                                BubbleEvent = False

                                Exit Sub
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If pval.ItemUID = "TVer" Then
                            BubbleEvent = False
                        End If
                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        CreateButton(FormUID)
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pval.ItemUID = "Verify" Then
                            Verify(FormUID)
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
                    If GEVerfication(BusinessObjectInfo.FormUID) Then

                    Else
                        BubbleEvent = False
                    End If
                End If

            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End Try
        Else
            CloseGateEntry()
        End If
    End Sub
    Private Sub CloseGateEntry()

        If objAddOn.HANA Then
            strSQL = "UPDATE ""@MIGTIN"" set ""Status""='C' where ""DocEntry"" ='" & GEDocEntry & "'"
        Else

            strSQL = "Update [@MIGTIN] set Status='C' where DocEntry='" & GEDocEntry & "'"
        End If
        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRS.DoQuery(strSQL)
        objRS = Nothing
    End Sub
    Public Sub CreateButton(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objButton As SAPbouiCOM.Button
        Dim objItem As SAPbouiCOM.Item
        Try
            objButton = objForm.Items.Item("Verify").Specific
        Catch ex As Exception
            objItem = objForm.Items.Add("Verify", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Left = objForm.Items.Item("2").Left + 100
            objItem.Width = objForm.Items.Item("2").Width
            objItem.Top = objForm.Items.Item("2").Top
            objItem.Height = objForm.Items.Item("2").Height
            objButton = objItem.Specific
            objButton.Caption = "Verify GE"
        End Try
        Dim objText As SAPbouiCOM.EditText
        Try
            objText = objForm.Items.Item("TVer").Specific
        Catch ex As Exception


            objItem = objForm.Items.Add("TVer", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Left = objForm.Items.Item("2").Left + 200
            objItem.Width = objForm.Items.Item("2").Width
            objItem.Top = objForm.Items.Item("2").Top
            objItem.Height = objForm.Items.Item("2").Height
            objText = objItem.Specific
            objText.DataBind.SetBound(True, "OPDN", "U_gever")
        End Try

        'Dim objLink As SAPbouiCOM.LinkedButton
        'Try
        '    objLink = objForm.Items.Item("gelink").Specific
        'Catch ex As Exception


        '    objItem = objForm.Items.Add("gelink", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
        '    objItem.Left = objForm.Items.Item("2").Left + 180
        '    objItem.Width = 20
        '    objItem.Top = objForm.Items.Item("2").Top
        '    objItem.Height = objForm.Items.Item("2").Height
        '    objLink = objItem.Specific
        '    'objText.DataBind.SetBound(True, "OPDN", "U_gever")
        '    objLink.LinkedObjectType = "MIGTIN"
        '    objItem.LinkTo = "TVer"
        'End Try

    End Sub
    Private Function Verify(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objMatrix As SAPbouiCOM.Matrix
        objMatrix = objForm.Items.Item("38").Specific
        For i As Integer = 1 To objMatrix.RowCount
            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                Dim BaseEntry As Long = CLng(objMatrix.Columns.Item("45").Cells.Item(i).Specific.String)
                Dim BaseLine As Integer = CInt(objMatrix.Columns.Item("46").Cells.Item(i).Specific.String)
                Dim BaseQty As Double = CDbl(objMatrix.Columns.Item("11").Cells.Item(i).Specific.String)
                
                    If objAddOn.HANA Then
                    strSQL = "SELECT T0.""DocEntry"" FROM ""@MIGTIN1"" T0 WHERE T0.""U_basentry"" = " & BaseEntry & " AND  T0.""U_baseline"" = " & BaseLine & " AND " & _
                        " T0.""U_qty"" = " & BaseQty & " AND T0.""U_basetype"" = 22;"

                    Else
                    strSQL = "select T0.DocEntry from [@MIGTIN1] T0 where T0.U_basentry = " & BaseEntry & " and T0.U_baseline= " & BaseLine & " and" & _
                   "  T0.U_qty = " & BaseQty & " and T0.U_basetype=22"
                    End If
                    GEDocEntry = objAddOn.objGenFunc.getSingleValue(strSQL)
                If GEDocEntry = "" Then
                    objForm.Items.Item("TVer").Specific.String = GEDocEntry
                    objAddOn.objApplication.MessageBox("Please check the quantity at line : " & CStr(i))

                    ' objForm.DataSources.DBDataSources.Item("OPDN").SetValue("U_gever", 0, "Open")
                    Return False
                End If

            End If
        Next
        'objAddOn.objApplication.MessageBox("All ok")
        objForm.Items.Item("TVer").Specific.String = CStr(GEDocEntry)
        'objForm.DataSources.DBDataSources.Item("OPDN").SetValue("U_gever", 0, "Verified")

        Return True
    End Function
    Private Function GEVerfication(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        GEDocEntry = Trim(CStr(objForm.DataSources.DBDataSources.Item("OPDN").GetValue("U_gever", 0)))
        If GEDocEntry = "" Then
            objAddOn.objApplication.SetStatusBarMessage("Please verify with Gate Entry")
            Return False
        End If

        Return True
    End Function
End Class
