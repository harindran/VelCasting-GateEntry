Public Class clsItemDetails
    Public Const Formtype As String = "ItemDetails"
    Dim objForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim IOFormUID As String = ""
    Dim objGrid As SAPbouiCOM.Grid
    Dim strSQL As String = ""
    Public Sub LoadScreen(ByVal CallerFormUID As String, ByVal DocType As String, ByVal CardCode As String, ByVal CutDate As String)
        IOFormUID = CallerFormUID
        objForm = objAddOn.objUIXml.LoadScreenXML("ItemDetails1.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
        objForm.DataSources.DataTables.Add("Items")
        If CallerFormUID.Contains(clsGateOutward.Formtype) Then
            LoadItemDetails(objForm.UniqueID, DocType, Trim(CardCode), CutDate)
        ElseIf CallerFormUID.Contains(clsGateInward.Formtype) Then
            LoadInwardItemDetails(objForm.UniqueID, DocType, Trim(CardCode), CutDate)
        End If

        objForm.Visible = True
    End Sub
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        If pval.BeforeAction Then
            If pval.ItemUID = "101" Then
                LoadDetailsToGT(FormUID)
            End If
        Else
        End If
    End Sub
    Private Sub LoadDetailsToGT(ByVal FormUID As String)
        Dim OutwardForm As SAPbouiCOM.Form
        Dim OutwardMatrix As SAPbouiCOM.Matrix
        Dim DBLine As String = ""
        If IOFormUID.Contains(clsGateOutward.Formtype) Then
            DBLine = "@MIGTOT1"
        ElseIf IOFormUID.Contains(clsGateInward.Formtype) Then
            DBLine = "@MIGTIN1"
        End If
        OutwardForm = objAddOn.objApplication.Forms.GetForm(IOFormUID, 1)
        OutwardMatrix = OutwardForm.Items.Item("36").Specific
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objGrid = objForm.Items.Item("3").Specific
        OutwardForm.DataSources.DBDataSources.Item(DBLine).Clear()
        For i As Integer = objGrid.Rows.Count To 1 Step -1
            If objGrid.DataTable.GetValue("Select", i - 1) = "Y" Then
                With OutwardForm.DataSources.DBDataSources.Item(DBLine)
                    .InsertRecord(0)
                    .SetValue("U_basetype", 0, objGrid.DataTable.GetValue("ObjType", i - 1))
                    .SetValue("U_basenum", 0, objGrid.DataTable.GetValue("DocNum", i - 1))
                    .SetValue("U_basentry", 0, objGrid.DataTable.GetValue("DocEntry", i - 1))
                    .SetValue("U_baseline", 0, CStr(objGrid.DataTable.GetValue("LineNum", i - 1)))
                    .SetValue("U_itemcode", 0, objGrid.DataTable.GetValue("ItemCode", i - 1))
                    .SetValue("U_itemdesc", 0, objGrid.DataTable.GetValue("Dscription", i - 1))
                    '  .SetValue("U_itemdet", 0, objGrid.DataTable.GetValue("Dscription", i - 1))
                    .SetValue("U_itemdet1", 0, objGrid.DataTable.GetValue("Details", i - 1))
                    .SetValue("U_orderqty", 0, objGrid.DataTable.GetValue("Quantity", i - 1))
                    .SetValue("U_pendqty", 0, objGrid.DataTable.GetValue("PendQty", i - 1))
                    .SetValue("U_qty", 0, objGrid.DataTable.GetValue("PendQty", i - 1))
                    .SetValue("U_unitpric", 0, objGrid.DataTable.GetValue("Price", i - 1))
                    .SetValue("U_linetot", 0, CDbl(objGrid.DataTable.GetValue("Quantity", i - 1)) * CDbl(objGrid.DataTable.GetValue("Price", i - 1)))
                End With
            End If
        Next
        OutwardMatrix.LoadFromDataSourceEx()
        OutwardMatrix.AutoResizeColumns()
        objForm.Close()

    End Sub
    Private Function ReturnQueryOutward(ByVal doctype As String) As String
        Return ""
    End Function
    Public Sub LoadItemDetails(ByVal FormUID As String, ByVal DType As String, ByVal CardCode As String, ByVal CutoffDate As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objGrid = objForm.Items.Item("3").Specific
        If objAddOn.HANA Then




            Select Case Trim(DType)
                Case "SI" 'sales invoice
                    strSQL = "SELECT '' AS ""Select"", (SELECT  ""DocNum"" FROM OINV WHERE ""DocEntry"" = T0.""DocEntry"") AS ""DocNum"", T0.""DocEntry"", T0.""LineNum"", " &
                        " T0.""ItemCode"", T0.""Dscription"",IFNULL(T0.""Text"",'') AS ""Details"", T0.""Quantity"", T0.""Price"", T0.""LineTotal"", T0.""ObjType"", " &
                        " T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" = cast (T0.""DocEntry"" as varchar) AND ""U_baseline"" = T0.""LineNum"") AS ""PendQty"" " &
                        " FROM INV1 T0 WHERE T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType""  " &
                        " AND ""U_basentry"" = cast (T0.""DocEntry"" as varchar) AND ""U_baseline"" = T0.""LineNum"") > 0 AND T0.""BaseCard"" = '" & CardCode & "' AND T0.""DocDate"">='" & CutoffDate & "';"
                Case "NR" 'Delivery

                    strSQL = "SELECT '' AS ""Select"", (SELECT  ""DocNum"" FROM ODLN WHERE ""DocEntry"" = T0.""DocEntry"") AS ""DocNum"", T0.""DocEntry"", T0.""LineNum"", " &
                        " T0.""ItemCode"", T0.""Dscription"",IFNULL(T0.""Text"",'') AS ""Details"", T0.""Quantity"", T0.""Price"", T0.""LineTotal"", T0.""ObjType"", " &
                        " T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" = cast (T0.""DocEntry"" as varchar) AND ""U_baseline"" = T0.""LineNum"") AS ""PendQty"" " &
                        " FROM DLN1 T0 WHERE T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType""  " &
                        " AND ""U_basentry"" = cast (T0.""DocEntry"" as varchar) AND ""U_baseline"" = T0.""LineNum"") > 0 AND T0.""BaseCard"" = '" & CardCode & "' AND T0.""DocDate"">='" & CutoffDate & "';"
                Case "MO" 'goods issue

                    strSQL = "SELECT '' AS ""Select"", (SELECT  ""DocNum"" FROM OIGE WHERE ""DocEntry"" = T0.""DocEntry"") AS ""DocNum"", T0.""DocEntry"", T0.""LineNum"", " &
                        " T0.""ItemCode"", T0.""Dscription"",IFNULL(T0.""Text"",'') AS ""Details"", T0.""Quantity"", T0.""Price"", T0.""LineTotal"", T0.""ObjType"", " &
                        " T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" = cast (T0.""DocEntry"" as varchar) AND ""U_baseline"" = T0.""LineNum"") AS ""PendQty"" " &
                        " FROM IGE1 T0 WHERE T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType""  " &
                        " AND ""U_basentry"" = cast (T0.""DocEntry"" as varchar) AND ""U_baseline"" = T0.""LineNum"") > 0  AND T0.""DocDate"">='" & CutoffDate & "';" 'AND T0.""BaseCard"" = '" & CardCode & "';"

                'Case "SR" 'goods return

                '    strSQL = "SELECT '' AS ""Select"", (SELECT  ""DocNum"" FROM ORPD WHERE ""DocEntry"" = T0.""DocEntry"") AS ""DocNum"", T0.""DocEntry"", T0.""LineNum"", " &
                '        " T0.""ItemCode"", T0.""Dscription"",IFNULL(T0.""Text"",'') AS ""Details"", T0.""Quantity"", T0.""Price"", T0.""LineTotal"", T0.""ObjType"", " &
                '        " T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" = T0.""DocEntry"" AND ""U_baseline"" = T0.""LineNum"") AS ""PendQty"" " &
                '        " FROM RPD1 T0 WHERE T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType""  " &
                '        " AND ""U_basentry"" = T0.""DocEntry"" AND ""U_baseline"" = T0.""LineNum"") > 0 AND T0.""BaseCard"" = '" & CardCode & "' AND T0.""DocDate"">='" & CutoffDate & "';"
                Case "SC" 'AP credit memo

                    strSQL = "SELECT '' AS ""Select"", (SELECT  ""DocNum"" FROM ORPC WHERE ""DocEntry"" = T0.""DocEntry"") AS ""DocNum"", T0.""DocEntry"", T0.""LineNum"", " &
                        " T0.""ItemCode"", T0.""Dscription"",IFNULL(T0.""Text"",'') AS ""Details"", T0.""Quantity"", T0.""Price"", T0.""LineTotal"", T0.""ObjType"", " &
                        " T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" = cast (T0.""DocEntry"" as varchar) AND ""U_baseline"" = T0.""LineNum"") AS ""PendQty"" " &
                        " FROM RPC1 T0 WHERE T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType""  " &
                        " AND ""U_basentry"" = cast (T0.""DocEntry"" as varchar) AND ""U_baseline"" = T0.""LineNum"") > 0 AND T0.""BaseCard"" = '" & CardCode & "' AND T0.""DocDate"">='" & CutoffDate & "';"

                Case "SR", "JO", "SO", "RT", "RW", "RJ", "ST", "IU", "JW" ' stock transfer

                    strSQL = "SELECT '' AS ""Select"", (SELECT  ""DocNum"" FROM OWTR WHERE ""DocEntry"" = T0.""DocEntry"") AS ""DocNum"", T0.""DocEntry"", T0.""LineNum"", " &
                        " T0.""ItemCode"", T0.""Dscription"",IFNULL(T0.""Text"",'') AS ""Details"", T0.""Quantity"", T0.""Price"", T0.""LineTotal"", T0.""ObjType"", " &
                        " T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" = cast (T0.""DocEntry"" as varchar) AND ""U_baseline"" = T0.""LineNum"") AS ""PendQty"" " &
                        " FROM WTR1 T0 WHERE T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTOT1"" WHERE ""U_basetype"" = T0.""ObjType""  " &
                        " AND ""U_basentry"" = cast (T0.""DocEntry"" as varchar) AND ""U_baseline"" = T0.""LineNum"") > 0 AND T0.""DocDate"">='" & CutoffDate & "';" 'AND T0.""BaseCard"" = '" & CardCode & "';"
            End Select
        Else
            Select Case Trim(DType)
                Case "SI" 'sales invoice
                    strSQL = "SELECT '' 'Select', (select top 1 DocNum from oinv where DocEntry = T0.DocEntry) DocNum, T0.DocEntry , T0.LineNum , T0.[ItemCode], T0.[Dscription], " &
                " isnull(T0.Text,'') Details, T0.[Quantity],  T0.[Price], T0.[LineTotal], T0.ObjType, " &
                " T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0)  from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) PendQty " &
                " FROM INV1 T0 where T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0) from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) >0 and T0.BaseCard='" & CardCode & "'"

                Case "NR" 'Delivery
                    strSQL = " SELECT '' 'Select', (select top 1 DocNum from ODLN where DocEntry = T0.DocEntry) DocNum, T0.DocEntry , T0.LineNum , T0.[ItemCode], T0.[Dscription], " &
                    " isnull(T0.Text,'') Details, T0.[Quantity],  T0.[Price], T0.[LineTotal], T0.ObjType, " &
                    " T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0)  from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) PendQty  " &
                    " FROM DLN1 T0 where T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0) from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) >0 and T0.BaseCard='" & CardCode & "'"

                Case "MO" 'goods issue
                    strSQL = "SELECT '' 'Select', (select top 1 DocNum from OIGE where DocEntry = T0.DocEntry) DocNum, T0.DocEntry , T0.LineNum , T0.[ItemCode], T0.[Dscription], " &
                 " isnull(T0.Text,'') Details, T0.[Quantity],  T0.[Price], T0.[LineTotal], T0.ObjType, " &
                  " T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0)  from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) PendQty " &
                  " FROM IGE1 T0 where T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0) from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) >0"

                Case "SR" 'goods return
                    strSQL = "SELECT '' 'Select', (select top 1 DocNum from ORPD where DocEntry = T0.DocEntry) DocNum, T0.DocEntry , T0.LineNum , T0.[ItemCode], T0.[Dscription], " &
                   " isnull(T0.Text,'') Details, T0.[Quantity],  T0.[Price], T0.[LineTotal], T0.ObjType, " &
                    " T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0)  from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) PendQty " &
                    " FROM RPD1 T0 where T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0) from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) >0 and T0.BaseCard='" & CardCode & "'"

                Case "SC" 'AP credit memo
                    strSQL = "  SELECT '' 'Select', (select top 1 DocNum from ORPC where DocEntry = T0.DocEntry) DocNum, T0.DocEntry , T0.LineNum , T0.[ItemCode], T0.[Dscription], " &
                    " isnull(T0.Text,'') Details, T0.[Quantity],  T0.[Price], T0.[LineTotal], T0.ObjType, " &
                    "  T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0)  from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) PendQty " &
                    "   FROM RPC1 T0 where T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0) from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) >0 and T0.BaseCard='" & CardCode & "'"


                Case "SR", "JO", "SO", "RT", "RW", "RJ", "ST", "IU", "JW" ' stock transfer
                    strSQL = "SELECT '' 'Select', (select top 1 DocNum from OWTR where DocEntry = T0.DocEntry) DocNum, T0.DocEntry , T0.LineNum , T0.[ItemCode], T0.[Dscription], " &
                    " isnull(T0.Text,'') Details, T0.[Quantity],  T0.[Price], T0.[LineTotal], T0.ObjType, " &
                    " T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0)  from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) PendQty " &
                    " FROM WTR1 T0 where T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0) from [@MIGTOT1] where U_basetype = T0.objtype and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) >0"

            End Select
        End If
        objForm.DataSources.DataTables.Item("Items").ExecuteQuery(strSQL)
        objGrid.DataTable = objForm.DataSources.DataTables.Item("Items")
        objGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

    End Sub
    Public Sub LoadInwardItemDetails(ByVal FormUID As String, ByVal DType As String, ByVal CardCode As String, ByVal CutoffDate As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objGrid = objForm.Items.Item("3").Specific

        If objAddOn.HANA Then
            Select Case Trim(DType)
                Case "PO", "GR"
                    strSQL = "SELECT '' AS ""Select"", (SELECT ""DocNum"" FROM OPOR WHERE ""DocEntry"" = T0.""DocEntry"") AS ""DocNum"", T0.""DocEntry"", T0.""LineNum"", " &
            " T0.""ItemCode"", T0.""Dscription"",IFNULL(T0.""Text"",'') AS ""Details"", T0.""Quantity"", T0.""Price"", T0.""LineTotal"", T0.""ObjType"", " &
            " T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTIN1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" = Cast(T0.""DocEntry"" as Varchar) AND ""U_baseline"" = T0.""LineNum"") AS ""PendQty"" " &
            " FROM POR1 T0 WHERE T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTIN1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry""  =  Cast(T0.""DocEntry"" as Varchar) " &
            " AND ""U_baseline"" = T0.""LineNum"") > 0 AND T0.""LineStatus"" ='O' AND T0.""BaseCard"" = '" & CardCode & "';"

                Case "SR", "IN"
                    strSQL = "SELECT '' AS ""Select"", (SELECT ""DocNum"" FROM OINV WHERE ""DocEntry"" = T0.""DocEntry"") AS ""DocNum"", T0.""DocEntry"", T0.""LineNum"", " &
            " T0.""ItemCode"", T0.""Dscription"",IFNULL(T0.""Text"",'') AS ""Details"", T0.""Quantity"", T0.""Price"", T0.""LineTotal"", T0.""ObjType"", " &
            " T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTIN1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" =  Cast(T0.""DocEntry"" as Varchar) AND ""U_baseline"" = T0.""LineNum"") AS ""PendQty"" " &
            " FROM INV1 T0 WHERE T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTIN1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" = T0.""DocEntry"" " &
            " AND ""U_baseline"" = T0.""LineNum"") > 0 AND T0.""BaseCard"" = '" & CardCode & "' AND T0.""DocDate"">='" & CutoffDate & "';"
                Case "DR", "WI"
                    strSQL = "SELECT '' AS ""Select"", (SELECT ""DocNum"" FROM ODLN WHERE ""DocEntry"" = T0.""DocEntry"") AS ""DocNum"", T0.""DocEntry"", T0.""LineNum"", " &
            " T0.""ItemCode"", T0.""Dscription"",IFNULL(T0.""Text"",'') AS ""Details"", T0.""Quantity"", T0.""Price"", T0.""LineTotal"", T0.""ObjType"", " &
            " T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTIN1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" = Cast(T0.""DocEntry"" as Varchar) AND ""U_baseline"" = T0.""LineNum"") AS ""PendQty"" " &
            " FROM DLN1 T0 WHERE T0.""Quantity"" - (SELECT IFNULL(SUM(IFNULL(""U_qty"", 0)), 0) FROM ""@MIGTIN1"" WHERE ""U_basetype"" = T0.""ObjType"" AND ""U_basentry"" = T0.""DocEntry"" " &
            " AND ""U_baseline"" = T0.""LineNum"") > 0 AND T0.""BaseCard"" = '" & CardCode & "' AND T0.""DocDate"">='" & CutoffDate & "';"
                Case "DC", "JR", "SR", "WM", "RW", "ST", "SO", "CP", "JO", "JW", "MI", "HR" ' Cash Purchase
                    strSQL = "SELECT '' as ""Select"", '0' as ""DocNum"",'0' as ""DocEntry"" ,0 as ""LineNum"",""ItemCode"", ""ItemName""  as  ""Dscription"",""UserText""  as ""Details"",0 as ""Quantity"", 0 as ""Price"", 0 as ""LineTotal"",'4' as ""ObjType"", 0 as ""PendQty"" FROM OITM;"
            End Select

        Else


            Select Case Trim(DType)
                Case "PO", "GR" 'Purchase Order
                    strSQL = "SELECT '' 'Select', (select top 1 DocNum from OPOR where DocEntry = T0.DocEntry) DocNum, T0.DocEntry , T0.LineNum , T0.[ItemCode], T0.[Dscription], " &
                    " isnull(T0.Text,'') Details, T0.[Quantity],  T0.[Price], T0.[LineTotal], T0.ObjType, " &
                    " T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0)  from [@MIGTIN1] where U_basetype = T0.ObjType and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) PendQty " &
                    " FROM POR1 T0 where T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0) from [@MIGTIN1] " &
                    " where U_basetype = T0.ObjType and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) >0 and T0.LineStatus ='O' and T0.BaseCard='" & CardCode & "'"


                Case "SR", "IN" 'Sales Return with Invoice
                    strSQL = "SELECT '' 'Select', (select top 1 DocNum from OINV where DocEntry = T0.DocEntry) DocNum, T0.DocEntry , T0.LineNum , T0.[ItemCode], T0.[Dscription], " &
                " isnull(T0.Text,'') Details, T0.[Quantity],  T0.[Price], T0.[LineTotal], T0.ObjType, " &
                "  T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0)  from [@MIGTIN1] where U_basetype = T0.ObjType and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) PendQty " &
                " FROM INV1 T0 where T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0) from [@MIGTIN1] " &
                " where U_basetype = T0.ObjType and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) >0 and T0.BaseCard='" & CardCode & "'"


                Case "DR", "WI"
                    strSQL = "SELECT '' 'Select', (select top 1 DocNum from ODLN where DocEntry = T0.DocEntry) DocNum, T0.DocEntry , T0.LineNum , T0.[ItemCode], T0.[Dscription], " &
                 " isnull(T0.Text,'') Details, T0.[Quantity],  T0.[Price], T0.[LineTotal], T0.ObjType, " &
                 "  T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0)  from [@MIGTIN1] where U_basetype = T0.ObjType and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) PendQty " &
                 " FROM DLN1 T0 where T0.Quantity- (select isnull(sum(isnull(U_qty,0)),0) from [@MIGTIN1] " &
                 " where U_basetype = T0.ObjType and U_basentry=Cast(T0.DocEntry as varchar) and U_baseline=T0.LineNum) >0 and T0.BaseCard='" & CardCode & "'"
                Case "DC", "JR", "SR", "WM", "RW", "ST", "SO", "CP", "JO", "JW", "MI", "HR" ' Cash Purchase
                    strSQL = "SELECT '' 'Select', '' DocNum,'' DocEntry,'' LineNum,ItemCode,ItemName Dscription,UserText Details,0 Quantity,0 as Price, 0 as LineTotal,'4' ObjType, 0 PendQty from OITM"

            End Select
        End If
        objForm.DataSources.DataTables.Item("Items").ExecuteQuery(strSQL)
        objGrid.DataTable = objForm.DataSources.DataTables.Item("Items")
        objGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
    End Sub
End Class
