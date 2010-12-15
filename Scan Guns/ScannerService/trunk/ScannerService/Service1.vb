' NOTE: If you change the class name "ScannerService" here, you must also update the reference to "ScannerService" in App.config.
Imports System.IO


Public Class ScannerService
    Implements IScannerService


#Region "Cycle Counts"

    Public Function CheckBin(ByVal binNo As String) As String Implements IScannerService.CheckBin
        
        CheckBin = "1"
        Using oc As New oracleClient()
            oc.GetDataTable("ICTSCAN1", "SELECT * FROM ICTSCAN1 WHERE ITEM_BIN=:PARM1 AND SCAN_STATUS='P'", False, New Object() {binNo})

            If oc("ICTSCAN1").Rows.Count > 0 Then
                CheckBin = "0"
            End If
        End Using

    End Function

    Public Function GetBin(ByVal itemUpcCode As String) As String Implements IScannerService.GetBin

        GetBin = ""
        Using oc As New oracleClient()
            Dim sql As String = "Select * from ICTITEM1 WHERE (ITEM_CODE=:PARM1 OR ITEM_UPC_CODE=:PARM2 OR ITEM_OPC_CODE=:PARM3 OR ITEM_EAN_CODE=:PARM4) AND ITEM_STATUS='A'"
            oc.GetDataTable("ICTITEM1", sql, False, New Object() {itemUpcCode, itemUpcCode, itemUpcCode, itemUpcCode})
            If oc("ICTITEM1").Rows.Count > 0 Then
                GetBin = oc("ICTITEM1").Rows(0).Item("ITEM_BIN") & ""
            End If
        End Using

    End Function

    Public Function IsBinValid(ByVal bin As String, ByVal whse As String) As String Implements IScannerService.IsBinValid

        IsBinValid = "0"
        Using oc As New oracleClient()
            Dim sql As String = "Select * from ICTPHYX2 WHERE ITEM_BIN=:PARM1 AND WHSE_CODE=:PARM2"
            oc.GetDataTable("ICTPHYX2", sql, False, New Object() {bin, whse})
            If oc("ICTPHYX2").Rows.Count > 0 Then
                IsBinValid = "1"
            End If
        End Using

    End Function

    Function GetScanData(ByVal binNo As String) As DataSet Implements IScannerService.GetScanData

        Dim ds As DataSet = Nothing
        Using oc As New oracleClient()
            Dim sql As String = "SELECT ICTITEM1.*,ICTSCAN2.SCAN_QTY,ICTSTAT2.WHSE_QTY_ON_HAND FROM ICTSCAN1,ICTSCAN2,ICTITEM1,ICTSTAT2 WHERE ICTSCAN1.SCAN_NO=ICTSCAN2.SCAN_NO AND " & _
                " ICTITEM1.ITEM_CODE=ICTSCAN2.ITEM_CODE AND" & _
                " ICTSCAN1.ITEM_BIN=:PARM1 AND ICTSCAN1.SCAN_STATUS='P' AND" & _
                " ICTSTAT2.ITEM_CODE=ICTSCAN2.ITEM_CODE AND ICTSTAT2.WHSE_CODE=ICTSCAN1.WHSE_CODE"
            oc.GetDataTable("Table", sql, False, New Object() {binNo})
            ds = oc.dst.Copy()
        End Using
        Return ds

    End Function

    Public Function GetItemInfo(ByVal itemUpcCode As String, ByVal whseCode As String) As DataSet Implements IScannerService.GetItemInfo

        Dim ds As DataSet = Nothing
        Using oc As New oracleClient()
            Dim sql As String = "Select ICTITEM1.*,ICTSTAT2.WHSE_QTY_ON_HAND from ICTITEM1,ICTSTAT2 WHERE ICTITEM1.ITEM_CODE=ICTSTAT2.ITEM_CODE AND ICTSTAT2.WHSE_CODE=:PARM1 AND ((ICTITEM1.ITEM_UPC_CODE=:PARM2 OR ICTITEM1.ITEM_OPC_CODE=:PARM3 OR ICTITEM1.ITEM_CODE=:PARM4 OR ICTITEM1.ITEM_EAN_CODE=:PARM5) AND ITEM_STATUS='A')"
            oc.GetDataTable("Table", sql, False, New Object() {whseCode, itemUpcCode, itemUpcCode, itemUpcCode, itemUpcCode})
            ds = oc.dst.Copy()
        End Using
        Return ds

    End Function


    Public Function UpdateItemInfo(ByVal dst As DataSet, ByVal binNo As String, ByVal priceCatgy As String, ByVal operId As String, ByVal whseCode As String, ByVal forUpdate As String) As String Implements IScannerService.UpdateItemInfo

        Dim returnStatus = "1" 'update was successful

        Using oc As New oracleClient()
            Dim SCAN_NO As String = ""
            Dim sql As String
            oc.BeginTrans()
            If forUpdate = "1" Then
                sql = "SELECT SCAN_NO FROM ICTSCAN1 WHERE ITEM_BIN=:PARM1 AND SCAN_STATUS='P'"
                SCAN_NO = oc.GetDataValue(sql, New Object() {binNo})
                sql = "Select * from ICTSCAN2 WHERE SCAN_NO=:PARM1"
                oc.GetDataTable("ICTSCAN2", sql, True, New String() {SCAN_NO})

                For Each row As DataRow In dst.Tables(0).Rows
                    If oc("ICTSCAN2").Select("ITEM_CODE='" & row.Item("ITEM_CODE") & "'").Length > 0 Then
                        'update quantities for existing rows
                        Dim trow As DataRow = oc("ICTSCAN2").Select("ITEM_CODE='" & row.Item("ITEM_CODE") & "'")(0)
                        trow.Item("SCAN_QTY") = row.Item("SCAN_QTY")
                    Else
                        'add new rows
                        Dim nrow As DataRow = oc("ICTSCAN2").NewRow()
                        nrow.Item("SCAN_NO") = SCAN_NO
                        nrow.Item("ITEM_CODE") = row.Item("ITEM_CODE")
                        nrow.Item("SCAN_QTY") = row.Item("SCAN_QTY")
                        oc("ICTSCAN2").Rows.Add(nrow)
                    End If
                Next
                'Check for deleted rows
                For Each row As DataRow In oc("ICTSCAN2").Select("1=1")
                    If dst.Tables(0).Select("ITEM_CODE='" & row.Item("ITEM_CODE") & "'").Length = 0 Then
                        oc("ICTSCAN2").Rows.Remove(row)
                    End If
                Next
            Else
                SCAN_NO = Next_Control_No(oc, "ICTSCAN1.SCAN_NO")
                sql = "SELECT * FROM ICTSCAN1 WHERE ROWNUM < 1"
                oc.GetDataTable("ICTSCAN1", sql, True, Nothing)
                Dim nrow As DataRow = oc("ICTSCAN1").NewRow()
                nrow.Item("SCAN_NO") = SCAN_NO
                nrow.Item("WH_OPER_ID") = operId
                nrow.Item("ITEM_BIN") = binNo
                nrow.Item("PRICE_CATGY_CODE") = priceCatgy
                nrow.Item("INIT_DATE") = Now
                nrow.Item("SCAN_STATUS") = "P" 'pending review status
                nrow.Item("WHSE_CODE") = whseCode
                oc("ICTSCAN1").Rows.Add(nrow)

                sql = "Select * from ICTSCAN2 WHERE ROWNUM < 1"
                oc.GetDataTable("ICTSCAN2", sql, True, Nothing)
                For Each row As DataRow In dst.Tables(0).Rows
                    Dim nrow2 As DataRow = oc("ICTSCAN2").NewRow()
                    nrow2.Item("SCAN_NO") = SCAN_NO
                    nrow2.Item("ITEM_CODE") = row.Item("ITEM_CODE")
                    nrow2.Item("SCAN_QTY") = row.Item("SCAN_QTY")
                    oc("ICTSCAN2").Rows.Add(nrow2)
                Next

                oc.Update("ICTSCAN1")
            End If

            oc.Update("ICTSCAN2")
            oc.Commit()
            If oc.errorText <> "" Then
                returnStatus = oc.errorText
            End If
        End Using

        Return returnStatus
    End Function

#End Region


#Region "Receivings"

    Public Function LoadPO(ByVal invPackUPC As String) As DataSet Implements IScannerService.LoadPO
        Dim dst As New DataSet()
        Using oc As New oracleClient()
            Dim poOrderNo As String = oc.GetDataValue("SELECT PO_ORDER_NO FROM ICTPOREC WHERE INV_PACK_UPC=:PARM1 AND ROWNUM < 2", invPackUPC)
            If poOrderNo Is Nothing Then
                Throw New Exception("No PO Located")
            End If
            If poOrderNo.Length = 0 Then
                LoadPO = dst
            Else
                Dim sql As String = ""
                oc.GetDataTable("POTORDR1", "Select * from POTORDR1 WHERE PO_ORDER_NO=:PARM1", False, poOrderNo)

                sql = "SELECT " _
                    & "  PO2.*," _
                    & "  DECODE(NVL(IP1.PO_QTY_MULT,0), 0,PO2.PO_UOM_CONV_FACTOR,IP1.PO_QTY_MULT) SCAN_MULT," _
                    & "  IC1.ITEM_UPC_CODE," _
                    & "  IC1.ITEM_OPC_CODE," _
                    & "  PO2.PO_QTY_OPN - NVL(PS2.SCAN_QTY,0) OPEN_BALANCE " _
                    & "FROM " _
                    & "  POTORDR2 PO2" _
                    & "   JOIN" _
                    & "  ICTITEM1 IC1 ON (IC1.ITEM_CODE=PO2.ITEM_CODE)" _
                    & "   JOIN" _
                    & "  ICTPCAT1 IP1 ON (IC1.PRICE_CATGY_CODE=IP1.PRICE_CATGY_CODE)" _
                    & "   LEFT JOIN" _
                    & "  POTSCAN1 PS1 ON (PS1.STATUS='O' AND PS1.PO_ORDER_NO=PO2.PO_ORDER_NO)" _
                    & "   LEFT JOIN" _
                    & "  POTSCAN2 PS2 ON (PS1.SCAN_NO=PS2.SCAN_NO AND PS2.ITEM_CODE=PO2.ITEM_CODE) " _
                    & "WHERE " _
                    & "  PO2.PO_ORDER_NO=:PARM1 AND" _
                    & "  PO2.ITEM_CODE=IC1.ITEM_CODE"
                oc.GetDataTable("POTORDR2", sql, False, poOrderNo)

                dst = oc.dst.Copy()
                LoadPO = dst
            End If
        End Using
        
    End Function

    Public Function UpdatePO(ByVal WH_OPER_ID As String, ByVal dst As DataSet) As String Implements IScannerService.UpdatePO
        'Take in the dataset from the gun and write to POTSCAN tables with the info in it
        UpdatePO = "1"
        Try
            Using oc As New oracleClient()
                Dim sql As String = "SELECT * FROM POTSCAN1 WHERE ROWNUM < 1"
                oc.GetDataTable("POTSCAN1", sql, True)
                sql = "SELECT * FROM POTSCAN2 WHERE ROWNUM < 1"
                oc.GetDataTable("POTSCAN2", sql, True)

                oc.BeginTrans()

                Dim SCAN_NO = Next_Control_No(oc, "POTSCAN1.SCAN_NO")

                Dim nrow As Data.DataRow = oc("POTSCAN1").NewRow
                nrow.Item("SCAN_NO") = SCAN_NO
                nrow.Item("PO_ORDER_NO") = dst.Tables("POTORDR1").Rows(0).Item("PO_ORDER_NO")
                nrow.Item("INIT_OPER") = WH_OPER_ID
                nrow.Item("INIT_DATE") = Now
                nrow.Item("STATUS") = "O"
                oc("POTSCAN1").Rows.Add(nrow)

                For Each row As Data.DataRow In dst.Tables("POTORDR2").Rows
                    Dim nrow2 As DataRow = oc("POTSCAN2").NewRow
                    nrow2.Item("SCAN_NO") = SCAN_NO
                    nrow2.Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                    nrow2.Item("SCAN_LNO") = row.Item("PO_ORDER_LNO")
                    nrow2.Item("ITEM_CODE") = row.Item("ITEM_CODE")
                    nrow2.Item("SCAN_QTY") = row.Item("SCAN_QTY")
                    oc("POTSCAN2").Rows.Add(nrow2)
                Next

                oc.Update("POTSCAN1")
                oc.Update("POTSCAN2")
                oc.Commit()
            End Using
        Catch ex As Exception
            UpdatePO = ex.Message
        End Try

    End Function

#End Region

#Region "Shared Functions"

    Private Function Next_Control_No(ByVal oc As oracleClient, ByVal CTL_NO_TYPE As String)
        Return oc.ExecuteSF("TAPCTLN2", New String() {"CTL_NO_TYPE_IN", "CTL_NO_LEN_IN", "HOW_MANY_IN", "INIT_OPER_IN"}, New Object() {CTL_NO_TYPE, 10, 1, "scanserv"})
    End Function
    'create or replace FUNCTION "TAPCTLN2" ("CTL_NO_TYPE_IN" VARCHAR2, "CTL_NO_LEN_IN" NUMBER, "HOW_MANY_IN" NUMBER, "INIT_OPER_IN" VARCHAR2) RETURN VARCHAR2 AS
    'BEGIN
    'BEGIN
    'INSERT INTO TATCTLN1 (CTL_NO_TYPE) VALUES (CTL_NO_TYPE_IN);
    'EXCEPTION
    'WHEN OTHERS THEN NULL;
    'END;
    'BEGIN
    ' DECLARE
    '  CURSOR C1 IS SELECT * FROM TATCTLN1 WHERE CTL_NO_TYPE = CTL_NO_TYPE_IN FOR UPDATE;
    '  CTL_NO NUMBER;
    '  CTL_NO_RETURN VARCHAR2(10);
    '  HOW_MANY_WORK NUMBER;
    '  CTL_NO_LEN_WORK NUMBER;
    ' BEGIN
    '  HOW_MANY_WORK := NVL(HOW_MANY_IN,0);
    '  IF NVL(HOW_MANY_WORK,0) = 0 THEN HOW_MANY_WORK := 1; END IF;
    '  CTL_NO_LEN_WORK := NVL(CTL_NO_LEN_IN,0);
    '  IF NVL(CTL_NO_LEN_WORK,0) = 0 THEN CTL_NO_LEN_WORK := 10; END IF;
    '  FOR R1 IN C1 LOOP
    '   IF NVL(R1.CTL_NO_LENGTH,0) <> 0 THEN
    '    CTL_NO_LEN_WORK := R1.CTL_NO_LENGTH;
    '   END IF;
    '   CTL_NO := NVL(R1.CTL_NO_LAST,0) + 1;
    '   IF CTL_NO >= 10 ** CTL_NO_LEN_WORK THEN CTL_NO := 1; END IF;
    '   CTL_NO_RETURN := LPAD(TO_CHAR(CTL_NO),CTL_NO_LEN_WORK,'0');
    '   IF HOW_MANY_WORK > 1 THEN
    '	CTL_NO := CTL_NO + HOW_MANY_WORK -1;
    '    IF CTL_NO >= 10 ** CTL_NO_LEN_WORK THEN CTL_NO := CTL_NO - 10 ** CTL_NO_LEN_WORK + 1; END IF;
    '   END IF;
    '   UPDATE TATCTLN1 SET CTL_NO_LAST = CTL_NO WHERE CURRENT OF C1;
    '   Insert into TATCTLN2 (CTL_NO_TYPE,CTL_NO_LAST,CTL_NO_KEY,HOW_MANY,INIT_DATE,INIT_OPER) VALUES (ctl_no_type_in, ctl_no, ctl_no_return, how_many_work,SYSDATE, init_oper_in);
    '  END LOOP;
    '  RETURN (CTL_NO_RETURN);
    ' END;
    ' END;
    'END "TAPCTLN2";

#End Region

End Class
