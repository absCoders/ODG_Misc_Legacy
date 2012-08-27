Public Class DelJobService

    Private Shared clsDataConnection As New Oracle.DataAccess.Client.OracleConnection
    'Public Shared oraCmd As New OracleCommand
    'Public Shared oraSP As New OracleCommand    ' Used to execute Stored Procedures
    'Public Shared T As OracleTransaction
    Private Shared clsDataAdapter As New Oracle.DataAccess.Client.OracleDataAdapter
    Private Shared tablePrefix As String = String.Empty

    ''' <summary>
    ''' Structure to hold validation errors
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DelJobValidationError

        Private sField As String
        Private sMessage As String

        Public Sub New(ByVal field As String, ByVal message As String)
            sField = field
            sMessage = message
        End Sub

        Public ReadOnly Property Field() As String
            Get
                Return sField
            End Get
        End Property

        Public ReadOnly Property Message()
            Get
                Return sMessage
            End Get
        End Property
    End Class

    ''' <summary>
    ''' DETJOBM1
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure JobValidatorFields
        Dim DelJobId As Int16
        Dim JobNumber As String
        Dim SubmittedOn As Date
        Dim RedoOfJobNumber As String
        Dim ReasonCodeRedo As String
        Dim PractitionerID As Int16
        Dim AccountNumber As String
        Dim ShipToID As String
        Dim PatientId As Int16
        Dim PatientFirstName As String
        Dim PatientLastName As String
        Dim SaveNewPatient As Boolean
        Dim Tray As String
        Dim JobType As String 'FINISHED
        Dim WhichLenses As String ' LENS_ORDER
        Dim Brand As String
        Dim LensDesign As String
        Dim LensDesigner As String
        Dim ColorType As String
        Dim Material As String
        Dim Color As String ' COLOR_CODE_WEB
        Dim Tint As String
        Dim TintColor As String
        Dim TintPercent As Int16
        Dim UseWorkingDistances As Int16
        Dim WorkingDistancesDistance As Int16
        Dim WorkingDistancesViewingAngle As String
        Dim WorkingDistancesActivity As String
        Dim WorkingDistancesNear As Int16
        Dim WorkingDistancesFar As Int16
        Dim ArCoating As String
        Dim BacksideOnly As String ' AR_BACKSIDE_ONLY
        Dim MirrorCoating As String
        Dim MirrorCoatingColor As String
        Dim Polish As String
        Dim TintableHardcoat As String

        Dim DistanceRxRightSph As Decimal
        Dim DistanceRxRightCyl As Decimal
        Dim DistanceRxRightAxis As Decimal

        Dim DistanceRxLeftSph As Decimal
        Dim DistanceRxLeftCyl As Decimal
        Dim DistanceRxLeftAxis As Decimal

        Dim NearRxRightAdd As Decimal
        Dim NearRxLeftAdd As Decimal

        Dim IncludePrism As Boolean
        Dim PrismRightIn As Decimal
        Dim PrismRightInAxis As String
        Dim PrismRightUp As Decimal
        Dim PrismRightUpAxis As String
        Dim PrismLeftIn As Decimal
        Dim PrismLeftInAxis As String
        Dim PrismLeftUp As Decimal
        Dim PrismLeftUpAxis As String

        Dim VertexPantoRightFittingVertex As Decimal
        Dim VertexPantoRightRefractiveVertex As Decimal
        Dim VertexPantoRightPantoTilt As Decimal
        Dim VertexPantoRightWrapAngle As Decimal

        Dim VertexPantoLeftFittingVertex As Decimal
        Dim VertexPantoLeftRefractiveVertex As Decimal
        Dim VertexPantoLeftPantoTilt As Decimal
        Dim VertexPantoLeftWrapAngle As Decimal

        Dim FittingMeasurementsRightMonocularPD As Decimal
        Dim FittingMeasurementsRightFittingHeight As Decimal
        Dim FittingMeasurementsLeftMonocularPD As Decimal
        Dim FittingMeasurementsLeftFittingHeight As Decimal

        Dim FrameDataA As Decimal
        Dim FrameDataB As Decimal
        Dim FrameDataDbl As String
        Dim FrameDataED As Decimal

        Dim FrameType As String
        Dim EdgeThickness As String
        Dim BaseCurve As Decimal

        Dim SpecialInstructions As String
        Dim Finishedwork As String
        Dim FrameManufacturer As String
        Dim FrameModelNumber As String
        Dim FrameColor As String
        Dim ShapePattern As String
        Dim FrameSize As String
        Dim TrackingNumber As String
        Dim JobStatusCode As String 'STATUS_TYPE_WEB
        Dim JobStatusText As String ' STATUS_DESC_WEB
        Dim IsJobStatusDelayed As Int16
        Dim JobStatusOn As Date ' Init_Date
        Dim InvoicedOn As Date
        Dim InvoiceNumber As String ' InvNo
        Dim ShipViaDescriptionWeb As String 'SHIP_VIA_DESC_WEB
        Dim CarrierCode As String
        Dim CarrierTrackingNumber As String 'SHIP_REF
        Dim AddedOn As Date
        Dim UpdatedOn As Date
        Dim UpdatedBy As String

        '****************************************

        'Dim CustName As String
        'Dim CustShipToName As String
        'Dim OrdrCustPo As String
        'Dim OrdrSource As String
        'Dim ShipViaCode As String
        'Dim ColorCode As String
        'Dim CorridorLength As String
        'Dim Finished As String

        'Dim FrameAWidth As Decimal
        'Dim FrameBHeight As Decimal
        'Dim FrameDblBridge As Decimal
        'Dim FrameEdDiagonal As Decimal
        'Dim JobStatus As String
        'Dim InitOper As String
        'Dim LastOper As String
        'Dim LastDate As Date
        'Dim JobTypeCode As String
        'Dim LensDesignerCode As String
        'Dim ListPrice As Decimal
        'Dim CommentInv As String
        'Dim PatientName As String
        'Dim RxPrism As String
        'Dim OrdrNo As String
        'Dim OrdrMessageLab As String
        'Dim JobCancelledReason As String
        'Dim TraceFrom As String
        'Dim Fpc As String
        'Dim TrcInd As String
        'Dim TrcQueued As Date
        'Dim TrcCompleted As Date
        'Dim LdsInd As String
        'Dim LdsQueued As Date
        'Dim LdsCompleted As Date
        'Dim LmsInd As String
        'Dim LmsQueued As Date
        'Dim LmsCompleted As Date
        'Dim TktInd As String
        'Dim TktQueued As Date
        'Dim TktCompleted As Date
        'Dim RushOrder As String
        'Dim UseThinningPrism As String
        'Dim JobScanVerified As String
        'Dim BalanceLens As String
        'Dim BlankSelection As String
        'Dim SuppressCrib As String
        'Dim Edging As String
        'Dim NoFreight As String
        'Dim JobError As String
        'Dim InvDate As Date
        'Dim JobNoReplFrom As String
        'Dim InvFreight As Decimal
        'Dim InvMiscChg As Decimal
        'Dim InvStax As Decimal
        'Dim InvTotalAmount As Decimal
        'Dim OpsYyyypp As String
        'Dim StaxCode As String
        'Dim TermCode As String
        'Dim UseDiscPct As String
        'Dim InvDiscPct As Decimal
        'Dim TktPrintCount As Int16
        'Dim InvSales As Decimal
        'Dim JobNoTraceFrom As String
        'Dim ReasonCodeCanc As String
        'Dim ReasonCodeDisc As String
        'Dim Nwd As Decimal
        'Dim Fwd As Decimal
        'Dim Nwa As Decimal
        'Dim Fwa As Decimal
        'Dim FittingVertex As Decimal
        'Dim RefractiveVertex As Decimal
        'Dim PantoscopicTilt As Decimal
        'Dim PanoramicAngle As Decimal
        'Dim JobMustShipToday As String
        'Dim WorkingDistance As Decimal
        'Dim ViewingAngle As Decimal
        'Dim Activity As String
        'Dim LabHold As String
        'Dim OrdrCallerName As String
        'Dim JobHoldLab As String
        'Dim JobHoldLabReason As String
        'Dim JobHoldInv As String
        'Dim JobHoldInvReason As String
        'Dim BinNoDel As Int16
        'Dim JobScanVerifiedExc As String
        'Dim OpcCodeSvR As String
        'Dim OpcCodeSvL As String
        'Dim JobChecked As String
        'Dim JobCheckedOper As String
        'Dim JobCheckedDate As Date
        'Dim JobInQueue As String
        'Dim JobRequiresReview As String
        'Dim JobQueuedDate As Date
        'Dim CouponCode As String
        'Dim JobInspctSup As String
        'Dim JobReportedInd As String
        'Dim JobReportedXno As String
        'Dim Nod As Decimal
        'Dim JobReportedDate As Date
        'Dim WrapEdge As String
        'Dim CustomFrameFile As String
        'Dim CustomFrameNew As String
        'Dim PatternNo As String
        'Dim CustomFrameNo As String
        'Dim DatabaseTraceExecuted As String
        'Dim JobPtPrinted As String
        'Dim CustClinit As String
        'Dim WrapEdgeSport As String
        'Dim JobLockPricing As String
        'Dim ShapeToBeModified As String
        'Dim FogFree As String
        'Dim LensDesignRndsegBlnd As Decimal
        'Dim JobNoBacksideCoat As String
        'Dim JobCaseHeld As String
        'Dim JobFrameHeld As String
        'Dim BinNo As Int16
        'Dim JobInspctSupDesc As String
        'Dim PromoCode As String
        'Dim FuelSurchargeInv As Decimal
    End Structure

    Public Structure OrderValidatorFields
        Dim OrderNumer As String

        Dim CustomerNameBillTo As String
        Dim CustomerAddressBillTo As String
        Dim CustomerCityBillTo As String
        Dim CustomerStateBillTo As String
        Dim CustomerZipcodeBillTo As String

        Dim CustomerNameShipTo As String
        Dim CustomerAddressShipTo As String
        Dim CustomerCityShipTo As String
        Dim CustomerStateShipTo As String
        Dim CustomerZipcodeShipTo As String



    End Structure

    Dim job As JobValidatorFields

#Region "Class Instantiation"

    Public Sub New()

    End Sub

    Public Sub New(ByVal DataConnection As Oracle.DataAccess.Client.OracleConnection, _
                   ByVal DataAdapter As Oracle.DataAccess.Client.OracleDataAdapter)

        clsDataConnection = DataConnection
        clsDataAdapter = DataAdapter
    End Sub

#End Region


#Region "Class Properties"

    ''' <summary>
    ''' Get / Set Data Connection 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DataConnection() As Oracle.DataAccess.Client.OracleConnection
        Get
            Return clsDataConnection
        End Get
        Set(ByVal value As Oracle.DataAccess.Client.OracleConnection)
            clsDataConnection = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Data Adapter
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DataAdapter() As Oracle.DataAccess.Client.OracleDataAdapter
        Get
            Return clsDataAdapter
        End Get
        Set(ByVal value As Oracle.DataAccess.Client.OracleDataAdapter)
            clsDataAdapter = value
        End Set
    End Property

#End Region

#Region "Data Access"

    ''' <summary>
    ''' Returns a datatable for the provided query
    ''' If an error occurs Nothing is returned
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetDataTable(ByVal sql As String) As DataTable
        Try
            If clsDataConnection Is Nothing OrElse clsDataAdapter Is Nothing Then
                Return Nothing
            End If

            clsDataAdapter.SelectCommand = clsDataConnection.CreateCommand
            clsDataAdapter.SelectCommand.CommandText = sql
            Dim tbl As DataTable = New DataTable
            clsDataAdapter.Fill(tbl)
            Return tbl

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Returns a datarow for the provided query
    ''' If more than one row is returned the first row in the results is returned
    ''' If an error occurs or not result is found, Nothing is returned
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetDataRow(ByVal sql As String) As DataRow
        Try
            If clsDataConnection Is Nothing OrElse clsDataAdapter Is Nothing Then
                Return Nothing
            End If

            clsDataAdapter.SelectCommand = clsDataConnection.CreateCommand
            clsDataAdapter.SelectCommand.CommandText = sql
            Dim tbl As DataTable = New DataTable
            clsDataAdapter.Fill(tbl)
            If tbl.Rows.Count > 0 Then
                Return tbl.Rows(0)
            Else
                Return Nothing
            End If

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Returns a value from a query
    ''' Query should be an aggregate query and return one row
    ''' If more than one row is retruned then the valur in the firyt row is returned
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetDataValue(ByVal sql As String) As Object
        Try
            If clsDataConnection Is Nothing OrElse clsDataAdapter Is Nothing Then
                Return Nothing
            End If

            clsDataAdapter.SelectCommand = clsDataConnection.CreateCommand
            clsDataAdapter.SelectCommand.CommandText = sql
            Dim tbl As DataTable = New DataTable
            clsDataAdapter.Fill(tbl)

            If tbl IsNot Nothing AndAlso tbl.Rows.Count > 0 Then
                Return tbl.Rows(0).Item(0)
            Else
                Return Nothing
            End If

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Returns the value of a field in a datatable
    ''' </summary>
    ''' <param name="dt">Datatable to search</param>
    ''' <param name="whereClause">Where clause to limit datarows</param>
    ''' <param name="returnFieldname">Name of field to extract data.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function SelectValue(ByRef dt As DataTable, ByVal whereClause As String, ByVal returnFieldname As String) As Object

        If dt Is Nothing Then Return Nothing
        If dt.Rows.Count = 0 Then Return Nothing
        If Not dt.Columns.Contains(returnFieldname) Then Return Nothing

        Try
            If dt.Select(whereClause).Length > 0 Then
                Return dt.Select(whereClause)(0).Item(returnFieldname)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

#End Region

#Region " GetXXX Methods for Data Entry Dropdowns/Options "

    ''' <summary>
    ''' Get Lens Designer Codes
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetLensDesignerCodes() As DataTable
        Return GetLensDesignerCodes(False)
    End Function

    ''' <summary>
    ''' Get Lens Designer Codes
    ''' </summary>
    ''' <param name="WebEnabledOnly">Boolean indicating if only return Designers avaliable for the web</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetLensDesignerCodes(ByVal WebEnabledOnly As Boolean) As DataTable

        Dim TSql As String = "SELECT * FROM " & tablePrefix & "DETDSGN0 WHERE LENS_DESIGNER_CODE <> 'ODG' "

        If WebEnabledOnly Then
            TSql &= " AND (SELECT COUNT(*) FROM " & tablePrefix & "DETDSGN1 WHERE LENS_DESIGNER_CODE = " & tablePrefix & "DETDSGN0.LENS_DESIGNER_CODE AND LENS_DESIGN_ACTIVATE_FOR_WEB='1') > 1"
        End If

        TSql &= " ORDER BY LENS_DESIGNER_NAME"

        Return GetDataTable(TSql)

    End Function

    ''' <summary>
    ''' Get Lens Designs
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetLensDesigns() As DataTable

        Return GetLensDesigns(String.Empty)

    End Function

    ''' <summary>
    ''' Get Lens Designs
    ''' </summary>
    ''' <param name="lensDesignerCode">Designs for a specific Lens Designer</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetLensDesigns(ByVal lensDesignerCode As String) As DataTable

        Dim TSql As String = "SELECT * FROM " & tablePrefix & "DETDSGN1"

        If lensDesignerCode.Length > 0 Then TSql &= " WHERE LENS_DESIGNER_CODE = '{0}'"

        Return GetDataTable(String.Format(TSql, lensDesignerCode))

    End Function

    ''' <summary>
    ''' Return Materials
    ''' </summary>
    ''' <param name="lensDesignerCode">Lens Designer Code</param>
    ''' <param name="lensDesignCode">Lens Design code</param>
    ''' <param name="colorTypeCode">Color Type Code</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetMaterials(ByVal lensDesignerCode As String, ByVal lensDesignCode As String, ByVal colorTypeCode As String) As DataTable

        Dim TSql As String = _
         "SELECT " & _
         " DISTINCT " & tablePrefix & "DETMATL1.MATL_CODE, " & _
         " " & tablePrefix & "DETMATL1.MATL_DESC " & _
         " FROM " & tablePrefix & "DETMATL1 " & _
         " JOIN " & tablePrefix & "DETDSGN3 " & _
         " ON (" & tablePrefix & "DETMATL1.MATL_CODE = " & tablePrefix & "DETDSGN3.MATL_CODE) " & _
         " JOIN " & tablePrefix & "DETCOLR1 " & _
         " ON (" & tablePrefix & "DETDSGN3.COLOR_CODE = " & tablePrefix & "DETCOLR1.COLOR_CODE) " & _
         " JOIN " & tablePrefix & "DETCOLR2 " & _
         " ON (" & tablePrefix & "DETCOLR1.COLOR_CODE_WEB = " & tablePrefix & "DETCOLR2.COLOR_CODE_WEB)" & _
         " WHERE " & tablePrefix & "DETDSGN3.LENS_DESIGN_CODE = '{0}' " & _
                  " AND " & tablePrefix & "DETMATL1.MATL_STATUS = 'A' " & _
                  " AND " & tablePrefix & "DETDSGN3.ACTIVE = '1'" & _
         " AND CASE " & tablePrefix & "DETDSGN3.COLOR_CODE " & _
         "      WHEN 'CLEAR' THEN 'C' " & _
         "      ELSE CASE NVL(" & tablePrefix & "DETCOLR1.POLARIZED,'0') " & _
         "              WHEN '1' THEN 'P' " & _
         "              ELSE 'T' END END = '{1}' " & _
         " ORDER BY MATL_DESC "

        Return GetDataTable(String.Format(TSql, lensDesignCode, colorTypeCode))

    End Function

    ''' <summary>
    ''' Return Materials
    ''' </summary>
    ''' <param name="lensDesignerCode">Lens Designer Code</param>
    ''' <param name="lensDesignCode">Lens Design code</param>
    ''' <param name="colorTypeCode">Color Type Code</param>
    ''' <param name="colorCode">Color Code</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetMaterials(ByVal lensDesignerCode As String, ByVal lensDesignCode As String, ByVal colorTypeCode As String, ByVal colorCode As String) As DataTable

        Dim TSql As String = _
         "SELECT " & _
         " DISTINCT " & tablePrefix & "DETMATL1.MATL_CODE, " & _
         " " & tablePrefix & "DETMATL1.MATL_DESC " & _
         " FROM " & tablePrefix & "DETMATL1 " & _
         " JOIN " & tablePrefix & "DETDSGN3 " & _
         " ON (" & tablePrefix & "DETMATL1.MATL_CODE = " & tablePrefix & "DETDSGN3.MATL_CODE) " & _
         " JOIN " & tablePrefix & "DETCOLR1 " & _
         " ON (" & tablePrefix & "DETDSGN3.COLOR_CODE = " & tablePrefix & "DETCOLR1.COLOR_CODE) " & _
         " JOIN " & tablePrefix & "DETCOLR2 " & _
         " ON (" & tablePrefix & "DETCOLR1.COLOR_CODE_WEB = " & tablePrefix & "DETCOLR2.COLOR_CODE_WEB) " & _
         " WHERE " & tablePrefix & "DETDSGN3.LENS_DESIGN_CODE = '{0}' " & _
                  " AND " & tablePrefix & "DETMATL1.MATL_STATUS = 'A' " & _
                  " AND " & tablePrefix & "DETDSGN3.ACTIVE = '1' " & _
         " AND CASE " & tablePrefix & "DETDSGN3.COLOR_CODE " & _
         "      WHEN 'CLEAR' THEN 'C' " & _
         "      ELSE CASE NVL(" & tablePrefix & "DETCOLR1.POLARIZED,'0') " & _
         "              WHEN '1' THEN 'P' " & _
         "              ELSE 'T' END END = '{1}' " & _
         " AND " & tablePrefix & "DETCOLR2.COLOR_CODE_WEB = '{2}' " & _
         " ORDER BY MATL_DESC "

        Return GetDataTable(String.Format(TSql, lensDesignCode, colorTypeCode, colorCode))

    End Function

    ''' <summary>
    ''' Return Color List
    ''' </summary>
    ''' <param name="lensDesignerCode">Lens Designer Code</param>
    ''' <param name="lensDesignCode">Lens Design Code</param>
    ''' <param name="colorTypeCode">Color Type Code</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetColors(ByVal lensDesignerCode As String, ByVal lensDesignCode As String, ByVal colorTypeCode As String) As DataTable

        ' Return Color list for Brand + LensDesign + ColorType
        Dim TSql As String = _
        "SELECT " & _
        " DISTINCT " & tablePrefix & "DETCOLR2.COLOR_CODE_WEB AS COLOR_CODE, " & _
        " " & tablePrefix & "DETCOLR2.COLOR_DESC_WEB AS COLOR_DESC" & _
        " FROM " & tablePrefix & "DETDSGN3 " & _
        " JOIN " & tablePrefix & "DETCOLR1 " & _
        " ON (" & tablePrefix & "DETDSGN3.COLOR_CODE = " & tablePrefix & "DETCOLR1.COLOR_CODE) " & _
        " JOIN " & tablePrefix & "DETCOLR2 " & _
        " ON (" & tablePrefix & "DETCOLR1.COLOR_CODE_WEB = " & tablePrefix & "DETCOLR2.COLOR_CODE_WEB) " & _
        " WHERE " & tablePrefix & "DETDSGN3.LENS_DESIGN_CODE = '{0}' " & _
        " AND " & tablePrefix & "DETDSGN3.ACTIVE = '1'" & _
        " AND CASE " & tablePrefix & "DETDSGN3.COLOR_CODE " & _
        "      WHEN 'CLEAR' THEN 'C' " & _
        "      ELSE CASE NVL(" & tablePrefix & "DETCOLR1.POLARIZED,'0') " & _
        "              WHEN '1' THEN 'P' " & _
        "              ELSE 'T' END END = '{1}' " & _
        " ORDER BY " & tablePrefix & "DETCOLR2.COLOR_DESC_WEB "

        Return GetDataTable(String.Format(TSql, lensDesignCode, colorTypeCode))

    End Function

    ''' <summary>
    ''' Return Color List
    ''' </summary>
    ''' <param name="lensDesignerCode">Lens Designer Code</param>
    ''' <param name="lensDesignCode">Lens Design Code</param>
    ''' <param name="colorTypeCode">Color Type Code</param>
    ''' <param name="materialCode">Material Code</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetColors(ByVal lensDesignerCode As String, ByVal lensDesignCode As String, ByVal colorTypeCode As String, ByVal materialCode As String) As DataTable

        ' Return Color list for Brand + LensDesign + ColorType, further restricted by Material
        Dim TSql As String = _
        "SELECT " & _
        " DISTINCT " & tablePrefix & "DETCOLR2.COLOR_CODE_WEB AS COLOR_CODE, " & _
        " " & tablePrefix & "DETCOLR2.COLOR_DESC_WEB AS COLOR_DESC " & _
        " FROM " & tablePrefix & "DETDSGN3 " & _
        " JOIN " & tablePrefix & "DETCOLR1 " & _
        " ON (" & tablePrefix & "DETDSGN3.COLOR_CODE = " & tablePrefix & "DETCOLR1.COLOR_CODE) " & _
        " JOIN " & tablePrefix & "DETCOLR2 " & _
        " ON (" & tablePrefix & "DETCOLR1.COLOR_CODE_WEB = " & tablePrefix & "DETCOLR2.COLOR_CODE_WEB) " & _
        " WHERE " & tablePrefix & "DETDSGN3.LENS_DESIGN_CODE = '{0}' " & _
        " AND " & tablePrefix & "DETDSGN3.ACTIVE = '1' " & _
        " AND CASE " & tablePrefix & "DETDSGN3.COLOR_CODE " & _
        "      WHEN 'CLEAR' THEN 'C' " & _
        "      ELSE CASE NVL(" & tablePrefix & "DETCOLR1.POLARIZED,'0') " & _
        "              WHEN '1' THEN 'P' " & _
        "              ELSE 'T' END END = '{1}' " & _
        " AND " & tablePrefix & "DETDSGN3.MATL_CODE = '{2}' " & _
        " ORDER BY " & tablePrefix & "DETCOLR2.COLOR_DESC_WEB "

        Return GetDataTable(String.Format(TSql, lensDesignCode, colorTypeCode, materialCode))

    End Function

    ''' <summary>
    ''' Return Mirror Coating Color list
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetMirrorCoatingColors() As DataTable

        Dim TSql As String = _
        "SELECT MIRROR_COATING_COLOR FROM " & tablePrefix & "DETCOLMC WHERE NVL(MIRROR_ACTIVATE_FOR_WEB,'0') = '1' "

        Return GetDataTable(TSql)

    End Function

    ''' <summary>
    ''' Return Frame Types
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFrameTypes() As DataTable

        Dim TSql As String = _
         "SELECT " & _
         "  FRAME_TYPE_CODE, " & _
         "  FRAME_TYPE_DESC, " & _
         "  FRAME_FTYP, " & _
         "  CAST((CASE FRAME_TYPE_CODE WHEN 'RIMLESS' THEN 1 ELSE 0 END) AS BIT) USE_SHAPE_EDGE " & _
         " FROM " & tablePrefix & "DETFRAM1 " & _
         " WHERE FRAME_TYPE_CODE <> 'NONE' " & _
         " ORDER BY FRAME_FTYP"

        Return GetDataTable(TSql)

    End Function

    ''' <summary>
    ''' Get Tints
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTints() As DataTable

        Dim TSql As String = _
          "SELECT " & _
          "  TINT_CODE, " & _
          "  TINT_DESC, " & _
          "  TINT_SEQ " & _
          " FROM " & tablePrefix & "DETTINT1 " & _
          " ORDER BY TINT_SEQ"

        Return GetDataTable(TSql)

    End Function

    ''' <summary>
    ''' Get Redo Reasons
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetRedoReasons() As DataTable

        Dim TSql As String = _
          "SELECT " & _
          "  REASON_CODE, " & _
          "  REASON_DESC, " & _
          "  REASON_CODE REASON_SEQ" & _
          " FROM " & tablePrefix & "DETREAS1 " & _
          " WHERE REASON_TYPE = 'R' " & _
          " AND NVL(INT_OR_WEB,'') = 'W' " & _
          " ORDER BY REASON_CODE"

        Return GetDataTable(TSql)


    End Function

#End Region

#Region " Get Rules (Display Messages + RX Ranges + RX Defaults) "

    ''' <summary>
    ''' Returns Error Messages for Job Validation
    ''' </summary>
    ''' <param name="job">Header Datarow for the Job</param>
    ''' <param name="withDefaults"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetRules(ByVal job As JobValidatorFields, ByVal withDefaults As Boolean) As DelJobRules

        Dim rules = New DelJobRules()

        rules.DisplayMessages = GetDisplayMessages(job)
        rules.RxRanges = GetRxRanges(job)
        If withDefaults Then
            rules.RxDefaults = GetDefaults(job)
        End If

        Return rules

    End Function

    ''' <summary>
    ''' Gets RX Ranges for a Design / Material
    ''' </summary>
    ''' <param name="job"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetRxRanges(ByVal job As JobValidatorFields) As DelJobRules.Ranges

        Dim TSql As String = "SELECT * FROM " & tablePrefix & "DETDSGN6 WHERE LENS_DESIGN_CODE = '{0}' AND MATL_CODE = '{1}'"

        Dim dt = GetDataTable(String.Format(TSql, job.LensDesign, job.Material))

        If dt.Rows.Count <> 1 Then
            Throw New Exception(String.Format("Invalid Design/Material Combination ({0}, {1})", job.LensDesign, job.Material))
            'Return rules
        End If

        Dim dr As DataRow = dt.Rows(0)

        Dim ranges = New DelJobRules.Ranges()

        With ranges

            .DistanceSph.SetRange(GetValue(dr.Item("SPHERE_NEG_MAX"), -17), GetValue(dr.Item("SPHERE_POS_MAX"), 17))

            Dim cylinderValue = GetValue(dr.Item("CYL_NEG_MAX"), -17)
            .DistanceCyl.SetRange(cylinderValue, -1 * cylinderValue)

            .NearAdd.SetRange(-17, 17) 'only applicable for progressives

            .DistanceAxis.SetRange(0, 180)
            .MonocularPD.SetRange(22, 40)

            Dim sqlMinFittingHeight As String = "SELECT MIN(MIN_FITTING_HEIGHT) FROM " & tablePrefix & "DETDSGN2 WHERE LENS_DESIGN_CODE = '{0}'"
            Dim minFittingHeight As Integer = Val(GetDataValue(String.Format(sqlMinFittingHeight, job.LensDesign)) & String.Empty)

            .FittingHeight.SetRange(minFittingHeight, 999)    '" & tablePrefix & "DETDSGN2.MIN_FITTING_HEIGHT - is a warning, not a hard error

            .PrismIn.SetRange(-17, 17)
            .PrismUp.SetRange(-17, 17)

            .FittingVertex.SetRange(-999, 999)
            .RefractiveVertex.SetRange(-999, 999)
            .PantoTilt.SetRange(-999, 999)
            .WrapAngle.SetRange(-999, 999)

            'MAX VALUES IN DETPARM1
            Dim sqlDETPARM1 As String = "SELECT *  FROM " & tablePrefix & "DETPARM1"
            Dim dtDETPARM1 = GetDataTable(sqlDETPARM1)
            If dt.Rows.Count <> 1 Then
                'Throw New Exception(String.Format("Invalid Design/Material Combination ({0}, {1})", job.LensDesign, job.Material))
                Return ranges
            End If

            Dim drDETPARM1 = dtDETPARM1.Rows(0)

            .FrameA.SetRange(0, GetReaderHighValue(drDETPARM1.Item("DE_PARM_MAX_A")))
            .FrameB.SetRange(0, GetReaderHighValue(drDETPARM1.Item("DE_PARM_MAX_B")))
            .FrameDbl.SetRange(0, GetReaderHighValue(drDETPARM1.Item("DE_PARM_MAX_DBL")))
            .FrameED.SetRange(0, GetReaderHighValue(drDETPARM1.Item("DE_PARM_MAX_ED")))
        End With

        Return ranges

    End Function

    Private Shared Function GetValue(ByVal readerValue As Object, ByVal ifNullValue As Decimal) As Decimal

        If readerValue Is System.DBNull.Value Then
            Return ifNullValue
        End If

        Return Convert.ToDecimal(readerValue)

    End Function

    Private Shared Function GetReaderLowValue(ByVal readerValue As Object) As Decimal

        If readerValue Is System.DBNull.Value Then
            Return -999.99
        End If

        Return Convert.ToDecimal(readerValue)

    End Function

    Private Shared Function GetReaderHighValue(ByVal readerValue As Object) As Decimal

        If readerValue Is System.DBNull.Value Then
            Return 999.99
        End If

        Return Convert.ToDecimal(readerValue)

    End Function

    Private Shared Function GetDefaults(ByVal job As JobValidatorFields)

        Dim defaults = New DelJobRules.Defaults()

        Dim sqlDefaults = "SELECT * FROM " & tablePrefix & "DETDEFD1 WHERE LENS_DESIGN_CODE = '{0}'"

        Dim dtDefaults = GetDataTable(String.Format(sqlDefaults, job.LensDesign))

        With defaults
            .FittingVertex = Val(SelectValue(dtDefaults, "VCA_KEY='BVD'", "VALUE_NUM"))
            .RefractiveVertex = Val(SelectValue(dtDefaults, "VCA_KEY='RVD'", "VALUE_NUM"))
            .PantoTilt = Val(SelectValue(dtDefaults, "VCA_KEY='PANTO'", "VALUE_NUM"))
            .WrapAngle = Val(SelectValue(dtDefaults, "VCA_KEY='ZTILT'", "VALUE_NUM"))
            Dim VIEWP() = Split(SelectValue(dtDefaults, "VCA_KEY='VIEWP'", "VALUE"), ";")
            Dim CLLIFE = SelectValue(dtDefaults, "VCA_KEY='CLLIFE'", "VALUE")
            If VIEWP.GetLength(0) > 1 Then
                .WorkingDistancesDistance = Val(VIEWP(1))
                .WorkingDistancesViewingAngle = Convert.ToString(Val(VIEWP(2)) / 20 + 1)  'value will be -20, 0, or 20 currently - but web is looking for 0, 1, 2
                .WorkingDistancesActivity = IIf(CLLIFE = "ACTIVE", "1", "0")
            Else
                .WorkingDistancesDistance = 0
                .WorkingDistancesViewingAngle = 0
                .WorkingDistancesActivity = String.Empty
            End If
            .WorkingDistancesNear = Val(SelectValue(dtDefaults, "VCA_KEY='NWD'", "VALUE_NUM"))
            .WorkingDistancesFar = Val(SelectValue(dtDefaults, "VCA_KEY='FWD'", "VALUE_NUM"))
        End With

        Return defaults

    End Function

    Private Shared Function GetDisplayMessages(ByVal job As JobValidatorFields) As List(Of String)

        Dim list = New List(Of String)

        ' TBD: Confirm need for frame trace message
        If Math.Sign(job.DistanceRxLeftSph) <> Math.Sign(job.DistanceRxRightSph) Then
            list.Add("WARNING: Opposite signs on lens sphere")
        End If

        ' TBD: Add more messages you want the user to see here

        Return list

    End Function

#End Region

#Region " Validate Job "

    Public Function ValidateJobData(ByVal job As JobValidatorFields) As List(Of DelJobValidationError)

        ' Other things to consider to validate
        'PROMO_CODE


        Dim errors As New List(Of DelJobValidationError)

        If job.PatientFirstName.Trim.Length = 0 And job.PatientLastName.Trim.Length = 0 And job.Tray.Trim.Length = 0 Then
            errors.Add(New DelJobValidationError("PatientTray", "Patient First Name, Last Name, or Tray is required"))
        End If

        If (job.PatientFirstName.Trim.Length = 0 OrElse job.PatientLastName.Trim.Length = 0) And job.SaveNewPatient Then
            errors.Add(New DelJobValidationError("PatientTray", "Patient First and Last Name is required to save a new patient file"))
        End If

        If job.RedoOfJobNumber.Trim().Length > 0 Then
            If job.ReasonCodeRedo.Trim().Length = 0 Then
                errors.Add(New DelJobValidationError("RedoReason", "Redo Reason is required"))
            End If
        End If

        Dim isvalidAccount As Boolean = False
        If job.AccountNumber.Trim().Length = 0 Then
            errors.Add(New DelJobValidationError("AccountNumber", "A Valid Account Number is required"))
        Else
            If GetDataValue("select cust_code from artcust1 where cust_code = '" & job.AccountNumber.Trim & "'") & String.Empty = String.Empty Then
                errors.Add(New DelJobValidationError("AccountNumber", "A Valid Account Number is required"))
            Else
                isvalidAccount = True
            End If
        End If

        If isvalidAccount AndAlso job.ShipToID.Length > 0 Then
            If GetDataValue("select cust_code from artcust2 where cust_code = '" & job.AccountNumber.Trim & "' and cust_ship_to_to = '" & job.ShipToID & "'") & String.Empty = String.Empty Then
                errors.Add(New DelJobValidationError("AccountNumber", "A Valid Account Number is required"))
            End If
        End If

        If job.JobType.Trim().Length = 0 Then
            errors.Add(New DelJobValidationError("JobType", "Job Type is required"))
        End If

        If job.LensDesigner.Trim().Length = 0 Then
            errors.Add(New DelJobValidationError("LensDesigner", "Lens Designer is required"))
        End If

        If job.LensDesign.Trim().Length = 0 Then
            errors.Add(New DelJobValidationError("LensDesign", "Lens Design is required"))
        End If

        If job.ColorType.Trim().Length = 0 Then
            errors.Add(New DelJobValidationError("ColorType", "Color Type is required"))
        End If

        If job.Material.Trim().Length = 0 Then
            errors.Add(New DelJobValidationError("Material", "Material is required"))
        End If

        If job.Color.Trim().Length = 0 Then
            errors.Add(New DelJobValidationError("Color", "Color is required"))
        End If

        If job.Tint.Trim().Length = 0 Then
            errors.Add(New DelJobValidationError("Tint", "Tint is required"))
        Else
            If job.Tint <> "NONE" Then
                If job.TintColor.Trim().Length = 0 Then
                    errors.Add(New DelJobValidationError("TintColor", "Tint Color is required"))
                End If
                'If job.TintPercent < 0 OrElse job.TintPercent > 100 Then
                '    errors.Add(New DelJobValidationError("TintPercent", "Tint % must be between 0 and 100"))
                'End If
                If Not "5, 10, 20, 30, 40, 50, 60, 70, 80".Contains(job.TintPercent) Then
                    errors.Add(New DelJobValidationError("TintPercent", "Tint % must be one of the following values: 5, 10, 20, 30, 40, 50, 60, 70, 80"))
                End If
            End If
        End If

        'DistanceRxLeftAxis, DistanceRxRightAxis, DistanceRxLeftCyl, DistanceRxRightCyl, NearRxRightAdd, NearRxLeftAdd, FittingMeasurementsRightMonocularPD, FittingMeasurementsLeftMonocularPD

        Dim axisErrs = 0
        If job.DistanceRxLeftAxis = 0 And job.DistanceRxLeftCyl <> 0 And job.WhichLenses <> "R" Then
            axisErrs += 1
        End If

        If job.DistanceRxRightAxis = 0 And job.DistanceRxRightCyl <> 0 And job.WhichLenses <> "L" Then
            axisErrs += 2
        End If

        Select Case axisErrs
            Case 1
                errors.Add(New DelJobValidationError("DistanceRxAxisL", "Value for left axis required"))
            Case 2
                errors.Add(New DelJobValidationError("DistanceRxAxisR", "Value for right axis required"))
            Case 3
                errors.Add(New DelJobValidationError("DistanceRxAxisR", "Value for right axis required"))
                errors.Add(New DelJobValidationError("DistanceRxAxisL", "Value for left axis required"))
        End Select

        If job.DistanceRxLeftCyl = 0 And job.DistanceRxLeftAxis <> 0 And job.WhichLenses <> "R" Then
            errors.Add(New DelJobValidationError("DistanceRxCylL", "Value for left cylinder required"))
        End If
        If job.DistanceRxRightCyl = 0 And job.DistanceRxRightAxis <> 0 And job.WhichLenses <> "L" Then
            errors.Add(New DelJobValidationError("DistanceRxCylR", "Value for right cylinder required"))
        End If

        If job.WhichLenses = "B" Then
            If Math.Sign(job.DistanceRxLeftCyl) <> Math.Sign(job.DistanceRxRightCyl) And job.DistanceRxRightCyl <> 0 And job.DistanceRxLeftCyl <> 0 Then
                errors.Add(New DelJobValidationError("DistanceRxCylL", "Left and right cylinders must be the same sign (both negative or both positive)"))
            End If
        End If

        If Math.Ceiling(job.DistanceRxRightAxis * 4) <> Math.Floor(job.DistanceRxRightAxis * 4) And job.WhichLenses <> "L" Then
            errors.Add(New DelJobValidationError("DistanceRxAxisR", "Values for right axis must be in quarter increments"))
        End If
        If Math.Ceiling(job.DistanceRxLeftAxis * 4) <> Math.Floor(job.DistanceRxLeftAxis * 4) And job.WhichLenses <> "R" Then
            errors.Add(New DelJobValidationError("DistanceRxAxisL", "Values for left axis must be in quarter increments"))
        End If

        If Math.Ceiling(job.NearRxRightAdd * 4) <> Math.Floor(job.NearRxRightAdd * 4) And job.WhichLenses <> "L" Then
            errors.Add(New DelJobValidationError("NearRxAddR", "Values for right add power must be in quarter increments"))
        End If
        If Math.Ceiling(job.NearRxLeftAdd * 4) <> Math.Floor(job.NearRxLeftAdd * 4) And job.WhichLenses <> "R" Then
            errors.Add(New DelJobValidationError("NearRxAddL", "Values for left add power must be in quarter increments"))
        End If

        If Math.Ceiling(job.FittingMeasurementsRightMonocularPD * 2) <> Math.Floor(job.FittingMeasurementsRightMonocularPD * 2) And job.WhichLenses <> "L" Then
            errors.Add(New DelJobValidationError("MonocularPDR", "Values for right Monocular PD must be in half increments"))
        End If
        If Math.Ceiling(job.FittingMeasurementsLeftMonocularPD * 2) <> Math.Floor(job.FittingMeasurementsLeftMonocularPD * 2) And job.WhichLenses <> "R" Then
            errors.Add(New DelJobValidationError("MonocularPDL", "Values for left Monocular PD must be in half increments"))
        End If

        If Math.Ceiling(job.DistanceRxRightCyl * 4) <> Math.Floor(job.DistanceRxRightCyl * 4) And job.WhichLenses <> "L" Then
            errors.Add(New DelJobValidationError("DistanceRxCylR", "Values for cylinder must be in quarter increments"))
        End If
        If Math.Ceiling(job.DistanceRxLeftCyl * 4) <> Math.Floor(job.DistanceRxLeftCyl * 4) And job.WhichLenses <> "R" Then
            errors.Add(New DelJobValidationError("DistanceRxCylL", "Values for cylinder must be in quarter increments"))
        End If

        If Math.Ceiling(job.DistanceRxRightSph * 4) <> Math.Floor(job.DistanceRxRightSph * 4) And job.WhichLenses <> "L" Then
            errors.Add(New DelJobValidationError("DistanceRxSphR", "Values for sphere must be in quarter increments"))
        End If
        If Math.Ceiling(job.DistanceRxLeftSph * 4) <> Math.Floor(job.DistanceRxLeftSph * 4) And job.WhichLenses <> "R" Then
            errors.Add(New DelJobValidationError("DistanceRxSphL", "Values for sphere must be in quarter increments"))
        End If

        If job.IncludePrism Then
            If "IO".IndexOf(job.PrismRightInAxis) = -1 And job.WhichLenses <> "L" Then
                errors.Add(New DelJobValidationError("PrismInR", "Right RX Prism In/Out axis is required"))
            End If
            If "IO".IndexOf(job.PrismLeftInAxis) = -1 And job.WhichLenses <> "R" Then
                errors.Add(New DelJobValidationError("PrismInL", "Left RX Prism In/Out axis is required"))
            End If
            If "UD".IndexOf(job.PrismRightUpAxis) = -1 And job.WhichLenses <> "L" Then
                errors.Add(New DelJobValidationError("PrismUpR", "Right Prism Up/Down axis is required"))
            End If
            If "UD".IndexOf(job.PrismLeftUpAxis) = -1 And job.WhichLenses <> "R" Then
                errors.Add(New DelJobValidationError("PrismUpL", "Left Prism Up/Down axis is required"))
            End If

            If ((job.PrismLeftIn = 0 And job.PrismLeftInAxis <> "") Or (job.PrismLeftIn <> 0 And job.PrismLeftInAxis = "")) And job.WhichLenses <> "R" Then
                errors.Add(New DelJobValidationError("PrismInL", "Must select both a Left Prism In value and Left Prism In Axis value"))
            End If
            If ((job.PrismRightIn = 0 And job.PrismRightInAxis <> "") Or (job.PrismRightIn <> 0 And job.PrismRightInAxis = "")) And job.WhichLenses <> "L" Then
                errors.Add(New DelJobValidationError("PrismInR", "Must select both a Right Prism In value and Right Prism In Axis value"))
            End If
            If ((job.PrismLeftUp = 0 And job.PrismLeftUpAxis <> "") Or (job.PrismLeftUp <> 0 And job.PrismLeftUpAxis = "")) And job.WhichLenses <> "R" Then
                errors.Add(New DelJobValidationError("PrismUpL", "Must select both a Left Prism Up value and Left Prism Up Axis value"))
            End If
            If ((job.PrismRightUp = 0 And job.PrismRightUpAxis <> "") Or (job.PrismRightUp <> 0 And job.PrismRightUpAxis = "")) And job.WhichLenses <> "L" Then
                errors.Add(New DelJobValidationError("PrismUpR", "Must select both a Right Prism Up value and Right Prism Up Axis value"))
            End If
        End If

        ValidateRxNumbers(job, errors)

        If job.FrameType.Trim().Length = 0 Then
            errors.Add(New DelJobValidationError("FrameType", "Frame Type is required"))
        Else
            If job.FrameType <> "NONE" Then
                If job.FrameDataA = 0 Then
                    errors.Add(New DelJobValidationError("FrameDataA", "Frame A Width value is required"))
                End If
                If job.FrameDataB = 0 Then
                    errors.Add(New DelJobValidationError("FrameDataB", "Frame B Height value is required"))
                End If
                If job.FrameDataED = 0 Then
                    errors.Add(New DelJobValidationError("FrameDataED", "Frame ED Diagonal value is required"))
                End If

                Dim frameDataDbl = GetFrameDataDbl(job)
                If frameDataDbl Is Nothing Then
                    errors.Add(New DelJobValidationError("FrameDataDbl", "Frame Dbl Bridge value is required"))
                Else
                    If Math.Ceiling(frameDataDbl * 2) <> Math.Floor(frameDataDbl * 2) Then
                        errors.Add(New DelJobValidationError("FrameDataDbl", "Frame Dbl Bridge value must be in 0.5 increments"))
                    End If
                End If

                If Math.Ceiling(job.FrameDataA * 2) <> Math.Floor(job.FrameDataA * 2) Then
                    errors.Add(New DelJobValidationError("FrameDataA", "Frame A Width value must be in 0.5 increments"))
                End If

                If Math.Ceiling(job.FrameDataB * 2) <> Math.Floor(job.FrameDataB * 2) Then
                    errors.Add(New DelJobValidationError("FrameDataB", "Frame B Height value must be in 0.5 increments"))
                End If

                If Math.Ceiling(job.FrameDataED * 2) <> Math.Floor(job.FrameDataED * 2) Then
                    errors.Add(New DelJobValidationError("FrameDataED", "Frame ED Diagonal value must be in 0.5 increments"))
                End If

                If job.FrameDataED < job.FrameDataA Then
                    errors.Add(New DelJobValidationError("FrameDataED", String.Format("Frame ED Diagonal ({0}) cannot be less than Frame A Width ({1})", job.FrameDataED, job.FrameDataA)))
                End If
                If job.FrameDataED < job.FrameDataB Then
                    errors.Add(New DelJobValidationError("FrameDataED", String.Format("Frame ED Diagonal ({0}) cannot be less than Frame B Height ({1})", job.FrameDataED, job.FrameDataB)))
                End If


                Dim sqlFrameMatl As String = _
                   "SELECT " & tablePrefix & "DETMATL1.MATL_DESC," & tablePrefix & "DETFRAM1.FRAME_ETYP, " & tablePrefix & "DETMATL1.PERMITTED_WITH_RIMLESS FROM " & tablePrefix & "DETFRAM1," & tablePrefix & "DETMATL1 " & _
                   "WHERE " & tablePrefix & "DETFRAM1.FRAME_TYPE_CODE='{0}' AND " & tablePrefix & "DETMATL1.MATL_CODE='{1}'"

                Dim dtFrameMatl = GetDataTable(String.Format(sqlFrameMatl, job.FrameType, job.Material))

                If dtFrameMatl.Rows.Count > 0 Then
                    Dim drFrameMatl = dtFrameMatl.Rows(0)
                    If drFrameMatl.Item("PERMITTED_WITH_RIMLESS") <> "1" And drFrameMatl.Item("FRAME_ETYP") = "2" Then
                        errors.Add(New DelJobValidationError("Material", String.Format("Material {0} is not available for rimless", drFrameMatl.Item("MATL_DESC"))))
                    End If
                End If
            End If
        End If

        ' Validations for Finished Work
        If job.JobType = "F" Then
            If job.Finishedwork.Trim().Length = 0 Then
                errors.Add(New DelJobValidationError("FinishedWork", "Finished Work Frame-To-Come or Shape/Edge is required"))
            ElseIf job.Finishedwork.Trim().Length <> 1 OrElse Not "FS".Contains(job.JobType.Trim()) Then
                errors.Add(New DelJobValidationError("FinishedWork", "Invalid value for Finished Work"))
            End If
            If job.FrameManufacturer.Trim().Length = 0 AndAlso job.FrameModelNumber.Trim().Length = 0 Then
                errors.Add(New DelJobValidationError("FrameManufacturer", "Finished Work Frame Manufacturer or Model Number is required"))
            End If
        End If


        If job.LensDesigner <> "SEIKO" Then 'And Absx1.optFor("FINISHED").Value <> "O" Then
            'If LookUp("DETDSGN4", New String() {Absx1.txtFor("LENS_DESIGNER_CODE").Text, Absx1.txtFor("MATL_CODE").Text & "", Absx1.txtFor("COLOR_CODE").Text}) Is Nothing Then
            'errors.Add(New DelJobValidationError("LMATID", "No LMATID on file for Designer, Material & Color Specified"))
            'End If
        End If


        If job.FittingMeasurementsLeftMonocularPD = 0 AndAlso job.WhichLenses <> "R" Then
            errors.Add(New DelJobValidationError("MonocularPDL", "Left Monocular PD is required"))
        End If
        If job.FittingMeasurementsRightMonocularPD = 0 AndAlso job.WhichLenses <> "L" Then
            errors.Add(New DelJobValidationError("MonocularPDR", "Right Monocular PD is required"))
        End If

        If job.FittingMeasurementsLeftFittingHeight <= 0 AndAlso job.WhichLenses <> "R" Then
            errors.Add(New DelJobValidationError("FittingHeightL", "Left Fitting Height must be greater than 0"))
        End If
        If job.FittingMeasurementsRightFittingHeight <= 0 AndAlso job.WhichLenses <> "L" Then
            errors.Add(New DelJobValidationError("FittingHeightR", "Right Fitting Height must be greater than 0"))
        End If

        Dim drLensDesign As DataRow = GetDataRow("select * from detdsgn1 where lens_design_code = '" & job.LensDesign & "'")
        If drLensDesign IsNot Nothing Then
            If drLensDesign.Item("LENS_TYPE") & String.Empty = "P" Then 'PROGRESSIVE LENS
                Dim addErrs = 0
                If job.NearRxLeftAdd <= 0 And job.WhichLenses <> "R" Then
                    addErrs += 1
                End If
                If job.NearRxRightAdd <= 0 And job.WhichLenses <> "L" Then
                    addErrs += 2
                End If
                Select Case addErrs
                    Case 1
                        errors.Add(New DelJobValidationError("NearRxAddL", "Left add power is required for progressive lens designs"))
                    Case 2
                        errors.Add(New DelJobValidationError("NearRxAddR", "Right add power is required for progressive lens designs"))
                    Case 3
                        errors.Add(New DelJobValidationError("NearRxAddL", "Right add power is required for progressive lens designs"))
                        errors.Add(New DelJobValidationError("NearRxAddR", "Left add power is required for progressive lens designs"))
                End Select
            Else 'SINGLE VISION LENS
                If job.WhichLenses = "B" Or job.WhichLenses = "R" Then
                    If job.FittingMeasurementsRightFittingHeight < 0.25 * job.FrameDataB Or job.FittingMeasurementsRightFittingHeight > 0.75 * job.FrameDataB Then
                        errors.Add(New DelJobValidationError("FittingHeightR", String.Format("Right fitting height must be between 25% ({0}) and 75% ({1}) of Frame B Height", 0.25 * job.FrameDataB, 0.75 * job.FrameDataB)))
                    End If
                End If
                If job.WhichLenses = "B" Or job.WhichLenses = "L" Then
                    If job.FittingMeasurementsLeftFittingHeight < 0.25 * job.FrameDataB Or job.FittingMeasurementsLeftFittingHeight > 0.75 * job.FrameDataB Then
                        errors.Add(New DelJobValidationError("FittingHeightL", String.Format("Left fitting height must be between 25% ({0}) and 75% ({1}) of Frame B Height", 0.25 * job.FrameDataB, 0.75 * job.FrameDataB)))
                    End If
                End If
            End If
        End If

        Return errors

    End Function

    Private Shared Function GetFrameDataDbl(ByVal job As JobValidatorFields) As Decimal?

        If job.FrameDataDbl.Trim().Length = 0 Then
            Return Nothing
        End If

        Dim asDecimal As Decimal
        If Not Decimal.TryParse(job.FrameDataDbl, asDecimal) Then
            Return Nothing
        End If

        Return asDecimal

    End Function

    Private Shared Sub ValidateRxNumbers(ByVal job As JobValidatorFields, ByVal errors As List(Of DelJobValidationError))

        Dim ranges As DelJobRules.Ranges
        Try
            ranges = GetRules(job, False).RxRanges
        Catch ex As Exception
            ' If we can't get the rules, then we can't validate the RX numbers, but
            ' will still allow validation to continue without an exception. We rely on
            ' the cause of the exception retrieving the rules (e.g., missing color) 
            ' to be flagged with its own validation rule (e.g., color is required)
            Return
        End Try

        Dim checkRight = (job.WhichLenses = "B" Or job.WhichLenses = "R")
        Dim checkLeft = (job.WhichLenses = "B" Or job.WhichLenses = "L")

        ValidateRxNumber(errors, checkRight, checkLeft, ranges.DistanceSph, job.DistanceRxRightSph, job.DistanceRxLeftSph, "DistanceRxSph", "Distance RX Sphere")
        ValidateRxNumber(errors, checkRight, checkLeft, ranges.DistanceCyl, job.DistanceRxRightCyl, job.DistanceRxLeftCyl, "DistanceRxCyl", "Distance RX Cylinder")
        ValidateRxNumber(errors, checkRight, checkLeft, ranges.DistanceAxis, job.DistanceRxRightAxis, job.DistanceRxLeftAxis, "DistanceRxAxis", "Distance RX Axis")
        ValidateRxNumber(errors, checkRight, checkLeft, ranges.NearAdd, job.NearRxRightAdd, job.NearRxLeftAdd, "NearRxAdd", "Near RX Add")
        If job.IncludePrism Then
            ValidateRxNumber(errors, checkRight, checkLeft, ranges.PrismIn, job.PrismRightIn, job.PrismLeftIn, "PrismIn", "Prism In")
            ValidateRxNumber(errors, checkRight, checkLeft, ranges.PrismUp, job.PrismRightUp, job.PrismLeftUp, "PrismUp", "Prism Up")
        End If
        ValidateRxNumber(errors, checkRight, checkLeft, ranges.FittingVertex, job.VertexPantoRightFittingVertex, job.VertexPantoLeftFittingVertex, "FittingVertex", "Fitting Vertex")
        ValidateRxNumber(errors, checkRight, checkLeft, ranges.RefractiveVertex, job.VertexPantoRightRefractiveVertex, job.VertexPantoLeftRefractiveVertex, "RefractiveVertex", "Refractive Vertex")
        ValidateRxNumber(errors, checkRight, checkLeft, ranges.PantoTilt, job.VertexPantoRightPantoTilt, job.VertexPantoLeftPantoTilt, "PantoTilt", "Panto Tilt")
        ValidateRxNumber(errors, checkRight, checkLeft, ranges.WrapAngle, job.VertexPantoRightWrapAngle, job.VertexPantoLeftWrapAngle, "WrapAngle", "Wrap Angle")
        ValidateRxNumber(errors, checkRight, checkLeft, ranges.MonocularPD, job.FittingMeasurementsRightMonocularPD, job.FittingMeasurementsLeftMonocularPD, "MonocularPD", "Monocular PD")
        ValidateRxNumber(errors, checkRight, checkLeft, ranges.FittingHeight, job.FittingMeasurementsRightFittingHeight, job.FittingMeasurementsLeftFittingHeight, "FittingHeight", "Fitting Height")
        ValidateRxNumber(errors, ranges.FrameA, job.FrameDataA, "FrameDataA", "Frame A (Width)")
        ValidateRxNumber(errors, ranges.FrameB, job.FrameDataB, "FrameDataB", "Frame B (Height)")
        ValidateRxNumber(errors, ranges.FrameDbl, GetFrameDataDbl(job), "FrameDataDbl", "Frame DBL (Bridge)")
        ValidateRxNumber(errors, ranges.FrameED, job.FrameDataED, "FrameDataED", "Frame ED (Diagonal)")

    End Sub

    Private Shared Sub ValidateRxNumber( _
      ByVal errors As List(Of DelJobValidationError), _
      ByVal checkRight As Boolean, _
      ByVal checkLeft As Boolean, _
      ByVal range As DelJobRules.Range, _
      ByVal rightNumber As Decimal?, _
      ByVal leftNumber As Decimal?, _
      ByVal rxNumberKey As String, _
      ByVal rxNumberDescription As String)

        If checkRight Then
            If rightNumber.HasValue AndAlso rightNumber <> 0 AndAlso (rightNumber < range.MinValue OrElse rightNumber > range.MaxValue) Then
                errors.Add(New DelJobValidationError(String.Format("{0}R", rxNumberKey), String.Format("{0} Right value must be between {1} and {2}", rxNumberDescription, range.MinValue, range.MaxValue)))
            End If

        End If
        If checkLeft Then
            If leftNumber.HasValue AndAlso leftNumber <> 0 AndAlso (leftNumber < range.MinValue OrElse leftNumber > range.MaxValue) Then
                errors.Add(New DelJobValidationError(String.Format("{0}L", rxNumberKey), String.Format("{0} Left value must be between {1} and {2}", rxNumberDescription, range.MinValue, range.MaxValue)))
            End If
        End If

    End Sub

    Private Shared Sub ValidateRxNumber( _
     ByVal errors As List(Of DelJobValidationError), _
     ByVal range As DelJobRules.Range, _
     ByVal number As Decimal?, _
     ByVal rxNumberKey As String, _
     ByVal rxNumberDescription As String)

        If number.HasValue AndAlso number <> 0 AndAlso (number < range.MinValue OrElse number > range.MaxValue) Then
            errors.Add(New DelJobValidationError(rxNumberKey, String.Format("{0} value must be between {1} and {2}", rxNumberDescription, range.MinValue, range.MaxValue)))
        End If

    End Sub

    Public Shared Function ValidateOrderdata(ByVal order As OrderValidatorFields) As List(Of DelJobValidationError)

        Dim errors As New List(Of DelJobValidationError)


        If order.CustomerNameBillTo.Length = 0 _
            OrElse order.CustomerAddressBillTo.Length = 0 _
            OrElse order.CustomerCityBillTo.Length = 0 _
            OrElse order.CustomerStateBillTo.Length = 0 _
            OrElse order.CustomerZipcodeBillTo.Length = 0 Then
            errors.Add(New DelJobValidationError("BillToAddress", "Name, Street address, City, State and Zip are required"))
        End If

        If order.CustomerNameShipTo.Length = 0 _
            OrElse order.CustomerAddressShipTo.Length = 0 _
            OrElse order.CustomerCityShipTo.Length = 0 _
            OrElse order.CustomerStateShipTo.Length = 0 _
            OrElse order.CustomerZipcodeShipTo.Length = 0 Then
            errors.Add(New DelJobValidationError("ShipToAddress", "Name, Street address, City, State and Zip are required"))
        End If


        Return errors

    End Function

#End Region

#Region " Allocate Job Number "

    Public Shared Function AllocateNextJobNumber(ByVal accountNumber As String, ByVal ipAddress As String) As String
        ' Commented out by Ed
        'Dim inTrans = True
        'Dim delJobNo = ""
        'If connInfo.Transaction Is Nothing Then
        '    inTrans = False
        '    connInfo._txn = connInfo.Connection.BeginTransaction()
        'End If

        'Const getNextJobNumber = "SELECT MIN(JOB_NO) FROM " & tablePrefix & "DETJNUM1 (UPDLOCK) WHERE AVAILABLE='1'"
        'Dim drJobNo = DataAccessHelper.FillRow(connInfo, getNextJobNumber)
        'delJobNo = drJobNo.Item(0).ToString.PadLeft(6, "0")

        'Dim an As AccountNumber = New AccountNumber(accountNumber)
        'DataAccessHelper.ExecuteSql(connInfo, String.Format("UPDATE " & tablePrefix & "DETJNUM1 SET AVAILABLE='2',ALLOCATED_DATE=GETDATE(),CUST_CODE='{1}',CUST_SHIP_TO_NO='{2}',IP_ADDRESS='{3}' WHERE JOB_NO = '{0}'", delJobNo, an.CustomerCode, If(an.HasShipTo, an.CustomerShipTo, ""), ipAddress))

        'If Not inTrans Then
        '    connInfo.Transaction.Commit()
        '    connInfo._txn.Dispose()
        '    connInfo._txn = Nothing
        'End If

        'Return delJobNo

        Return String.Empty
    End Function

#End Region

#Region " Get Job "

    ''' <summary>
    ''' Create a Job Object used for validation
    ''' </summary>
    ''' <param name="DETJOBM1"></param>
    ''' <param name="DETJOBM3"></param>
    ''' <param name="JobStatusCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateJobObject(ByRef DETJOBM1 As DataRow, ByVal DETJOBM3 As DataTable, ByVal JobStatusCode As String) As JobValidatorFields

        Dim job As New JobValidatorFields
        Dim sql As String = String.Empty

        With job
            .DelJobId = 0
            .JobNumber = DETJOBM1.Item("JOB_NO") & String.Empty
            .SubmittedOn = DETJOBM1.Item("ORDR_DATE") & String.Empty
            .RedoOfJobNumber = DETJOBM1.Item("JOB_NO_ORIG") & String.Empty
            .ReasonCodeRedo = DETJOBM1.Item("REASON_CODE_REDO") & String.Empty
            .PractitionerID = 0
            .AccountNumber = DETJOBM1.Item("CUST_CODE") & String.Empty
            .ShipToID = DETJOBM1.Item("CUST_SHIP_TO_NO") & String.Empty
            .PatientId = 0

            If DETJOBM1.Item("PATIENT_FIRST_NAME") & String.Empty = String.Empty _
                AndAlso DETJOBM1.Item("PATIENT_FIRST_NAME") & String.Empty = String.Empty _
                AndAlso DETJOBM1.Item("TRAY") & String.Empty = String.Empty Then
                .PatientFirstName = DETJOBM1.Item("PATIENT_NAME") & String.Empty
            Else
                .PatientFirstName = DETJOBM1.Item("PATIENT_FIRST_NAME") & String.Empty
            End If

            .PatientLastName = DETJOBM1.Item("PATIENT_LAST_NAME") & String.Empty
            .SaveNewPatient = False
            .Tray = DETJOBM1.Item("TRAY") & String.Empty

            If DETJOBM1.Item("FINISHED") & String.Empty = "O" Then
                .JobType = "F"
            Else
                .JobType = DETJOBM1.Item("FINISHED") & String.Empty
            End If

            .WhichLenses = DETJOBM1.Item("LENS_ORDER") & String.Empty

            .LensDesigner = DETJOBM1.Item("LENS_DESIGNER_CODE") & String.Empty
            .LensDesign = DETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty

            Dim row As DataRow = GetDataRow("SELECT * FROM " & tablePrefix & "DETDSGN1 WHERE LENS_DESIGNER_CODE = '" & .LensDesigner & "'")
            If row Is Nothing OrElse row.Item("PRIVATE_LABEL_DESIGN") & String.Empty <> "1" Then
                .Brand = .LensDesigner
            Else
                .Brand = row.Item("LENS_DESIGNER_NAME_INV") & String.Empty
            End If

            .ColorType = DETJOBM1.Item("COLOR_TYPE") & String.Empty
            .Material = DETJOBM1.Item("MATL_CODE") & String.Empty
            .Color = DETJOBM1.Item("COLOR_CODE") & String.Empty
            .Tint = DETJOBM1.Item("TINT_CODE") & String.Empty
            .TintColor = DETJOBM1.Item("TINT_COLOR") & String.Empty
            .TintPercent = Val(DETJOBM1.Item("TINT_PCT") & String.Empty)

            Select Case DETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty
                Case "AUTOOFFICE" : .UseWorkingDistances = 1
                Case "LIFEMADEW" : .UseWorkingDistances = 2
                Case Else : .UseWorkingDistances = 0
            End Select

            If DETJOBM1.Item("WORKING_DISTANCE") & String.Empty = String.Empty Then
                .WorkingDistancesDistance = -1
            Else
                .WorkingDistancesDistance = DETJOBM1.Item("WORKING_DISTANCE")
            End If


            '" NVL(J1.NWD,0) WorkingDistancesNear," & _
            '" NVL(J1.FWD,0) WorkingDistancesFar," & _

            Dim VIEWING_ANGLE As Decimal = 40
            If DETJOBM1.Item("VIEWING_ANGLE") IsNot DBNull.Value Then
                VIEWING_ANGLE = DETJOBM1.Item("VIEWING_ANGLE")
            End If
            VIEWING_ANGLE /= 20
            VIEWING_ANGLE += 1

            .WorkingDistancesViewingAngle = VIEWING_ANGLE

            If DETJOBM1.Item("ACTIVITY") IsNot DBNull.Value Then
                .WorkingDistancesActivity = DETJOBM1.Item("ACTIVITY")
            Else
                .WorkingDistancesActivity = -1
            End If

            .WorkingDistancesNear = Val(DETJOBM1.Item("NWD") & String.Empty)
            .WorkingDistancesFar = Val(DETJOBM1.Item("FWD") & String.Empty)

            .ArCoating = Val(DETJOBM1.Item("AR_COATING") & String.Empty)
            .BacksideOnly = Val(DETJOBM1.Item("AR_BACKSIDE_ONLY") & String.Empty)
            .MirrorCoating = Val(DETJOBM1.Item("MIRROR_COATING") & String.Empty)
            .MirrorCoatingColor = DETJOBM1.Item("MIRROR_COATING_COLOR") & String.Empty
            .Polish = Val(DETJOBM1.Item("POLISHING") & String.Empty)
            .TintableHardcoat = Val(DETJOBM1.Item("UNCUT_TINTABLE") & String.Empty)

            sql = "JOB_NO = '" & .JobNumber & "' AND RL = 'R'"
            If DETJOBM3.Select(sql).Length > 0 Then
                With DETJOBM3.Select(sql)(0)
                    job.DistanceRxRightSph = Val(.Item("SPHERE") & String.Empty)
                    job.DistanceRxRightCyl = Val(.Item("CYLINDER") & String.Empty)
                    job.DistanceRxRightAxis = Val(.Item("AXIS") & String.Empty)
                    job.NearRxRightAdd = Val(.Item("ADD_POWER") & String.Empty)

                    job.PrismRightIn = Val(.Item("PRISM_IN") & String.Empty)
                    job.PrismRightInAxis = .Item("PRISM_IN_AXIS") & String.Empty
                    job.PrismRightUp = Val(.Item("PRISM_UP") & String.Empty)
                    job.PrismRightUpAxis = .Item("PRISM_UP_AXIS") & String.Empty

                    job.VertexPantoRightFittingVertex = Val(.Item("FITTING_VERTEX") & String.Empty)
                    job.VertexPantoRightRefractiveVertex = Val(.Item("REFRACTIVE_VERTEX") & String.Empty)
                    job.VertexPantoRightPantoTilt = Val(.Item("PANTOSCOPIC_TILT") & String.Empty)
                    job.VertexPantoRightWrapAngle = Val(.Item("PANORAMIC_ANGLE") & String.Empty)
                    job.FittingMeasurementsRightMonocularPD = Val(.Item("MONO_PD") & String.Empty)
                    job.FittingMeasurementsRightFittingHeight = Val(.Item("FITTING_HEIGHT") & String.Empty)


                End With
            Else
                job.DistanceRxRightSph = 0
                job.DistanceRxRightCyl = 0
                job.DistanceRxRightAxis = 0
                job.NearRxRightAdd = 0

                job.PrismRightIn = 0
                job.PrismRightInAxis = String.Empty
                job.PrismRightUp = 0
                job.PrismRightUpAxis = String.Empty

                job.VertexPantoRightFittingVertex = 0
                job.VertexPantoRightRefractiveVertex = 0
                job.VertexPantoRightPantoTilt = 0
                job.VertexPantoRightWrapAngle = 0
                job.FittingMeasurementsRightMonocularPD = 0
                job.FittingMeasurementsRightFittingHeight = 0


            End If

            sql = "JOB_NO = '" & .JobNumber & "' AND RL = 'L'"
            If DETJOBM3.Select(sql).Length > 0 Then
                With DETJOBM3.Select(sql)(0)
                    job.DistanceRxLeftSph = Val(.Item("SPHERE") & String.Empty)
                    job.DistanceRxLeftCyl = Val(.Item("CYLINDER") & String.Empty)
                    job.DistanceRxLeftAxis = Val(.Item("AXIS") & String.Empty)
                    job.NearRxLeftAdd = Val(.Item("ADD_POWER") & String.Empty)

                    job.PrismLeftIn = Val(.Item("PRISM_IN") & String.Empty)
                    job.PrismLeftInAxis = .Item("PRISM_IN_AXIS") & String.Empty
                    job.PrismLeftUp = Val(.Item("PRISM_UP") & String.Empty)
                    job.PrismLeftUpAxis = .Item("PRISM_UP_AXIS") & String.Empty

                    job.VertexPantoLeftFittingVertex = Val(.Item("FITTING_VERTEX") & String.Empty)
                    job.VertexPantoLeftRefractiveVertex = Val(.Item("REFRACTIVE_VERTEX") & String.Empty)
                    job.VertexPantoLeftPantoTilt = Val(.Item("PANTOSCOPIC_TILT") & String.Empty)
                    job.VertexPantoLeftWrapAngle = Val(.Item("PANORAMIC_ANGLE") & String.Empty)
                    job.FittingMeasurementsLeftMonocularPD = Val(.Item("MONO_PD") & String.Empty)
                    job.FittingMeasurementsLeftFittingHeight = Val(.Item("FITTING_HEIGHT") & String.Empty)
                End With
            Else
                job.DistanceRxLeftSph = 0
                job.DistanceRxLeftCyl = 0
                job.DistanceRxLeftAxis = 0
                job.NearRxLeftAdd = 0

                job.PrismLeftIn = 0
                job.PrismLeftInAxis = String.Empty
                job.PrismLeftUp = 0
                job.PrismLeftUpAxis = String.Empty

                job.VertexPantoLeftFittingVertex = 0
                job.VertexPantoLeftRefractiveVertex = 0
                job.VertexPantoLeftPantoTilt = 0
                job.VertexPantoLeftWrapAngle = 0
                job.FittingMeasurementsLeftMonocularPD = 0
                job.FittingMeasurementsLeftFittingHeight = 0

            End If

            .IncludePrism = Val(DETJOBM1.Item("RX_PRISM") & String.Empty)

            .FrameDataA = Val(DETJOBM1.Item("FRAME_A_WIDTH") & String.Empty)
            .FrameDataB = Val(DETJOBM1.Item("FRAME_B_HEIGHT") & String.Empty)
            .FrameDataDbl = Val(DETJOBM1.Item("FRAME_DBL_BRIDGE") & String.Empty)
            .FrameDataED = Val(DETJOBM1.Item("FRAME_ED_DIAGONAL") & String.Empty)

            .FrameType = DETJOBM1.Item("FRAME_TYPE_CODE") & String.Empty
            .EdgeThickness = DETJOBM1.Item("REQ_EDGE_THICKNESS") & String.Empty
            .BaseCurve = Val(DETJOBM1.Item("REQ_BASE_CURVE") & String.Empty)

            .SpecialInstructions = DETJOBM1.Item("COMMENT_LAB") & String.Empty
            .Finishedwork = IIf(DETJOBM1.Item("FRAME_STATUS") & String.Empty = "C", "F", DETJOBM1.Item("FRAME_STATUS") & String.Empty)

            .FrameManufacturer = DETJOBM1.Item("FRAME_MFG") & String.Empty
            .FrameModelNumber = DETJOBM1.Item("FRAME_MODEL_NO") & String.Empty
            .FrameColor = DETJOBM1.Item("FRAME_COLOR") & String.Empty
            .ShapePattern = DETJOBM1.Item("SHAPE_PATTERN") & String.Empty
            .FrameSize = DETJOBM1.Item("FRAME_SIZE") & String.Empty
            .TrackingNumber = DETJOBM1.Item("TRACKING_NUMBER") & String.Empty

            .JobStatusCode = JobStatusCode
            If DETJOBM1.Item("ORDR_SOURCE") & String.Empty = "W" Then
                .JobStatusText = GetDataValue("select " & tablePrefix & "detstat2.STATUS_DESC_WEB" & _
                                              " from " & tablePrefix & "detstat1, " & tablePrefix & "detstat2" & _
                                              " where " & tablePrefix & "detstat1.status_code_web = " & tablePrefix & "detstat2.status_code_web" & _
                                              " and status_code = '" & .JobStatusCode & "'" & _
                                              " and " & tablePrefix & "detstat1.status_code_web is not null") & String.Empty

            Else
                .JobStatusText = GetDataValue("select status_desc from " & tablePrefix & "detstat1 where status_code = '" & .JobStatusCode & "'") & String.Empty
            End If

            sql = " SELECT MAX(STATUS_CODE) STATUS_CODE "
            sql &= " FROM " & tablePrefix & "DETJOBM4 "
            sql &= " WHERE INIT_DATE = (SELECT MAX(INIT_DATE) FROM " & tablePrefix & "DETJOBM4 WHERE JOB_NO = '{0}') "
            sql &= " AND JOB_NO = '{1}'"
            sql &= " GROUP BY JOB_NO,INIT_DATE"
            sql = String.Format(sql, .JobNumber, .JobNumber)
            Dim STATUS_CODE As String = GetDataValue(sql) & String.Empty

            .IsJobStatusDelayed = Not (STATUS_CODE = String.Empty OrElse STATUS_CODE <> "800")

            .JobStatusOn = DETJOBM1.Item("INIT_DATE") & String.Empty
            .InvoiceNumber = DETJOBM1.Item("INV_NO") & String.Empty

            If .InvoiceNumber.Length > 0 Then
                sql = " SELECT INV_DATE FROM SOTINVH1 WHRE INV_NO = '{0}' AND INV_TYPE = 'I'"
                sql = String.Format(sql, .InvoiceNumber)
                .InvoicedOn = GetDataValue(sql) & String.Empty
            End If

            .ShipViaDescriptionWeb = DETJOBM1.Item("SHIP_VIA_CODE") & String.Empty
            .CarrierCode = String.Empty
            .CarrierTrackingNumber = String.Empty

            .AddedOn = DateTime.Now
            .UpdatedOn = DateTime.Now
            .UpdatedBy = String.Empty

        End With

        Return job
    End Function

    Public Function GetJob(ByVal jobNumber As String, ByVal UseWebValues As Boolean) As JobValidatorFields

        'UseWorkingDistances -
        ' 1 = AUTOOFFICE designs (use WORKING_DISTANCE field)
        ' 2 = LIFEMADEW designs (use NWD and FWD fields)
        Dim TSql As String = "SELECT " & _
           " 0 DelJobID," & _
           " J1.JOB_NO JobNumber," & _
           " J1.ORDR_DATE SubmittedOn," & _
           " NVL(J1.JOB_NO_ORIG,'') RedoOfJobNumber," & _
           " NVL(J1.REASON_CODE_REDO,'') RedoReason," & _
           " 0 PractitionerId," & _
           " J1.CUST_CODE  AccountNumber," & _
           " NVL(J1.CUST_SHIP_TO_NO, '') ShipToID," & _
           " J1.PATIENT_ID_WEB PatientId," & _
           " CASE WHEN J1.PATIENT_FIRST_NAME IS NULL AND J1.PATIENT_LAST_NAME IS NULL AND J1.TRAY IS NULL THEN NVL(J1.PATIENT_NAME,'') ELSE NVL(J1.PATIENT_FIRST_NAME,'') END PatientFirstName," & _
           " NVL(J1.PATIENT_LAST_NAME,'') PatientLastName," & _
           " 0 SaveNewPatient," & _
           " NVL(J1.TRAY,'') Tray," & _
           " CASE WHEN J1.FINISHED = 'O' THEN 'F' ELSE J1.FINISHED END JobType," & _
           " J1.LENS_ORDER WhichLenses," & _
           " CASE WHEN NVL(PRIVATE_LABEL_DESIGN,'0')='1' THEN D1.LENS_DESIGNER_NAME_INV ELSE J1.LENS_DESIGNER_CODE END Brand," & _
           " J1.LENS_DESIGNER_CODE LensDesigner," & _
           " J1.LENS_DESIGN_CODE LensDesign," & _
           " J1.COLOR_TYPE ColorType," & _
           " J1.MATL_CODE Material," & _
           IIf(UseWebValues, " C1.COLOR_CODE_WEB Color,", " C1.COLOR_CODE Color,") & _
           " NVL(J1.TINT_CODE, '') Tint," & _
           " NVL(J1.TINT_COLOR,'') TintColor," & _
           " NVL(J1.TINT_PCT,0) TintPercent," & _
           " CASE J1.LENS_DESIGN_CODE WHEN 'AUTOOFFICE' THEN 1 WHEN 'LIFEMADEW' THEN 2 ELSE 0 END UseWorkingDistances," & _
           " NVL(J1.WORKING_DISTANCE,-1) WorkingDistancesDistance," & _
           " (NVL(J1.VIEWING_ANGLE,40)/-20)+1 WorkingDistancesViewingAngle," & _
           " NVL(CAST(J1.ACTIVITY AS INT),-1) WorkingDistancesActivity," & _
           " NVL(J1.NWD,0) WorkingDistancesNear," & _
           " NVL(J1.FWD,0) WorkingDistancesFar," & _
           " CASE WHEN NVL(J1.AR_COATING,'0')='0' THEN 0 ELSE 1 END ARCoating," & _
           " CASE WHEN NVL(J1.AR_BACKSIDE_ONLY,'0')='0' THEN 0 ELSE 1 END BacksideOnly," & _
           " CASE WHEN NVL(J1.MIRROR_COATING,'0')='0' THEN 0 ELSE 1 END MirrorCoating," & _
           " NVL(J1.MIRROR_COATING_COLOR,'') MirrorCoatingColor," & _
           " NVL(J1.POLISHING,'0') Polish," & _
           " NVL(J1.UNCUT_TINTABLE,'0') TintableHardcoat," & _
           " NVL(J3R.SPHERE,0) DistanceRxRightSph," & _
           " NVL(J3R.CYLINDER,0) DistanceRxRightCyl," & _
           " NVL(J3R.AXIS,0) DistanceRxRightAxis," & _
           " NVL(J3L.SPHERE,0) DistanceRxLeftSph," & _
           " NVL(J3L.CYLINDER,0) DistanceRxLeftCyl," & _
           " NVL(J3L.AXIS,0) DistanceRxLeftAxis," & _
           " NVL(J3R.ADD_POWER,0) NearRxRightAdd," & _
           " NVL(J3L.ADD_POWER,0) NearRxLeftAdd," & _
           " CASE WHEN NVL(J1.RX_PRISM,'0')='0' THEN 0 ELSE 1 END IncludePrism," & _
           " NVL(J3R.PRISM_IN,0) PrismRightIn," & _
           " NVL(J3R.PRISM_IN_AXIS,'') PrismRightInAxis," & _
           " NVL(J3R.PRISM_UP,0) PrismRightUp," & _
           " NVL(J3R.PRISM_UP_AXIS,'') PrismRightUpAxis," & _
           " NVL(J3L.PRISM_IN,0) PrismLeftIn," & _
           " NVL(J3L.PRISM_IN_AXIS,'') PrismLeftInAxis," & _
           " NVL(J3L.PRISM_UP,0) PrismLeftUp," & _
           " NVL(J3L.PRISM_UP_AXIS,'') PrismleftUpAxis," & _
           " NVL(J3R.FITTING_VERTEX,0) RightFittingVertex," & _
           " NVL(J3R.REFRACTIVE_VERTEX,0) RightRefractiveVertex," & _
           " NVL(J3R.PANTOSCOPIC_TILT,0) RightPantoTilt," & _
           " NVL(J3R.PANORAMIC_ANGLE,0) RightWrapAngle," & _
           " NVL(J3L.FITTING_VERTEX,0) LeftFittingVertex," & _
           " NVL(J3L.REFRACTIVE_VERTEX,0) LeftRefractiveVertex," & _
           " NVL(J3L.PANTOSCOPIC_TILT,0) LeftPantoTilt," & _
           " NVL(J3L.PANORAMIC_ANGLE,0) LeftWrapAngle," & _
           " NVL(J3R.MONO_PD,0) RightMonocularPD," & _
           " NVL(J3R.FITTING_HEIGHT,0) RightFittingHeight," & _
           " NVL(J3L.MONO_PD,0) LeftMonocularPD," & _
           " NVL(J3L.FITTING_HEIGHT,0) LeftFittingHeight," & _
           " NVL(J1.FRAME_A_WIDTH,0) FrameDataA," & _
           " NVL(J1.FRAME_B_HEIGHT,0) FrameDataB," & _
           " NVL(J1.FRAME_DBL_BRIDGE,0) FrameDataDbl," & _
           " NVL(J1.FRAME_ED_DIAGONAL,0) FrameDataED," & _
           " J1.FRAME_TYPE_CODE FrameType," & _
           " NVL(J1.REQ_EDGE_THICKNESS,'') EdgeThickness," & _
           " NVL(J1.REQ_BASE_CURVE,0) BaseCurve," & _
           " CASE WHEN J1.ORDR_SOURCE='W' THEN NVL(J1.COMMENT_LAB,'') ELSE '' END SpecialInstructions," & _
           " CASE WHEN NVL(J1.FRAME_STATUS,'') = 'C' THEN 'F' ELSE NVL(J1.FRAME_STATUS,'') END FinishedWork," & _
           " NVL(J1.FRAME_MFG,'') FrameManufacturer," & _
           " NVL(J1.FRAME_MODEL_NO,'') FrameModelNumber," & _
           " NVL(J1.FRAME_COLOR,'') FrameColor," & _
           " NVL(J1.SHAPE_PATTERN,'') ShapePattern," & _
           " NVL(J1.FRAME_SIZE,'') FrameSize," & _
           " NVL(J1.TRACKING_NUMBER,'') TrackingNumber," & _
           " NVL(S2.STATUS_TYPE_WEB,'') JobStatusCode," & _
           " NVL(S2.STATUS_DESC_WEB,'') JobStatusText," & _
           " CASE WHEN J4D.STATUS_CODE IS NULL OR J4D.STATUS_CODE <> '800' THEN 0 ELSE 1 END IsJobStatusDelayed," & _
           " J4.INIT_DATE JobStatusOn," & _
           " J1.INV_DATE InvoicedOn," & _
           " NVL(J1.INV_NO,'') InvoiceNumber," & _
           " NVL(SV1.SHIP_VIA_DESC_WEB,'') ShipViaDescription," & _
           " NVL(SR1.CARRIER_CODE,'') CarrierCode," & _
           " NVL(SI1.SHIP_REF,'') CarrierTrackingNumber," & _
           " SYSDATE AddedOn," & _
           " SYSDATE UpdatedOn," & _
           " '00000000-0000-0000-0000-000000000000' UpdatedBy " & _
           " FROM" & _
           " " & tablePrefix & "DETJOBM1 J1" & _
           "   LEFT JOIN" & _
           " " & tablePrefix & "DETJOBM3 J3R ON (J1.JOB_NO=J3R.JOB_NO AND J3R.RL='R')" & _
           "   LEFT JOIN" & _
           " " & tablePrefix & "DETJOBM3 J3L ON (J1.JOB_NO=J3L.JOB_NO AND J3L.RL='L')" & _
           "   LEFT JOIN" & _
           " (SELECT JOB_NO, MAX(STATUS_CODE) STATUS_CODE, INIT_DATE FROM " & tablePrefix & "DETJOBM4 DJ4 WHERE INIT_DATE=(SELECT MAX(INIT_DATE) FROM " & tablePrefix & "DETJOBM4 WHERE JOB_NO=DJ4.JOB_NO) GROUP BY JOB_NO,INIT_DATE) J4 ON (J1.JOB_NO=J4.JOB_NO)" & _
           "   LEFT JOIN" & _
           " (SELECT JOB_NO, MAX(STATUS_CODE) STATUS_CODE FROM " & tablePrefix & "DETJOBM4 GROUP BY JOB_NO) J4D ON (J4.JOB_NO=J4D.JOB_NO AND J4D.STATUS_CODE BETWEEN '800' AND '920')" & _
           "   LEFT JOIN" & _
           " " & tablePrefix & "DETSTAT1 S1 ON (J4.STATUS_CODE=S1.STATUS_CODE)" & _
           "   LEFT JOIN" & _
           " " & tablePrefix & "DETSTAT2 S2 ON (S1.STATUS_CODE_WEB=S2.STATUS_CODE_WEB)" & _
           "   LEFT JOIN" & _
           " " & tablePrefix & "SOTINVH1 SI1 ON (J1.INV_NO=SI1.INV_NO)" & _
           "   LEFT JOIN" & _
           " " & tablePrefix & "SOTSVIA1 SV1 ON (SI1.SHIP_VIA_CODE=SV1.SHIP_VIA_CODE)" & _
           "   LEFT JOIN" & _
           " " & tablePrefix & "SOTROUT1 SR1 ON (SV1.ROUTE_CODE=SR1.ROUTE_CODE)" & _
           "   JOIN" & _
           " " & tablePrefix & "DETDSGN1 D1 ON (J1.LENS_DESIGN_CODE=D1.LENS_DESIGN_CODE)" & _
           "   JOIN" & _
           " " & tablePrefix & "DETCOLR1 C1 ON (J1.COLOR_CODE=C1.COLOR_CODE)" & _
           " WHERE" & _
           " J1.JOB_NO={0}"

        Dim sql = String.Format(TSql, jobNumber)

        Dim dt = GetDataTable(sql)

        Select Case dt.Rows.Count
            Case Is = 0
                Throw New Exception("Cannot locate Job No: " & jobNumber)

            Case Is > 1
                Throw New Exception("Multi Jobs returned for Job: " & jobNumber)
        End Select

        Dim job As New JobValidatorFields
        Dim row As DataRow = dt.rows(0)

        With job
            .DelJobId = row.Item("DelJobId")
            .JobNumber = row.Item("JobNumber")
            .SubmittedOn = row.Item("SubmittedOn")
            .RedoOfJobNumber = row.Item("RedoOfJobNumber")
            .ReasonCodeRedo = row.Item("ReasonCodeRedo")
            .PractitionerID = row.Item("PractitionerID")
            .AccountNumber = row.Item("AccountNumber")
            .ShipToID = row.Item("ShipToID")
            .PatientId = row.Item("PatientId")
            .PatientFirstName = row.Item("PatientFirstName")
            .PatientLastName = row.Item("PatientLastName")
            .SaveNewPatient = row.Item("SaveNewPatient")
            .Tray = row.Item("Tray")
            .JobType = row.Item("JobType")
            .WhichLenses = row.Item("WhichLenses")
            .Brand = row.Item("Brand")
            .LensDesigner = row.Item("LensDesigner")
            .LensDesign = row.Item("LensDesign")
            .ColorType = row.Item("ColorType")
            .Material = row.Item("Material")
            .Color = row.Item("Color")
            .Tint = row.Item("Tint")
            .TintColor = row.Item("TintColor")
            .TintPercent = row.Item("TintPercent")
            .UseWorkingDistances = row.Item("UseWorkingDistances")
            .WorkingDistancesDistance = row.Item("WorkingDistancesDistance")
            .WorkingDistancesViewingAngle = row.Item("WorkingDistancesViewingAngle")
            .WorkingDistancesActivity = row.Item("WorkingDistancesActivity")
            .WorkingDistancesNear = row.Item("WorkingDistancesNear")
            .WorkingDistancesFar = row.Item("WorkingDistancesFar")
            .ArCoating = row.Item("ArCoating")
            .BacksideOnly = row.Item("BacksideOnly")
            .MirrorCoating = row.Item("MirrorCoating")
            .MirrorCoatingColor = row.Item("MirrorCoatingColor")
            .Polish = row.Item("Polish")
            .TintableHardcoat = row.Item("TintableHardcoat")

            .DistanceRxRightSph = row.Item("DistanceRxRightSph")
            .DistanceRxRightCyl = row.Item("DistanceRxRightCyl")
            .DistanceRxRightAxis = row.Item("DistanceRxRightAxis")

            .DistanceRxLeftSph = row.Item("DistanceRxLeftSph")
            .DistanceRxLeftCyl = row.Item("DistanceRxLeftCyl")
            .DistanceRxLeftAxis = row.Item("DistanceRxLeftAxis")

            .NearRxRightAdd = row.Item("NearRxRightAdd")
            .NearRxLeftAdd = row.Item("NearRxLeftAdd")

            .IncludePrism = row.Item("IncludePrism")
            .PrismRightIn = row.Item("PrismRightIn")
            .PrismRightInAxis = row.Item("PrismRightInAxis")
            .PrismRightUp = row.Item("PrismRightUp")
            .PrismRightUpAxis = row.Item("PrismRightUpAxis")
            .PrismLeftIn = row.Item("PrismLeftIn")
            .PrismLeftInAxis = row.Item("PrismLeftInAxis")
            .PrismLeftUp = row.Item("PrismLeftUp")
            .PrismLeftUpAxis = row.Item("PrismLeftUpAxis")

            .VertexPantoRightFittingVertex = row.Item("RightFittingVertex")
            .VertexPantoRightRefractiveVertex = row.Item("RightRefractiveVertex")
            .VertexPantoRightPantoTilt = row.Item("RightPantoTilt")
            .VertexPantoRightWrapAngle = row.Item("RightWrapAngle")

            .VertexPantoLeftFittingVertex = row.Item("LeftFittingVertex")
            .VertexPantoLeftRefractiveVertex = row.Item("LeftRefractiveVertex")
            .VertexPantoLeftPantoTilt = row.Item("LeftPantoTilt")
            .VertexPantoLeftWrapAngle = row.Item("LeftWrapAngle")

            .FittingMeasurementsRightMonocularPD = row.Item("RightMonocularPD")
            .FittingMeasurementsRightFittingHeight = row.Item("RightFittingHeight")
            .FittingMeasurementsLeftMonocularPD = row.Item("LeftMonocularPD")
            .FittingMeasurementsLeftFittingHeight = row.Item("LeftFittingHeight")

            .FrameDataA = row.Item("FrameDataA")
            .FrameDataB = row.Item("FrameDataB")
            .FrameDataDbl = row.Item("FrameDataDbl")
            .FrameDataED = row.Item("FrameDataED")

            .FrameType = row.Item("FrameType")
            .EdgeThickness = row.Item("EdgeThickness")
            .BaseCurve = row.Item("BaseCurve")

            .SpecialInstructions = row.Item("SpecialInstructions")
            .Finishedwork = row.Item("Finishedwork")
            .FrameManufacturer = row.Item("FrameManufacturer")
            .FrameModelNumber = row.Item("FrameModelNumber")
            .FrameColor = row.Item("FrameColor")
            .ShapePattern = row.Item("ShapePattern")
            .FrameSize = row.Item("FrameSize")
            .TrackingNumber = row.Item("TrackingNumber")
            .JobStatusCode = row.Item("JobStatusCode")
            .JobStatusText = row.Item("JobStatusText")
            .IsJobStatusDelayed = row.Item("IsJobStatusDelayed")
            .JobStatusOn = row.Item("JobStatusOn")
            .InvoicedOn = row.Item("InvoicedOn")
            .InvoiceNumber = row.Item("InvoiceNumber")
            .ShipViaDescriptionWeb = row.Item("ShipViaDescriptionWeb")
            .CarrierCode = row.Item("CarrierCode")
            .CarrierTrackingNumber = row.Item("CarrierTrackingNumber")
            .AddedOn = row.Item("AddedOn")
            .UpdatedOn = row.Item("UpdatedOn")
            .UpdatedBy = row.Item("UpdatedBy")

        End With

        ' Clear "web-only" fields
        'job.DelJobID = 0  ' Can't change PK value, but doesn't matter... DelJobID a meaningless value for GetJob
        job.PractitionerID = 0
        job.SaveNewPatient = False
        job.AddedOn = DateTime.MinValue
        job.UpdatedBy = String.Empty
        job.UpdatedOn = DateTime.MinValue

        Return job

    End Function

#End Region


    ' Entire section commented out by Ed
#Region " Format For Display "

    'Public Shared Function GetCodeDescriptions(ByVal connInfo As ConnectionInfo, ByVal job As DelJobDS.DelJobRow) As DelJobCodeDescriptions

    '    Dim result = New DelJobCodeDescriptions()

    '    With result

    '        .RedoReason = GetRedoReasonDescription(connInfo, job.RedoReason)
    '        .JobType = GetJobTypeDescription(job.JobType)
    '        .WhichLenses = GetWhichLensesDescription(job.WhichLenses)
    '        .Brand = GetBrandDescription(connInfo, job.Brand)
    '        .LensDesign = GetLensDesignDescription(connInfo, job.Brand, job.LensDesign)
    '        .ColorType = GetColorTypeDescription(job.ColorType)
    '        .Material = GetMaterialDescription(connInfo, job.Brand, job.LensDesign, job.ColorType, job.Material)
    '        .Color = GetColorDescription(connInfo, job.Brand, job.LensDesign, job.ColorType, job.Color)
    '        .Tint = GetTintDescription(connInfo, job.Tint)
    '        .FrameType = GetFrameTypeDescription(connInfo, job.FrameType)
    '        .PrismRightInAxis = GetPrismAxisDescription(job.PrismRightInAxis)
    '        .PrismRightUpAxis = GetPrismAxisDescription(job.PrismRightUpAxis)
    '        .PrismLeftInAxis = GetPrismAxisDescription(job.PrismLeftInAxis)
    '        .PrismLeftUpAxis = GetPrismAxisDescription(job.PrismLeftUpAxis)
    '        .WorkingDistancesViewingAngle = GetWorkingDistancesViewingAngleDescription(job.WorkingDistancesViewingAngle)
    '        .WorkingDistancesActivity = GetWorkingDistancesActivityDescription(job.WorkingDistancesActivity)
    '        .FinishedWork = GetFinishedWorkDescription(job.FinishedWork)

    '    End With

    '    Return result

    'End Function

    'Public Shared Function GetRedoReasonDescription(ByVal connInfo As ConnectionInfo, ByVal redoReasonCode As String) As String

    '    Dim dt = GetRedoReasons(connInfo)

    '    Dim desc = _
    '     (From row In dt.AsEnumerable() _
    '     Where row.Field(Of String)("REASON_CODE") = redoReasonCode _
    '     Select row.Field(Of String)("REASON_DESC")).SingleOrDefault()

    '    If desc Is Nothing Then
    '        desc = redoReasonCode
    '    End If

    '    Return desc

    'End Function

    'Public Shared Function GetJobTypeDescription(ByVal jobTypeCode As String) As String

    '    Dim desc = jobTypeCode

    '    Select Case jobTypeCode
    '        Case "U"
    '            desc = "Uncut"
    '        Case "F"
    '            desc = "Finished"
    '    End Select

    '    Return desc

    'End Function

    'Public Shared Function GetWhichLensesDescription(ByVal whichLensesCode As String) As String

    '    Dim desc = whichLensesCode

    '    Select Case whichLensesCode
    '        Case "B"
    '            desc = "Both Lenses"
    '        Case "R"
    '            desc = "Right Lens Only"
    '        Case "L"
    '            desc = "Left Lens Only"
    '    End Select

    '    Return desc

    'End Function

    'Public Shared Function GetBrandDescription(ByVal connInfo As ConnectionInfo, ByVal brandCode As String) As String

    '    Dim dt = GetBrands(connInfo)

    '    Dim desc = _
    '     (From row In dt.AsEnumerable() _
    '     Where row.Field(Of String)("LENS_DESIGNER_CODE") = brandCode _
    '     Select row.Field(Of String)("LENS_DESIGNER_NAME")).SingleOrDefault()

    '    If desc Is Nothing Then
    '        desc = brandCode
    '    End If

    '    Return desc

    'End Function

    'Public Shared Function GetLensDesignDescription(ByVal connInfo As ConnectionInfo, ByVal brandCode As String, ByVal lensDesignCode As String) As String

    '    Dim dt = GetLensDesigns(connInfo, brandCode)

    '    Dim desc = _
    '     (From row In dt.AsEnumerable() _
    '     Where row.Field(Of String)("LENS_DESIGN_CODE") = lensDesignCode _
    '     Select row.Field(Of String)("LENS_DESIGN_NAME_WEB")).SingleOrDefault()

    '    If desc Is Nothing Then
    '        desc = brandCode + " " + lensDesignCode
    '    End If

    '    Return desc

    'End Function

    'Public Shared Function GetColorTypeDescription(ByVal colorTypeCode As String) As String

    '    Dim desc = colorTypeCode

    '    Select Case colorTypeCode
    '        Case "C"
    '            desc = "Clear"
    '        Case "P"
    '            desc = "Polarized"
    '        Case "T"
    '            desc = "Transitions"
    '    End Select

    '    Return desc

    'End Function

    'Public Shared Function GetMaterialDescription(ByVal connInfo As ConnectionInfo, ByVal brandCode As String, ByVal lensDesignCode As String, ByVal colorTypeCode As String, ByVal materialCode As String) As String

    '    Dim dt = GetMaterials(connInfo, brandCode, lensDesignCode, colorTypeCode)

    '    Dim desc = _
    '     (From row In dt.AsEnumerable() _
    '     Where row.Field(Of String)("MATL_CODE") = materialCode _
    '     Select row.Field(Of String)("MATL_DESC")).SingleOrDefault()

    '    If desc Is Nothing Then
    '        desc = materialCode
    '    End If

    '    Return desc

    'End Function

    'Public Shared Function GetColorDescription(ByVal connInfo As ConnectionInfo, ByVal brandCode As String, ByVal lensDesignCode As String, ByVal colorTypeCode As String, ByVal colorCode As String) As String

    '    Dim dt = GetColors(connInfo, brandCode, lensDesignCode, colorTypeCode)

    '    Dim desc = _
    '     (From row In dt.AsEnumerable() _
    '     Where row.Field(Of String)("COLOR_CODE") = colorCode _
    '     Select row.Field(Of String)("COLOR_DESC")).SingleOrDefault()

    '    If desc Is Nothing Then
    '        desc = colorCode
    '    End If

    '    Return desc

    'End Function

    'Public Shared Function GetTintDescription(ByVal connInfo As ConnectionInfo, ByVal tintCode As String) As String

    '    Dim dt = GetTints(connInfo)

    '    Dim desc = _
    '     (From row In dt.AsEnumerable() _
    '     Where row.Field(Of String)("TINT_CODE") = tintCode _
    '     Select row.Field(Of String)("TINT_DESC")).SingleOrDefault()

    '    If desc Is Nothing Then
    '        desc = tintCode
    '    End If

    '    Return desc

    'End Function

    'Public Shared Function GetFrameTypeDescription(ByVal connInfo As ConnectionInfo, ByVal tintCode As String) As String

    '    Dim dt = GetFrameTypes(connInfo)

    '    Dim desc = _
    '     (From row In dt.AsEnumerable() _
    '     Where row.Field(Of String)("FRAME_TYPE_CODE") = tintCode _
    '     Select row.Field(Of String)("FRAME_TYPE_DESC")).SingleOrDefault()

    '    If desc Is Nothing Then
    '        desc = tintCode
    '    End If

    '    Return desc

    'End Function

    'Public Shared Function GetPrismAxisDescription(ByVal axisCode As String) As String

    '    Dim desc = axisCode

    '    Select Case axisCode
    '        Case "I"
    '            desc = "In"
    '        Case "O"
    '            desc = "Out"
    '        Case "U"
    '            desc = "Up"
    '        Case "D"
    '            desc = "Down"
    '    End Select

    '    Return desc

    'End Function

    'Public Shared Function GetWorkingDistancesViewingAngleDescription(ByVal viewingAngleCode As String) As String

    '    Dim desc = viewingAngleCode

    '    Select Case viewingAngleCode
    '        Case "0"
    '            desc = "Above"
    '        Case "1"
    '            desc = "At Eye-Level"
    '        Case "2"
    '            desc = "Below"
    '    End Select

    '    Return desc

    'End Function

    'Public Shared Function GetWorkingDistancesActivityDescription(ByVal activityCode As String) As String

    '    Dim desc = activityCode

    '    Select Case activityCode
    '        Case "0"
    '            desc = "Static"
    '        Case "1"
    '            desc = "Active"
    '    End Select

    '    Return desc

    'End Function

    'Public Shared Function GetFinishedWorkDescription(ByVal finishedWorkCode As String) As String

    '    Dim desc = finishedWorkCode

    '    Select Case finishedWorkCode
    '        Case "F"
    '            desc = "Frame To Come"
    '        Case "S"
    '            desc = "Shape/Edge"
    '        Case "E"
    '            desc = "Enclosed"
    '    End Select

    '    Return desc

    'End Function

#End Region

#Region " Redo Functionality "

    Public Shared Function CanRedo(ByVal jobNumber As String) As Boolean

        Dim sqlIsJobInvoiced As String = "SELECT JOB_STATUS FROM " & tablePrefix & "DETJOBM1 WHERE JOB_NO = '{0}'"

        Dim jobRow = GetDataTable(String.Format(sqlIsJobInvoiced, jobNumber))

        'Can't redo jobs that have not been invoiced
        If jobRow Is Nothing Then Return False
        If jobRow.Item("JOB_STATUS") Is DBNull.Value OrElse jobRow.Item("JOB_STATUS") <> "F" Then Return False

        'can't redo job that uses a non-web design
        Dim sqlCheckDesign As String = "SELECT ISNULL(LENS_DESIGN_ACTIVATE_FOR_WEB,'0') ISACTIVE FROM " & tablePrefix & "DETDSGN1 D1 JOIN " & tablePrefix & "DETJOBM1 J1 ON (D1.LENS_DESIGN_CODE=J1.LENS_DESIGN_CODE) WHERE J1.JOB_NO='{0}'"
        Dim designRow = GetDataTable(String.Format(sqlCheckDesign, jobNumber))

        If designRow.Item("ISACTIVE") <> "1" Then
            Return False
        End If

        Dim sqlGetRootJob As String = "" & _
           "WITH " _
         & " JobRedoUp (JOB_NO,JOB_NO_ORIG)" _
         & " as" _
         & " (" _
         & " SELECT JOB_NO,JOB_NO_ORIG" _
         & " FROM " & tablePrefix & "DETJOBM1" _
         & " where JOB_NO='{0}'" _
         & " UNION ALL" _
         & " SELECT a1.JOB_NO, A1.JOB_NO_ORIG" _
         & " FROM " & tablePrefix & "DETJOBM1 a1" _
         & " INNER JOIN JobRedoUp" _
         & " on a1.JOB_NO = JobRedoUp.JOB_NO_ORIG" _
         & " )" _
         & " SELECT" _
         & " JOB_NO" _
         & " FROM" _
         & " JobRedoUp" _
         & " WHERE JOB_NO_ORIG IS NULL"
        Dim rootRow = GetDataTable(String.Format(sqlGetRootJob, jobNumber))

        If rootRow Is Nothing Then
            Return True
        End If

        Dim sqlGetJobCount As String = "" & _
          "WITH " _
           & " JobRedoDown (JOB_NO,JOB_NO_ORIG)" _
           & " as" _
           & " (" _
           & " SELECT JOB_NO,JOB_NO_ORIG" _
           & " FROM " & tablePrefix & "DETJOBM1" _
           & " where JOB_NO='{0}'" _
           & " UNION ALL" _
           & " SELECT a1.JOB_NO, A1.JOB_NO_ORIG" _
           & " FROM " & tablePrefix & "DETJOBM1 a1" _
           & " INNER JOIN JobRedoDown" _
           & " on a1.JOB_NO_ORIG = JobRedoDown.JOB_NO" _
           & " )" _
           & " SELECT" _
           & " COUNT(*)" _
           & " FROM" _
           & " JobRedoDown"

        Dim jobCount As DataRow = GetDataTable(String.Format(sqlGetJobCount, rootRow.Item(0))).Rows(0)

        If jobCount IsNot Nothing Then
            If Convert.ToInt32(jobCount.Item(0)) > 5 Then
                Return False
            End If
        End If

        Return True

    End Function

#End Region

End Class

