Public Class DECVALID

    ''' <summary>
    ''' The dst passed to this class must contain the following datatables loaded with data
    ''' DETJOBM1 - Job header table
    ''' DETJOBM2 - Lens pricing for the Job
    ''' DETJOBM3 - Lens Attributes for the Job 
    ''' 
    ''' SOTORDR5 - Sales Order Ship Address entrieds for the Job
    ''' SOTSVIA1 - Entire Ship Via Table or the Selected Ship Via
    ''' 
    ''' DETFRAM1 - Entire Frame Master table or the Frame on the Job
    ''' DETJOBT1 - Filled with data from the Original Job No DETJOBM1.JOB_NO_ORIG
    ''' DETMATL1 - Entire Material Master or the Material on the Job
    ''' DETPROM1 - Promo Code entry for the PROMO_CODE on the Job
    ''' DETPARM1 - Entire parameter table
    ''' 
    ''' ARTCUST1 - Customer Master for the Customer on the joB
    ''' ARTCUST2 - Ship To master for the ship to on the Job
    '''
    ''' DETDSGN0 - Entire table or Selected LENS_DESIGNER_CODE
    ''' DETDSGN1 - Entire table or Selected LENS_DESIGN_CODE
    ''' DETDSGN2 - Entire table or Selected LENS_DESIGN_CODE
    ''' DETDSGN3 - Entire table or Selected LENS_DESIGN_CODE
    ''' DETDSGN4 - Entire table or Selected LENS_DESIGNER_CODE or Selected LENS_DESIGNER_CODE / MATL_CODE
    ''' DETDSGN6 - Entire table or Selected LENS_DESIGNER_CODE or Selected LENS_DESIGNER_CODE / MATL_CODE
    ''' 
    ''' DETDRIL1 - For the selected Pattern No
    ''' DETCFRM1 - For the selected Frame No
    ''' DETTRCP2 - For XREF_TYPE = 'FPC' and XREF_CODE = the Selected FPC
    ''' 
    ''' DETJOBMP - JOB_NO, CUST_CODE, PROMO_CODE, JOB_STATUS from DETJOBM1 for jobs with the Promo Code (PROM_CODE) on this Job
    ''' 
    ''' </summary>
    ''' <remarks></remarks>

    Private dst As DataSet = New DataSet

    Private BlankSelectionException As New Dictionary(Of String, String)
    Private prevent_blank_selection As Boolean = False
    Private RL_ISSUES As New Dictionary(Of String, String)

    Public ErrorMasterCodeList As Hashtable = New Hashtable
    Private sErrorList As Hashtable = New Hashtable

    Private rowDETJOBM1 As DataRow = Nothing
    Private rowDETJOBM2 As DataRow = Nothing
    Private rowDETJOBM3 As DataRow = Nothing

    Private rowSOTORDR5 As DataRow = Nothing
    Private rowSOTSVIA1 As DataRow = Nothing

    Private rowDETFRAM1 As DataRow = Nothing
    Private rowDETJOBT1 As DataRow = Nothing
    Private rowDETMATL1 As DataRow = Nothing
    Private rowDETPROM1 As DataRow = Nothing
    Private rowDETPARM1 As DataRow = Nothing

    Private rowARTCUST1 As DataRow = Nothing
    Private rowARTCUST2 As DataRow = Nothing

    Private rowDETDSGN0 As DataRow = Nothing
    Private rowDETDSGN1 As DataRow = Nothing
    Private rowDETDSGN2 As DataRow = Nothing
    Private rowDETDSGN3 As DataRow = Nothing
    Private rowDETDSGN4 As DataRow = Nothing
    Private rowDETDSGN6 As DataRow = Nothing

    Private rowDETDRIL1 As DataRow = Nothing
    Private rowDETCFRM1 As DataRow = Nothing
    Private rowDETTRCP2 As DataRow = Nothing
    Private rowDETJOBMP As DataRow = Nothing

#Region "Instantiate Class"

    Public Sub New()
        InitializeObjects()
    End Sub

    Public Sub New(ByRef dDst As DataSet)
        InitializeObjects()
        dst = dDst.Copy
    End Sub

#End Region

#Region "Class Properties"

    ''' <summary>
    ''' Gets a list of errors / validation check errors
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetErrors() As Hashtable
        Get
            Return sErrorList
        End Get
    End Property

#End Region

#Region "Class Procedures"

    Public Sub MapTable(ByRef dTable As DataTable)
        If dst.Tables.Contains(dTable.TableName) Then
            dst.Tables.Remove(dTable.TableName)
        End If

        dst.Tables.Add(dTable)
    End Sub

    Private Sub InitializeObjects()

        dst = New DataSet
        CreateErrorCodeList()
        InitializeDatarows()
    End Sub

    Private Sub InitializeDatarows()
        rowDETJOBM1 = Nothing
        rowDETJOBM2 = Nothing
        rowDETJOBM3 = Nothing

        rowSOTORDR5 = Nothing
        rowSOTSVIA1 = Nothing

        rowDETFRAM1 = Nothing
        rowDETJOBT1 = Nothing
        rowDETMATL1 = Nothing
        rowDETPROM1 = Nothing
        rowDETPARM1 = Nothing

        rowARTCUST1 = Nothing
        rowARTCUST2 = Nothing

        rowDETDSGN0 = Nothing
        rowDETDSGN1 = Nothing
        rowDETDSGN2 = Nothing
        rowDETDSGN3 = Nothing
        rowDETDSGN4 = Nothing
        rowDETDSGN6 = Nothing

        rowDETDRIL1 = Nothing
        rowDETCFRM1 = Nothing
        rowDETTRCP2 = Nothing
        rowDETJOBMP = Nothing
    End Sub

    Private Sub CreateErrorCodeList()

        ErrorMasterCodeList = New Hashtable

        ErrorMasterCodeList.Add("A", "Job not found in job header")
        ErrorMasterCodeList.Add("B", "Invalid Customer Code")
        ErrorMasterCodeList.Add("C", "Customer requires a Ship To")
        ErrorMasterCodeList.Add("D", "Invalid Ship To")
        ErrorMasterCodeList.Add("E", "Ship-To Name and Address (including City, State & Zip Code) is Required")
        ErrorMasterCodeList.Add("F", "Frame Type (None) is reserved for Jobs with certain Finishing Services only (like Coating)")
        ErrorMasterCodeList.Add("G", "No LMATID on file for Designer, Material & Color Specified")
        ErrorMasterCodeList.Add("H", "No record in Design Table for Material Specified")
        ErrorMasterCodeList.Add("I", "No record in Design Table for Material and Color Specified")
        ErrorMasterCodeList.Add("J", "The record in Design Table for Material and Color Specified is not active.")
        ErrorMasterCodeList.Add("K", "Issues with R Rx Data")
        ErrorMasterCodeList.Add("L", "Issues with L Rx Data")
        ErrorMasterCodeList.Add("M", "Invalid Design")
        ErrorMasterCodeList.Add("N", "You Must Select Misc-Finish when using Customer Supplied Blanks")
        ErrorMasterCodeList.Add("O", "Design chosen is not Active (for Job Order Entry)")
        ErrorMasterCodeList.Add("P", "Invalid OPC for eye - (No Record in Blank Selection Chart)")
        ErrorMasterCodeList.Add("Q", "No Corridor Selected")
        ErrorMasterCodeList.Add("R", "Invalid Design / Corridor Length Combination")
        ErrorMasterCodeList.Add("S", "Minimum Fitting Height for Corridor Selected")
        ErrorMasterCodeList.Add("T", "Invalid Frame Type")
        ErrorMasterCodeList.Add("U", "Specifying a Tint (or None) is Mandatory")
        ErrorMasterCodeList.Add("V", "Entry of a Color and a Color % is Mandatory")
        ErrorMasterCodeList.Add("W", "Invalid Drill Pattern")
        ErrorMasterCodeList.Add("X", "Invlaid Custom Frame")
        ErrorMasterCodeList.Add("Y", "Custom Frame Job cannot use a Drill Pattern")
        ErrorMasterCodeList.Add("Z", "You Must Specify a Source for the Trace")
        ErrorMasterCodeList.Add("AA", "Finished Work cannot have No Trace")
        ErrorMasterCodeList.Add("AB", "Tracing Points Database does Not have a Trace for FPC")
        ErrorMasterCodeList.Add("AC", "Custom Frame Job options apply to Finished Jobs only")
        ErrorMasterCodeList.Add("AD", "Drill Pattern options apply to Finished Jobs only")
        ErrorMasterCodeList.Add("AE", "No 'Other' Finishing Services Selected")
        ErrorMasterCodeList.Add("AF", "Cannot find record of Trace from Original Job")
        ErrorMasterCodeList.Add("AG", "Cannot Determine FPC from Frame Data Provided")
        ErrorMasterCodeList.Add("AH", "Frame Type is not NONE, so Values for A, B, ED & DBL are Mandatory")
        ErrorMasterCodeList.Add("AI", "Missing Parameter table DETPARM1")
        ErrorMasterCodeList.Add("AJ", "Max Values Frame Measurements")
        ErrorMasterCodeList.Add("AK", "Only Re-Do Jobs may use a Trace from an Original Order")
        ErrorMasterCodeList.Add("AL", "No Trace Data on file for Original Job")
        ErrorMasterCodeList.Add("AM", "Cannot access Frame Master table")
        ErrorMasterCodeList.Add("AN", "Cannot access Material Master table")
        ErrorMasterCodeList.Add("AO", "Material is not set up for Rimless (Drill-Mount)")
        ErrorMasterCodeList.Add("AP", "Customer PO No is Mandatory")
        ErrorMasterCodeList.Add("AQ", "Customer PO must begin with DEL")
        ErrorMasterCodeList.Add("AR", "Patient Name is Mandatory")
        ErrorMasterCodeList.Add("AS", "Invalid Ship Via Code")
        ErrorMasterCodeList.Add("AT", "Invalid Cylinder/Axis Combination")
        ErrorMasterCodeList.Add("AU", "Prism Values not in synch with In/Out and Up/Down")
        ErrorMasterCodeList.Add("AV", "Invalid Fitting Height for")
        ErrorMasterCodeList.Add("AW", "Progressive Designs must have a non-zero Add Power")
        ErrorMasterCodeList.Add("AX", "Reason Required for Holding Production")
        ErrorMasterCodeList.Add("AY", "Reason for Holding Production should be Blank")
        ErrorMasterCodeList.Add("AZ", "Reason Required for Holding Billing")
        ErrorMasterCodeList.Add("BA", "Reason for Holding Billing should be Blank")
        ErrorMasterCodeList.Add("BB", "Right + Left Mono PD values should be within Range of 45 thru 84")
        ErrorMasterCodeList.Add("BC", "Lens Cylinder must be the same Sign.")
        ErrorMasterCodeList.Add("BD", "Spheres have Opposite Signs")
        ErrorMasterCodeList.Add("BE", "Invalid Promotion Code")
        ErrorMasterCodeList.Add("BF", "Cannot locate promotion lookup table")
        ErrorMasterCodeList.Add("BG", "The selected Promotion is a One Time Only promotion and it has been used")
        ErrorMasterCodeList.Add("BH", "Job Price may not be Negative")
        'ErrorMasterCodeList.Add("", "")

        ErrorMasterCodeList.Add("WA", "Minimum Fitting Height for Corridor Selected")
        ErrorMasterCodeList.Add("WB", "Frame Type is not NONE, Value for DBL is 0")
        ErrorMasterCodeList.Add("WC", "Use Trace from the Original Job")

        ErrorMasterCodeList.Add("SE", "System Error")

    End Sub

    Public Function ValidateJobData(ByVal jobNumber As String, ByVal displayWarningMessages As Boolean) As Boolean

        Try
            Dim JOB_NO_ORIG As String = String.Empty
            Dim zMsg As String = String.Empty
            Dim sql As String = String.Empty

            InitializeDatarows()
            sErrorList.Clear()
            RL_ISSUES.Clear()

            If dst.Tables("DETJOBM1").Select("JOB_NO = '" & jobNumber & "'").Length = 0 Then
                sErrorList.Add("A", ErrorMasterCodeList("A"))
                Return False
            End If

            If dst.Tables.Contains("DETPARM1") AndAlso dst.Tables("DETPARM1").Rows.Count > 0 Then
                rowDETPARM1 = dst.Tables("DETPARM1").Rows(0)
            End If

            rowDETJOBM1 = dst.Tables("DETJOBM1").Select("JOB_NO = '" & jobNumber & "'")(0)
            Dim CUST_CODE As String = rowDETJOBM1.Item("CUST_CODE") & String.Empty
            Dim CUST_SHIP_TO_NO As String = rowDETJOBM1.Item("CUST_CODE") & String.Empty
            Dim ORDR_NO As String = rowDETJOBM1.Item("ORDR_NO") & String.Empty

            JOB_NO_ORIG = rowDETJOBM1("JOB_NO_ORIG") & String.Empty
            JOB_NO_ORIG = JOB_NO_ORIG.Trim
            If JOB_NO_ORIG.Length > 0 Then
                If dst.Tables.Contains("DETJOBT1") Then
                    If dst.Tables("DETJOBT1").Select("JOB_NO = '" & JOB_NO_ORIG & String.Empty & "'").Length = 0 Then
                        rowDETJOBT1 = dst.Tables("DETJOBT1").Select("JOB_NO = '" & JOB_NO_ORIG & String.Empty & "'")(0)
                    End If
                End If
            End If

            If dst.Tables.Contains("DETFRAM1") Then
                If dst.Tables("DETFRAM1").Select("FRAME_TYPE_CODE = '" & rowDETJOBM1("FRAME_TYPE_CODE") & String.Empty & "'").Length = 0 Then
                    rowDETFRAM1 = dst.Tables("DETFRAM1").Select("FRAME_TYPE_CODE = '" & rowDETJOBM1("FRAME_TYPE_CODE") & String.Empty & "'")(0)
                End If
            End If

            If dst.Tables.Contains("DETMATL1") Then
                If dst.Tables("DETMATL1").Select("MATL_CODE = '" & rowDETJOBM1("MATL_CODE") & String.Empty & "'").Length = 0 Then
                    rowDETMATL1 = dst.Tables("DETMATL1").Select("MATL_CODE = '" & rowDETJOBM1("MATL_CODE") & String.Empty & "'")(0)
                End If
            End If

            If dst.Tables.Contains("DETPROM1") AndAlso rowDETJOBM1("PROMO_CODE") & String.Empty <> String.Empty Then
                If dst.Tables("DETPROM1").Select("PROMO_CODE = '" & rowDETJOBM1("PROMO_CODE") & "'").Length > 0 Then
                    rowDETPROM1 = dst.Tables("DETPROM1").Select("PROMO_CODE = '" & rowDETJOBM1("PROMO_CODE") & "'")(0)
                End If
            End If

            If dst.Tables("ARTCUST1").Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                rowARTCUST1 = dst.Tables("ARTCUST1").Select("CUST_CODE = '" & CUST_CODE & "'")(0)
            End If

            If rowARTCUST1 Is Nothing Then
                sErrorList.Add("B", ErrorMasterCodeList("B"))
            ElseIf rowARTCUST1.Item("CUST_SHIP_TO_NO_REQD") & String.Empty = "1" AndAlso CUST_SHIP_TO_NO.Length = 0 Then
                sErrorList.Add("C", ErrorMasterCodeList("C"))
            End If

            If rowARTCUST1 IsNot Nothing AndAlso CUST_SHIP_TO_NO.Length > 0 Then
                If dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE & "' AND CUST_SHIP_TO_NO = '" & CUST_SHIP_TO_NO & "'").Length > 0 Then
                    rowARTCUST2 = dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE & "' AND CUST_SHIP_TO_NO = '" & CUST_SHIP_TO_NO & "'")(0)
                End If
                If rowARTCUST2 Is Nothing Then
                    sErrorList.Add("D", ErrorMasterCodeList("D"))
                End If
            End If

            rowSOTORDR5 = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'")(0)
            Dim CUST_NAME = rowSOTORDR5.Item("CUST_NAME") & String.Empty
            Dim CUST_ADDR1 = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
            Dim CUST_CITY = rowSOTORDR5.Item("CUST_CITY") & String.Empty
            Dim CUST_STATE = rowSOTORDR5.Item("CUST_STATE") & String.Empty
            Dim CUST_ZIP_CODE = rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty

            If CUST_NAME.Length = 0 _
                OrElse CUST_ADDR1.Length = 0 _
                OrElse CUST_CITY.Length = 0 _
                OrElse CUST_STATE.Length = 0 _
                OrElse CUST_ZIP_CODE.Length = 0 Then
                sErrorList.Add("E", ErrorMasterCodeList("E"))
            End If

            For Each rowDETJOBM3 As DataRow In dst.Tables("DETJOBM3").Select("JOB_NO = '" & jobNumber & "'")
                ValidateDetjobm3Row(rowDETJOBM3.Item("JOB_NO") & String.Empty, rowDETJOBM3.Item("RL") & String.Empty)
            Next

            Dim LENS_ORDER As String = rowDETJOBM1.Item("LENS_ORDER") & String.Empty
            BlankSelectionException.Clear()
            prevent_blank_selection = True
            'If Not BlankSelections(JOB_NO, True) Then

            If rowDETJOBM1.Item("FRAME_TYPE_CODE") & String.Empty = "NONE" Then
                If rowDETJOBM1.Item("FINISHED") & String.Empty = "O" _
                    AndAlso rowDETJOBM1.Item("EDGING") & String.Empty = "1" _
                    AndAlso rowDETJOBM1.Item("POLISHING") & String.Empty = "0" Then
                Else
                    sErrorList.Add("F", ErrorMasterCodeList("F"))
                End If
            End If

            If rowDETJOBM1.Item("LENS_DESIGNER_CODE") & String.Empty <> "SEIKO" _
                AndAlso rowDETJOBM1.Item("FINISHED") & String.Empty <> "O" Then

                If Not dst.Tables.Contains("DETDSGN4") OrElse _
                    dst.Tables("DETDSGN4").Select("LENS_DESIGNER_CODE = '" & rowDETJOBM1.Item("LENS_DESIGNER_CODE") & "' AND MATL_CODE = '" & rowDETJOBM1.Item("MATL_CODE") & "' AND COLOR_CODE = '" & rowDETJOBM1.Item("COLOR_CODE") & "'").Length = 0 Then
                    sErrorList.Add("G", ErrorMasterCodeList("G"))
                Else
                    rowDETDSGN4 = dst.Tables("DETDSGN4").Select("LENS_DESIGNER_CODE = '" & rowDETJOBM1.Item("LENS_DESIGNER_CODE") & "' AND MATL_CODE = '" & rowDETJOBM1.Item("MATL_CODE") & "' AND COLOR_CODE = '" & rowDETJOBM1.Item("COLOR_CODE") & "'")(0)
                End If

                If Not dst.Tables.Contains("DETDSGN6") OrElse _
                        dst.Tables("DETDSGN6").Select("LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "' AND MATL_CODE = '" & rowDETJOBM1.Item("MATL_CODE") & "'").Length = 0 Then
                    sErrorList.Add("H", ErrorMasterCodeList("H"))
                Else
                    rowDETDSGN6 = dst.Tables("DETDSGN6").Select("LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "' AND MATL_CODE = '" & rowDETJOBM1.Item("MATL_CODE") & "'")(0)
                End If

                If dst.Tables.Contains("DETDSGN3") AndAlso _
                        dst.Tables("DETDSGN3").Select("LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "' AND MATL_CODE = '" & rowDETJOBM1.Item("MATL_CODE") & "' AND COLOR_CODE = '" & rowDETJOBM1.Item("COLOR_CODE") & "'").Length > 0 Then
                    rowDETDSGN3 = dst.Tables("DETDSGN3").Select("LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "' AND MATL_CODE = '" & rowDETJOBM1.Item("MATL_CODE") & "' AND COLOR_CODE = '" & rowDETJOBM1.Item("COLOR_CODE") & "'")(0)
                End If
                If rowDETDSGN3 Is Nothing Then
                    sErrorList.Add("I", ErrorMasterCodeList("I"))
                ElseIf rowDETDSGN3.Item("ACTIVE") & String.Empty <> "1" Then
                    sErrorList.Add("J", ErrorMasterCodeList("J"))
                End If

                If LENS_ORDER = "R" Or LENS_ORDER = "B" Then
                    If RL_ISSUES("R") <> String.Empty Then
                        sErrorList.Add("K", ErrorMasterCodeList("K") & ":" & Mid(RL_ISSUES("R"), 2))
                    End If
                End If

                If LENS_ORDER = "L" Or LENS_ORDER = "B" Then
                    If RL_ISSUES("L") <> String.Empty Then
                        sErrorList.Add("L", ErrorMasterCodeList("L") & ":" & Mid(RL_ISSUES("L"), 2))
                    End If
                End If

                Dim BLANK_SELECTION_VIA_LDS As String = String.Empty
                If dst.Tables.Contains("DETDSGN0") OrElse _
                        dst.Tables("DETDSGN0").Select("LENS_DESIGNER_CODE = '" & rowDETJOBM1.Item("LENS_DESIGNER_CODE") & "'").Length > 0 Then
                    rowDETDSGN0 = dst.Tables("DETDSGN0").Select("LENS_DESIGNER_CODE = '" & rowDETJOBM1.Item("LENS_DESIGNER_CODE") & "'")(0)
                End If

                If rowDETDSGN0 IsNot Nothing Then
                    BLANK_SELECTION_VIA_LDS = rowDETDSGN0.Item("BLANK_SELECTION_VIA_LDS") & String.Empty
                End If

                If dst.Tables.Contains("DETDSGN1") OrElse _
                        dst.Tables("DETDSGN1").Select("LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "'").Length > 0 Then
                    rowDETDSGN1 = dst.Tables("DETDSGN1").Select("LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "'")(0)
                End If

                If rowDETDSGN1 IsNot Nothing Then
                    If rowDETDSGN1.Item("CUST_SUPPLIED_BLANKS") & String.Empty = "1" Then
                        If rowDETJOBM1.Item("FINISHED") & String.Empty <> "O" Then
                            sErrorList.Add("N", ErrorMasterCodeList("N") & "(" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & ")")
                        End If
                    End If
                    If rowDETDSGN1.Item("LENS_DESIGN_STATUS") & String.Empty <> "A" Then
                        sErrorList.Add("O", ErrorMasterCodeList("O"))
                    End If
                Else
                    sErrorList.Add("M", ErrorMasterCodeList("M"))
                End If

                If LENS_ORDER = "R" Or LENS_ORDER = "B" Then
                    rowDETJOBM3 = dst.Tables("DETJOBM3").Rows.Find(New Object() {jobNumber, "R"})
                    If rowDETJOBM3.Item("OPC_CODE") & String.Empty = String.Empty And BLANK_SELECTION_VIA_LDS <> "1" Then
                        sErrorList.Add("P", String.Format("Invalid OPC for {0} eye - (No Record in Blank Selection Chart)", dst.Tables("DETJOBM3").Rows(0).Item("RL")))
                    End If
                End If

                If LENS_ORDER = "L" Or LENS_ORDER = "B" Then
                    rowDETJOBM3 = dst.Tables("DETJOBM3").Rows.Find(New Object() {jobNumber, "L"})
                    If rowDETJOBM3.Item("OPC_CODE") & String.Empty = String.Empty And BLANK_SELECTION_VIA_LDS <> "1" Then
                        sErrorList.Add("P", String.Format("Invalid OPC for {0} eye - (No Record in Blank Selection Chart)", dst.Tables("DETJOBM3").Rows(1).Item("RL")))
                    End If
                End If

                Dim CORRIDOR_LENGTH As String = rowDETJOBM1.Item("CORRIDOR_LENGTH") & String.Empty
                Dim MATL_DESC As String = rowDETJOBM1.Item("MATL_DESC") & String.Empty
                If CORRIDOR_LENGTH.Length = 0 Then
                    sErrorList.Add("Q", ErrorMasterCodeList("Q"))
                Else
                    If Not dst.Tables.Contains("DETDSGN2") _
                            OrElse dst.Tables("DETDSGN2").Select("LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "' AND CORRIDOR_LENGTH = '" & rowDETJOBM1.Item("CORRIDOR_LENGTH") & "'").Length = 0 Then
                        sErrorList.Add("R", ErrorMasterCodeList("R"))
                    Else
                        rowDETDSGN2 = dst.Tables("DETDSGN2").Select("LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "' AND CORRIDOR_LENGTH = '" & rowDETJOBM1.Item("CORRIDOR_LENGTH") & "'")(0)
                        Dim MIN_FITTING_HEIGHT As Double = Val(rowDETDSGN2.Item("MIN_FITTING_HEIGHT") & String.Empty)
                        If dst.Tables("DETJOBM3").Rows.Count > 0 Then
                            If LENS_ORDER <> "B" Then
                                sql = "RL = '" & LENS_ORDER & "'"
                            End If
                            Dim FITTING_HEIGHT As Double = Val(dst.Tables("DETJOBM3").Compute("MIN(FITTING_HEIGHT)", SQL) & String.Empty)
                            If FITTING_HEIGHT < MIN_FITTING_HEIGHT Then
                                zMsg = "Minimum Fitting Height for Corridor Selected is " & CStr(MIN_FITTING_HEIGHT)
                                If rowDETJOBM1.Item("JOB_TYPE_CODE") & String.Empty = "R" Then
                                    If sErrorList.Count = 0 AndAlso displayWarningMessages Then
                                        If MsgBox(zMsg & Environment.NewLine & "Continue anyway?", MsgBoxStyle.YesNo, "Fitting Height") = MsgBoxResult.No Then
                                            sErrorList.Add("WA", zMsg)
                                            Return False
                                        End If
                                    Else
                                        sErrorList.Add("WA", zMsg)
                                    End If
                                Else
                                    sErrorList.Add("S", zMsg)
                                End If
                            End If
                        End If
                    End If
                End If

                If rowDETJOBM1.Item("FRAME_TYPE_CODE") & String.Empty = String.Empty Then
                    sErrorList.Add("T", ErrorMasterCodeList("T"))
                End If

                If rowDETJOBM1.Item("TINT_CODE") & String.Empty = String.Empty Then
                    sErrorList.Add("U", ErrorMasterCodeList("U"))
                Else
                    If rowDETJOBM1.Item("TINT_CODE") & String.Empty <> "NONE" Then
                        If rowDETJOBM1.Item("TINT_COLOR") & String.Empty = String.Empty OrElse Val(rowDETJOBM1.Item("TINT_PCT") & String.Empty) = 0 Then
                            sErrorList.Add("V", ErrorMasterCodeList("V"))
                        End If
                    End If
                End If

                If rowDETJOBM1.Item("PATTERN_NO") & String.Empty <> String.Empty AndAlso rowDETJOBM1.Item("TRACE_FROM") & String.Empty <> "D" Then
                    rowDETJOBM1.Item("TRACE_FROM") = "D"
                End If

                If rowDETJOBM1.Item("SHAPE_TO_BE_MODIFIED") & String.Empty = "1" OrElse (rowDETJOBM1.Item("PATTERN_NO") & String.Empty).ToString.StartsWith("M") Then
                    If dst.Tables("DETDRIL1").Select("PATTERN_NO = '" & rowDETJOBM1.Item("PATTERN_NO") & "'").Length = 0 Then
                        sErrorList.Add("W", ErrorMasterCodeList("W"))
                    Else
                        rowDETDRIL1 = dst.Tables("DETDRIL1").Select("PATTERN_NO = '" & rowDETJOBM1.Item("PATTERN_NO") & "'")(0)
                    End If
                End If

                If rowDETJOBM1.Item("CUSTOM_FRAME_NO") & String.Empty <> String.Empty Then
                    If dst.Tables("DETCFRM1").Select("CUSTOM_FRAME_NO = '" & rowDETJOBM1.Item("CUSTOM_FRAME_NO") & "'").Length = 0 Then
                        sErrorList.Add("X", ErrorMasterCodeList("X"))
                    Else
                        rowDETCFRM1 = dst.Tables("DETCFRM1").Select("CUSTOM_FRAME_NO = '" & rowDETJOBM1.Item("CUSTOM_FRAME_NO") & "'")(0)
                    End If
                End If

                If rowDETJOBM1.Item("PATTERN_NO") & String.Empty <> String.Empty Then
                    If rowDETJOBM1.Item("CUSTOM_FRAME_NO") & String.Empty <> String.Empty OrElse _
                        rowDETJOBM1.Item("CUSTOM_FRAME_NEW") & String.Empty = "1" Then
                        sErrorList.Add("Y", ErrorMasterCodeList("Y"))
                    End If
                End If


                Dim TRACE_FROM As String = rowDETJOBM1.Item("TRACE_FROM") & String.Empty

                If TRACE_FROM = String.Empty Then
                    sErrorList.Add("Z", ErrorMasterCodeList("Z"))
                End If

                Dim FPC As String = rowDETJOBM1.Item("FPC") & String.Empty
                If rowDETJOBM1.Item("FINISHED") & String.Empty = "F" _
                    OrElse (rowDETJOBM1.Item("FINISHED") & String.Empty = "O" AndAlso (rowDETJOBM1.Item("EDGING") & String.Empty = "1" OrElse rowDETJOBM1.Item("POLISHING") & String.Empty <> "0" OrElse rowDETJOBM1.Item("WRAP_EDGE") & String.Empty = "1")) Then
                    If TRACE_FROM = "N" Then
                        sErrorList.Add("AA", ErrorMasterCodeList("AA"))
                    ElseIf TRACE_FROM = "D" Then
                        If rowDETJOBM1.Item("PATTERN_NO") & String.Empty <> String.Empty Then
                            ' PATTERN WILL BE USED FOR THE DATABASE TRACE
                        Else
                            If Not dst.Tables.Contains("DETTRCP2") Then
                                sErrorList.Add("AB", ErrorMasterCodeList("AB") & " " & FPC)
                            Else
                                If dst.Tables("DETTRCP2").Select("XREF_TYPE = 'FPC' and XREF_CODE = '" & FPC & "'").Length = 0 Then
                                    sErrorList.Add("AB", ErrorMasterCodeList("AB") & " " & FPC)
                                Else
                                    rowDETTRCP2 = dst.Tables("DETTRCP2").Select("XREF_TYPE = 'FPC' and XREF_CODE = '" & FPC & "'")(0)
                                End If
                            End If
                        End If
                    End If
                Else
                    If rowDETJOBM1.Item("CUSTOM_FRAME_NO") & String.Empty <> String.Empty _
                        OrElse rowDETJOBM1.Item("CUSTOM_FRAME_NEW") & String.Empty = "1" Then
                        sErrorList.Add("AC", ErrorMasterCodeList("AC"))
                    End If

                    If rowDETJOBM1.Item("PATTERN_NO") & String.Empty <> String.Empty Then
                        sErrorList.Add("AD", ErrorMasterCodeList("AD"))
                    End If
                End If

                If rowDETJOBM1.Item("FINISHED") & String.Empty = "O" Then
                    If rowDETJOBM1.Item("EDGING") & String.Empty <> "1" _
                    AndAlso Not rowDETJOBM1.Item("POLISHING") & String.Empty <> "0" _
                    AndAlso rowDETJOBM1.Item("AR_COATING") & String.Empty <> "1" _
                    AndAlso rowDETJOBM1.Item("WRAP_EDGE") & String.Empty <> "1" _
                    AndAlso rowDETJOBM1.Item("MIRROR_COATING") & String.Empty <> "1" Then
                        sErrorList.Add("AE", ErrorMasterCodeList("AE"))
                    End If
                End If

                If TRACE_FROM = "O" Then
                    If rowDETJOBT1 Is Nothing Then
                        sErrorList.Add("AF", ErrorMasterCodeList("AF"))
                    End If
                End If

                If TRACE_FROM = "D" Then
                    If rowDETJOBM1.Item("PATTERN_NO") & String.Empty <> String.Empty Then
                        ' PATTERN WILL BE USED FOR THE DATABASE TRACE
                    Else
                        If FPC = String.Empty Then
                            sErrorList.Add("AG", ErrorMasterCodeList("AG"))
                        End If
                    End If
                End If

                If Val(rowDETJOBM1.Item("FRAME_A_WIDTH") & String.Empty) = 0 _
                    OrElse Val(rowDETJOBM1.Item("FRAME_B_HEIGHT") & String.Empty) = 0 _
                    OrElse Val(rowDETJOBM1.Item("FRAME_DBL_BRIDGE") & String.Empty) < 0 _
                    OrElse Val(rowDETJOBM1.Item("FRAME_ED_DIAGONAL") & String.Empty) = 0 Then
                    If rowDETJOBM1.Item("FRAME_TYPE_CODE") & String.Empty <> "NONE" Then
                        sErrorList.Add("AH", ErrorMasterCodeList("AH"))
                    End If
                End If

                If Val(rowDETJOBM1.Item("FRAME_DBL_BRIDGE") & String.Empty) = 0 Then
                    If rowDETJOBM1.Item("FRAME_TYPE_CODE") & String.Empty <> "NONE" Then
                        If sErrorList.Count = 0 AndAlso displayWarningMessages Then
                            If MsgBox(ErrorMasterCodeList("WB") & Environment.NewLine & "Continue Anyway?", MsgBoxStyle.YesNo, "Frame Type") = MsgBoxResult.No Then
                                sErrorList.Add("WB", ErrorMasterCodeList("WB"))
                                Return False
                            End If
                        Else
                            sErrorList.Add("WB", ErrorMasterCodeList("WB"))
                        End If

                    End If
                End If

                If rowDETPARM1 Is Nothing Then
                    sErrorList.Add("AI", ErrorMasterCodeList("AI"))
                Else
                    If Val(rowDETJOBM1.Item("FRAME_A_WIDTH") & String.Empty) > Val(rowDETPARM1.Item("DE_PARM_MAX_A") & String.Empty) _
                        OrElse Val(rowDETJOBM1.Item("FRAME_B_HEIGHT") & String.Empty) > Val(rowDETPARM1.Item("DE_PARM_MAX_B") & String.Empty) _
                        OrElse Val(rowDETJOBM1.Item("FRAME_DBL_BRIDGE") & String.Empty) > Val(rowDETPARM1.Item("DE_PARM_MAX_DBL") & String.Empty) _
                        OrElse Val(rowDETJOBM1.Item("FRAME_ED_DIAGONAL") & String.Empty) > Val(rowDETPARM1.Item("DE_PARM_MAX_ED") & String.Empty) _
                        OrElse Val(rowDETJOBM1.Item("FRAME_A_WIDTH") & String.Empty) < 0 _
                        OrElse Val(rowDETJOBM1.Item("FRAME_B_HEIGHT") & String.Empty) < 0 _
                        OrElse Val(rowDETJOBM1.Item("FRAME_DBL_BRIDGE") & String.Empty) < 0 _
                        OrElse Val(rowDETJOBM1.Item("FRAME_ED_DIAGONAL") & String.Empty) < 0 Then
                        sErrorList.Add("AJ", ErrorMasterCodeList("AJ") & ": A=" & CStr(Val(rowDETPARM1.Item("DE_PARM_MAX_A") & String.Empty)) & ", B=" & CStr(Val(rowDETPARM1.Item("DE_PARM_MAX_B") & String.Empty)) & ", ED=" & CStr(Val(rowDETPARM1.Item("DE_PARM_MAX_ED") & String.Empty)) & ", DBL=" & CStr(Val(rowDETPARM1.Item("DE_PARM_MAX_DBL") & String.Empty)) & String.Empty)
                    End If
                End If

                If Val(rowDETJOBM1.Item("FRAME_ED_DIAGONAL") & String.Empty) < Val(rowDETJOBM1.Item("FRAME_A_WIDTH") & String.Empty) Then
                    sErrorList.Add("AJ", ErrorMasterCodeList("AJ") & ": ED <" & " A Value " & Val(rowDETJOBM1.Item("FRAME_A_WIDTH") & String.Empty))
                End If


                If Val(rowDETJOBM1.Item("FRAME_ED_DIAGONAL") & String.Empty) < Val(rowDETJOBM1.Item("FRAME_B_HEIGHT") & String.Empty) Then
                    sErrorList.Add("AJ", ErrorMasterCodeList("AJ") & ": ED <" & " B Value " & Val(rowDETJOBM1.Item("FRAME_B_HEIGHT") & String.Empty))
                End If

                If TRACE_FROM = "O" Then
                    If rowDETJOBM1.Item("JOB_TYPE_CODE") & String.Empty = "O" Then
                        sErrorList.Add("AK", ErrorMasterCodeList("AK"))
                    Else

                        If rowDETJOBT1 Is Nothing Then
                            sErrorList.Add("AL", ErrorMasterCodeList("AL") & "  (" & JOB_NO_ORIG & ")")
                        Else
                            If Not My.Computer.FileSystem.FileExists(rowDETPARM1.Item("DE_PARM_ARCHIVE_TRC") & "\" & JOB_NO_ORIG & ".TRC") Then
                                sErrorList.Add("AL", ErrorMasterCodeList("AL") & "  (" & JOB_NO_ORIG & ")")
                            Else
                                If rowDETJOBM1.Item("FRAME_STATUS") & String.Empty = "C" Then
                                    If sErrorList.Count = 0 AndAlso displayWarningMessages Then
                                        If MsgBox(ErrorMasterCodeList("WC") & "(" & JOB_NO_ORIG & ")" & Environment.NewLine & "Continue Anyway?", MsgBoxStyle.YesNo, "Trace") = MsgBoxResult.No Then
                                            sErrorList.Add("WC", ErrorMasterCodeList("WC") & "(" & JOB_NO_ORIG & ")")
                                            Return False
                                        End If
                                    Else
                                        sErrorList.Add("WC", ErrorMasterCodeList("WC") & "(" & JOB_NO_ORIG & ")")
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If rowDETFRAM1 Is Nothing Then
                    sErrorList.Add("AM", ErrorMasterCodeList("AM"))
                Else
                    If rowDETFRAM1.Item("FRAME_ETYP") & String.Empty = "2" Then
                        If rowDETMATL1 Is Nothing Then
                            sErrorList.Add("AN", ErrorMasterCodeList("AN"))
                        Else
                            If rowDETMATL1.Item("PERMITTED_WITH_RIMLESS") & String.Empty <> "1" Then
                                sErrorList.Add("AO", ErrorMasterCodeList("AO") & " (" & rowDETJOBM1.Item("MATL_CODE") & ")")
                            End If
                        End If
                    End If

                End If

                If rowDETJOBM1.Item("ORDR_CUST_PO") & String.Empty = String.Empty Then
                    sErrorList.Add("AP", ErrorMasterCodeList("AP"))
                Else
                    If Mid(rowDETJOBM1.Item("ORDR_CUST_PO") & String.Empty, 1, 3) <> "DEL" Then
                        sErrorList.Add("AQ", ErrorMasterCodeList("AQ"))
                    End If
                End If

                If rowDETJOBM1.Item("PATIENT_NAME") & String.Empty = String.Empty Then
                    sErrorList.Add("AR", ErrorMasterCodeList("AR"))
                End If

                If Not dst.Tables.Contains("SOTSVIA1") OrElse dst.Tables("SOTSVIA1").Select("SHIP_VIA_CODE = '" & rowDETJOBM1.Item("SHIP_VIA_CODE") & "'").Length = 0 Then
                    sErrorList.Add("AS", ErrorMasterCodeList("AS"))
                End If


                For Each rowDETJOBM3 As DataRow In dst.Tables("DETJOBM3").Rows
                    Dim RL As String = rowDETJOBM3.Item("RL")
                    If LENS_ORDER = RL Or LENS_ORDER = "B" Then
                        If Val(rowDETJOBM3.Item("CYLINDER") & String.Empty) <> 0 And Val(rowDETJOBM3.Item("AXIS") & String.Empty) = 0 _
                        Or Val(rowDETJOBM3.Item("CYLINDER") & String.Empty) = 0 And Val(rowDETJOBM3.Item("AXIS") & String.Empty) <> 0 Then
                            sErrorList.Add("AT", ErrorMasterCodeList("AT") & " " & RL)
                        End If

                        If (Val(rowDETJOBM3.Item("PRISM_IN") & String.Empty) <> 0 And rowDETJOBM3.Item("PRISM_IN_AXIS") & String.Empty = String.Empty) _
                        OrElse (Val(rowDETJOBM3.Item("PRISM_IN") & String.Empty) = 0 And rowDETJOBM3.Item("PRISM_IN_AXIS") & String.Empty <> String.Empty) _
                        OrElse (Val(rowDETJOBM3.Item("PRISM_UP") & String.Empty) <> 0 And rowDETJOBM3.Item("PRISM_UP_AXIS") & String.Empty = String.Empty) _
                        OrElse (Val(rowDETJOBM3.Item("PRISM_UP") & String.Empty) = 0 And rowDETJOBM3.Item("PRISM_UP_AXIS") & String.Empty <> String.Empty) Then
                            sErrorList.Add("AU", ErrorMasterCodeList("AU") & " - Rx for " & RL)
                        End If

                        If Val(rowDETJOBM3.Item("FITTING_HEIGHT") & String.Empty) = 0 Then
                            sErrorList.Add("AV", ErrorMasterCodeList("AV") & " " & RL)
                        End If

                        If rowDETDSGN1 IsNot Nothing Then
                            If rowDETDSGN1.Item("LENS_TYPE") & String.Empty = "P" Then
                                If Val(rowDETJOBM3.Item("ADD_POWER") & String.Empty) = 0 Then
                                    sErrorList.Add("AW", ErrorMasterCodeList("AW") & " - Rx for " & RL)
                                End If
                            End If
                        End If
                    End If
                Next

                If rowDETJOBM1.Item("JOB_HOLD_LAB") & String.Empty = "1" Then
                    If rowDETJOBM1.Item("JOB_HOLD_LAB_REASON") & String.Empty = String.Empty Then
                        sErrorList.Add("AX", ErrorMasterCodeList("AX"))
                    End If
                Else
                    If rowDETJOBM1.Item("JOB_HOLD_LAB_REASON") & String.Empty <> String.Empty Then
                        sErrorList.Add("AY", ErrorMasterCodeList("AY"))
                    End If
                End If

                'JOB_HOLD_INV
                If rowDETJOBM1.Item("JOB_HOLD_INV") & String.Empty = "1" Then
                    If rowDETJOBM1.Item("JOB_HOLD_INV_REASON") & String.Empty = String.Empty Then
                        sErrorList.Add("AZ", ErrorMasterCodeList("AZ"))
                    End If
                Else
                    If rowDETJOBM1.Item("JOB_HOLD_INV_REASON") & String.Empty <> String.Empty Then
                        sErrorList.Add("BA", ErrorMasterCodeList("BA"))
                    End If
                End If

                If sErrorList.Count = 0 Then
                    If LENS_ORDER = "B" Then
                        With dst.Tables("DETJOBM3")
                            Dim R_MONO_PD As Double = Val(.Select("RL='R'")(0).Item("MONO_PD") & String.Empty)
                            Dim L_MONO_PD As Double = Val(.Select("RL='L'")(0).Item("MONO_PD") & String.Empty)
                            If R_MONO_PD + L_MONO_PD < 45 OrElse R_MONO_PD + L_MONO_PD > 84 Then
                                If sErrorList.Count = 0 AndAlso displayWarningMessages Then
                                    If MsgBox(ErrorMasterCodeList("BB") & Environment.NewLine & "Continue Anyway?", MsgBoxStyle.YesNo, "Rx Values Verification") = MsgBoxResult.No Then
                                        sErrorList.Add("BB", ErrorMasterCodeList("BB"))
                                        Return False
                                    End If
                                Else
                                    sErrorList.Add("BB", ErrorMasterCodeList("BB"))
                                End If
                            End If
                        End With
                    End If
                End If
            End If

            If LENS_ORDER = "B" Then
                With dst.Tables("DETJOBM3")
                    If Math.Sign(Val(.Select("RL = 'R'")(0).Item("CYLINDER") & String.Empty)) <> Math.Sign(Val(.Select("RL = 'L'")(0).Item("CYLINDER") & String.Empty)) _
                        AndAlso (Val(.Select("RL = 'R'")(0).Item("CYLINDER") & String.Empty) <> 0 AndAlso Val(.Select("RL = 'L'")(0).Item("CYLINDER") & String.Empty) <> 0) Then
                        sErrorList.Add("BC", ErrorMasterCodeList("BC"))
                    End If
                End With

                ' If there are no errors, look at the signs on the cylinder
                If sErrorList.Count = 0 Then
                    With dst.Tables("DETJOBM3")
                        If Math.Sign(Val(.Select("RL = 'R'")(0).Item("SPHERE") & String.Empty)) <> Math.Sign(Val(.Select("RL = 'L'")(0).Item("SPHERE") & String.Empty)) _
                            AndAlso (Val(.Select("RL = 'R'")(0).Item("SPHERE") & String.Empty) <> 0 AndAlso Val(.Select("RL = 'L'")(0).Item("SPHERE") & String.Empty) <> 0) Then
                            zMsg = "Opposite signs on Lens Sphere." & Environment.NewLine & Environment.NewLine & "Are you sure you want to Continue?"

                            If displayWarningMessages Then
                                If MsgBox(ErrorMasterCodeList("BD") & Environment.NewLine & "Continue Anyway?", MsgBoxStyle.YesNo, "") = MsgBoxResult.No Then
                                    sErrorList.Add("BD", ErrorMasterCodeList("BD"))
                                    Return False
                                End If
                            Else
                                sErrorList.Add("BD", ErrorMasterCodeList("BD"))
                            End If
                        End If
                    End With
                End If
            End If

            If rowDETJOBM1.Item("PROMO_CODE") & String.Empty <> String.Empty Then
                If rowDETPROM1 Is Nothing Then
                    sErrorList.Add("BE", ErrorMasterCodeList("BE"))
                Else
                    If rowDETPROM1.Item("PROMO_ONE_TIME_ONLY") & String.Empty = "1" Then
                        If Not dst.Tables.Contains("DETJOBMP") Then
                            sErrorList.Add("BF", ErrorMasterCodeList("BF"))
                        Else
                            sql = "CUST_CODE = '" & CUST_CODE & "'"
                            sql &= " and JOB_NO <> '" & jobNumber & "'"
                            sql &= " and PROMO_CODE = '" & rowDETJOBM1.Item("PROMO_CODE") & "'"
                            sql &= " and JOB_STATUS IN ('H','O','F')"

                            If dst.Tables("DETJOBMP").Select(sql).Length > 0 Then
                                sErrorList.Add("BG", ErrorMasterCodeList("BG") & " on Job: " & dst.Tables("DETJOBMP").Select(sql)(0).Item("JOB_NO"))
                            End If

                        End If
                    End If
                End If
            End If

            For Each rowDETJOBM2 As DataRow In dst.Tables("DETJOBM2").Select("JOB_QTY <> 0")
                Dim JOB_PRICE As Decimal = Val(rowDETJOBM2.Item("JOB_PRICE") & String.Empty)
                If JOB_PRICE < 0 Then
                    sErrorList.Add("BG", ErrorMasterCodeList("BG") & " (" & Format(JOB_PRICE, "#,###.00") & ")")
                End If
            Next

            Return sErrorList.Count > 0

        Catch ex As Exception
            sErrorList.Add("SE", ErrorMasterCodeList("SE") & " " & ex.Message)
            Return False
        End Try
    End Function

    Private Function ValidateDetjobm3Row(ByVal JOB_NO As String, ByVal RL As String) As Boolean

        Try
            Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)
            Dim rowDETJOBM3 As DataRow = dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "' AND RL = '" & RL & "'")(0)

            If RL_ISSUES.ContainsKey(RL) Then
                RL_ISSUES.Remove(RL)
            End If
            RL_ISSUES.Add(RL, String.Empty)

            Dim LENS_DESIGN_CODE As String = rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty
            Dim MATL_CODE As String = rowDETJOBM1.Item("MATL_CODE") & String.Empty
            Dim COLOR_CODE As String = rowDETJOBM1.Item("COLOR_CODE") & String.Empty
            Dim COLOR_TYPE As String = rowDETJOBM1.Item("COLOR_TYPE") & String.Empty

            Dim CYLINDER As Double = Val(rowDETJOBM1.Item("CYLINDER") & String.Empty)
            Dim SPHERE As Double = Val(rowDETJOBM1.Item("SPHERE") & String.Empty)
            Dim ADD_POWER As Double = Val(rowDETJOBM1.Item("ADD_POWER") & String.Empty)

            If CYLINDER > 0 Then ' switch to minus cylinder format
                SPHERE = SPHERE + CYLINDER
                CYLINDER = -1 * CYLINDER
            End If

            If System.Math.Round(SPHERE * 4, 2) <> CInt(SPHERE * 4) Then
                SPHERE = CInt(SPHERE * 4) / 4
            End If
            If System.Math.Round(CYLINDER * 4, 2) <> CInt(CYLINDER * 4) Then
                CYLINDER = CInt(CYLINDER * 4) / 4
            End If

            Dim rowDETDSGN1 As DataRow = Nothing

            If dst.Tables.Contains("DETDSGN1") AndAlso _
                    dst.Tables("DETDSGN1").Select("LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "'").Length > 0 Then
                rowDETDSGN1 = dst.Tables("DETDSGN1").Select("LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "'")(0)
            End If

            If rowDETDSGN1 Is Nothing Then
                sErrorList.Add("M", ErrorMasterCodeList("M"))
                Return False
            End If

            ' User must validate the Seiko Designer Code
            'If rowDETJOBM1.Item("LENS_DESIGNER_CODE") & String.Empty = "SEIKO" Then
            '    Dim rowBlankInfo As DataRow = Blank_Selection_SEIKO _
            '         (rowDETDSGN1.Item("BLANK_TABLE_DESIGN_CODE") & String.Empty _
            '          , MATL_CODE, COLOR_TYPE, SPHERE, CYLINDER, ADD_POWER)

            '    If rowBlankInfo.Item("BASE_CURVE") Is DBNull.Value Then
            '        RL_ISSUES(RL) &= "," & "Blank Selection"
            '    End If
            'Else
            '    'Stop ' NEED TO VALIDATE BLANK SELECTION ON OTHER THAN SEIKO
            'End If

            Dim rowDETDSGN6 As DataRow = Nothing
            If dst.Tables.Contains("DETDSGN6") Then
                rowDETDSGN6 = dst.Tables("DETDSGN6").Rows.Find(New String() {LENS_DESIGN_CODE, MATL_CODE})
            End If

            If rowDETDSGN6 Is Nothing Then
                sErrorList.Add("H", ErrorMasterCodeList("H"))
                Return False
            End If

            Dim SPHERE_POS_MAX As Double = Val(rowDETDSGN6.Item("SPHERE_POS_MAX") & String.Empty)
            Dim SPHERE_NEG_MAX As Double = Val(rowDETDSGN6.Item("SPHERE_NEG_MAX") & String.Empty)
            Dim CYL_NEG_MAX As Double = Val(rowDETDSGN6.Item("CYL_NEG_MAX") & String.Empty)
            Dim ADD_MAX As Double = Val(rowDETDSGN6.Item("ADD_MAX") & String.Empty)
            Dim ADD_MIN As Double = Val(rowDETDSGN6.Item("ADD_MIN") & String.Empty)

            If SPHERE_POS_MAX <> 0 Then
                If SPHERE > SPHERE_POS_MAX Then
                    RL_ISSUES(RL) &= "," & "Sphere Max Pos = " & CStr(SPHERE_POS_MAX)
                End If
            End If

            If SPHERE_NEG_MAX <> 0 Then
                If SPHERE < SPHERE_NEG_MAX Then
                    RL_ISSUES(RL) &= "," & "Sphere Max Neg = " & CStr(SPHERE_NEG_MAX)
                End If
            End If

            If System.Math.Round(Val(rowDETJOBM3.Item("SPHERE") & String.Empty) * 4, 2) <> CInt(Val(rowDETJOBM3.Item("SPHERE") & String.Empty) * 4) Then
                RL_ISSUES(RL) &= "," & "Sphere (1/4)"
            End If

            If CYL_NEG_MAX <> 0 Then
                If CYLINDER < CYL_NEG_MAX Then
                    RL_ISSUES(RL) &= "," & "Cylinder Max Neg = " & CStr(CYL_NEG_MAX)
                End If
            End If

            If System.Math.Round(Val(rowDETJOBM3.Item("CYLINDER") & String.Empty) * 4, 2) <> CInt(Val(rowDETJOBM3.Item("CYLINDER") & String.Empty) * 4) Then
                RL_ISSUES(RL) &= "," & "Cylinder (1/4)"
            End If

            If Val(rowDETJOBM3.Item("MONO_PD") & String.Empty) > 40 Or Val(rowDETJOBM3.Item("MONO_PD") & String.Empty) < 22 Then
                RL_ISSUES(RL) &= "," & "Mono PD"
            ElseIf System.Math.Round(Val(rowDETJOBM3.Item("MONO_PD") & String.Empty) * 4, 2) <> CInt(Val(rowDETJOBM3.Item("MONO_PD") & String.Empty) * 4) Then
                RL_ISSUES(RL) &= "," & "Mono PD"
            End If

            If Val(rowDETJOBM3.Item("AXIS") & String.Empty) > 180 Or Val(rowDETJOBM3.Item("AXIS") & String.Empty) < 0 Then
                RL_ISSUES(RL) &= "," & "Axis"
            End If

            If rowDETJOBM1.Item("CORRIDOR_LENGTH") & String.Empty = String.Empty Then

                Dim FITTING_HEIGHT As Double
                FITTING_HEIGHT = Val(dst.Tables("DETJOBM3").Compute("MAX(FITTING_HEIGHT)", String.Empty) & String.Empty)

                Dim SQL = "Select CORRIDOR_LENGTH from DETDSGN2 " _
                & " where LENS_DESIGN_CODE = :PARM1" _
                & " and MIN_FITTING_HEIGHT = " _
                & " (Select Max (MIN_FITTING_HEIGHT) from DETDSGN2 " _
                & " where LENS_DESIGN_CODE = :PARM2" _
                & " and MIN_FITTING_HEIGHT <= :PARM3)"

                Dim CORRIDOR_LENGTH As String = ABSolution.ASCDATA1.GetDataValue(SQL, "VVN", New Object() {rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty, rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty, FITTING_HEIGHT})
                If CORRIDOR_LENGTH <> String.Empty Then
                    rowDETJOBM1.Item("CORRIDOR_LENGTH") = CORRIDOR_LENGTH
                End If
            End If

            If rowDETJOBM3.Item("RL") & String.Empty = "R" Then
                Dim rowDETJOBM3L = dst.Tables("DETJOBM3").Rows.Find(New String() {JOB_NO, "L"})
                If rowDETJOBM3.Item("ADD_POWER") & String.Empty = String.Empty Then
                    If dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_COPY_R_TO_L") & String.Empty = "1" Then
                        rowDETJOBM3L.Item("ADD_POWER") = rowDETJOBM3.Item("ADD_POWER") & String.Empty
                        rowDETJOBM3L.Item("MONO_PD") = rowDETJOBM3.Item("MONO_PD") & String.Empty
                        rowDETJOBM3L.Item("FITTING_HEIGHT") = rowDETJOBM3.Item("FITTING_HEIGHT") & String.Empty
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            sErrorList.Add("SE", ErrorMasterCodeList("SE") & " " & ex.Message)
            Return False
        End Try
    End Function

    Private Function SetJobHeaderDefaults(ByRef rowDETJOBM1 As DataRow, ByVal ORDR_SOURCE As String, ByVal ORDR_CALLER_NAME As String) As Boolean

        Try
            sErrorList.Clear()
            ' default values taken from DETJOBM1
            rowDETJOBM1.Item("JOB_STATUS") = "O"

            ORDR_SOURCE = ORDR_SOURCE & String.Empty
            If ORDR_SOURCE.Length > rowDETJOBM1.Table.Columns("ORDR_SOURCE").MaxLength Then
                ORDR_SOURCE = ORDR_SOURCE.Substring(0, rowDETJOBM1.Table.Columns("ORDR_SOURCE").MaxLength)
            End If
            If ORDR_SOURCE.Length = 0 Then ORDR_SOURCE = "K"
            rowDETJOBM1.Item("ORDR_SOURCE") = ORDR_SOURCE

            rowDETJOBM1.Item("JOB_TYPE_CODE") = "O" ' Original
            rowDETJOBM1.Item("ORDR_DATE") = DateTime.Now.ToString("MM/dd/yyyy")
            rowDETJOBM1.Item("ORDR_CUST_PO") = "DEL/" & "SVC" & "/" & rowDETJOBM1.Item("JOB_NO") & String.Empty
            rowDETJOBM1.Item("COMMENT_LAB") = String.Empty
            rowDETJOBM1.Item("JOB_STATUS") = "O"
            rowDETJOBM1.Item("INIT_DATE") = DateTime.Now
            rowDETJOBM1.Item("LAST_DATE") = rowDETJOBM1.Item("INIT_DATE")
            'rowDETJOBM1.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
            rowDETJOBM1.Item("LENS_ORDER") = "B"
            rowDETJOBM1.Item("FINISHED") = "U"
            rowDETJOBM1.Item("FRAME_STATUS") = "N"
            rowDETJOBM1.Item("TINT_CODE") = "NONE"
            rowDETJOBM1.Item("TRACE_FROM") = "N"
            rowDETJOBM1.Item("POLISHING") = "0"

            ORDR_CALLER_NAME = ORDR_CALLER_NAME & String.Empty
            If ORDR_CALLER_NAME.Length > rowDETJOBM1.Table.Columns("ORDR_CALLER_NAME").MaxLength Then
                ORDR_CALLER_NAME = ORDR_CALLER_NAME.Substring(0, rowDETJOBM1.Table.Columns("ORDR_CALLER_NAME").MaxLength)
            End If
            rowDETJOBM1.Item("ORDR_CALLER_NAME") = ORDR_CALLER_NAME

            rowDETJOBM1.Item("USE_THINNING_PRISM") = "1"
            rowDETJOBM1.Item("ORDR_DATE") = DateTime.Now

        Catch ex As Exception
            sErrorList.Add("SE", ErrorMasterCodeList("SE") & " " & ex.Message)
            Return False
        End Try
    End Function

#End Region

End Class
