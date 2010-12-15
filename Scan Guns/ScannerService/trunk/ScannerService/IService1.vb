' NOTE: If you change the interface name "IScannerService" here, you must also update the reference to "IScannerService" in App.config.
<ServiceContract()> _
Public Interface IScannerService

    '<OperationContract()> _
    'Function GetData(ByVal value As Integer) As String

    '<OperationContract()> _
    'Function GetDataUsingDataContract(ByVal composite As CompositeType) As CompositeType

    ' TODO: Add your service operations here
    <OperationContract()> _
    Function GetItemInfo(ByVal itemUpcCode As String, ByVal whseCode As String) As DataSet

    <OperationContract()> _
    Function CheckBin(ByVal binNo As String) As String

    <OperationContract()> _
    Function GetBin(ByVal itemUpcCode As String) As String

    <OperationContract()> _
    Function IsBinValid(ByVal itemUpcCode As String, ByVal whse As String) As String

    <OperationContract()> _
    Function GetScanData(ByVal binNo As String) As DataSet

    <OperationContract()> _
    Function UpdateItemInfo(ByVal dst As DataSet, ByVal binNo As String, ByVal priceCatgy As String, ByVal operId As String, ByVal whseCode As String, ByVal forUpdate As String) As String

    <OperationContract()> _
    Function LoadPO(ByVal invPackUPC As String) As DataSet

    <OperationContract()> _
    Function UpdatePO(ByVal WH_OPER_ID As String, ByVal dst As DataSet) As String


End Interface

' Use a data contract as illustrated in the sample below to add composite types to service operations
'<DataContract()> _
'Public Class scannedInfo
'    Private _itemCode As String
'    Private _qty As Long
'    Private _scans As New Dictionary(Of String, Long)



'    Sub New()
'        _itemCode = ""
'        _qty = 0
'        _scans = New Dictionary(Of String, Long)
'    End Sub

'    <DataMember()> _
'    Property itemCode() As String
'        Get
'            Return _itemCode
'        End Get
'        Set(ByVal value As String)
'            _itemCode = value
'        End Set
'    End Property

'    <DataMember()> _
'    Property Qty() As Long
'        Get
'            Return _qty
'        End Get
'        Set(ByVal value As Long)
'            _qty = value
'        End Set
'    End Property


'    <DataMember()> _
'    Property AddItem() As Boolean
'        Get
'            Return False
'        End Get
'        Set(ByVal value As Boolean)
'            If value = True Then
'                If _scans Is Nothing Then
'                    _scans = New Dictionary(Of String, Long)
'                End If
'                If _scans.ContainsKey(itemCode) Then
'                    _scans(itemCode) += Qty
'                Else
'                    _scans.Add(itemCode, Qty)
'                End If
'            End If
'        End Set
'    End Property

'    <DataMember()> _
'        Property Scans() As Dictionary(Of String, Long)
'        Get
'            Return _scans
'        End Get
'        Set(ByVal value As Dictionary(Of String, Long))
'            _scans = value
'        End Set
'    End Property

'End Class

'<DataContract()> _
'Public Class CompositeType

'    Private boolValueField As Boolean
'    Private stringValueField As String

'    <DataMember()> _
'    Public Property BoolValue() As Boolean
'        Get
'            Return Me.boolValueField
'        End Get
'        Set(ByVal value As Boolean)
'            Me.boolValueField = value
'        End Set
'    End Property

'    <DataMember()> _
'    Public Property StringValue() As String
'        Get
'            Return Me.stringValueField
'        End Get
'        Set(ByVal value As String)
'            Me.stringValueField = value
'        End Set
'    End Property

'End Class

