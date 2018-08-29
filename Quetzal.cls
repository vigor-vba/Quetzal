VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Quetzal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#############################################################################
' Quetzal: A simple ADO handler.
' MIT License | github.com/vigor-vba/Quetzal
'#############################################################################
Option Explicit
#Const ISDEBUG = False
Private Const VERSION_MAJOR As Long = 1
Private Const VERSION_MINOR As Long = 1
Private Const VERSION_PATCH As Long = 1

'-----------------------------------------------------------------------------
' Enums
'-----------------------------------------------------------------------------
' ADODB: ObjectState
Private Enum ObjectState
    adStateClosed = 0
    adStateOpen = 1
    adStateConnecting = 2
    adStateExecuting = 4
    adStateFetching = 8
End Enum

' ADODB: CursorLocation
Private Enum CursorLocation
    adUseNone = 1
    adUseServer = 2
    adUseClient = 3
    adUseClientBatch = adUseClient
End Enum

' ADODB: CursorType
Private Enum CursorType
    adOpenUnspecified = -1
    adOpenForwardOnly = 0
    adOpenKeyset = 1
    adOpenDynamic = 2
    adOpenStatic = 3
End Enum

' ADODB: LockType
Private Enum LockType
    adLockUnspecified = -1
    adLockReadOnly = 1
    adLockPessimistic = 2
    adLockOptimistic = 3
    adLockBatchOptimistic = 4
End Enum

' ADODB: CommandType
Private Enum CommandType
    adCmdUnspecified = -1
    adCmdText = 1
    adCmdTable = 2
    adCmdStoredProc = 4
    adCmdUnknown = 8
    adCmdFile = 256
    adCmdTableDirect = 512
End Enum

' Quetzal: ConnectionType
Public Enum ConnectionType
    qzCustom = 0
    qzJetOLEDB = 1
    qzAceOLEDB = 2
    qzExcelDriver = 3
    qzTextDriver = 4
    qzOraOLEDB = 5
    qzMSDAORA = 6
    qzODBCOracle = 7
    dbMdb = qzJetOLEDB
    dbAccdb = qzAceOLEDB
    dbExcel = qzExcelDriver
    dbCSV = qzTextDriver
    dbOracle = qzOraOLEDB
End Enum

' Quetzal: StatementClass
Private Enum StatementClass
    Undefined = -1
    dmlSelect = 1
    dmlInsert = 2
    dmlUpdate = 3
    dmlDelete = 4
End Enum

'-----------------------------------------------------------------------------
' Types
'-----------------------------------------------------------------------------
Private Type Exception
    Title       As String
    Number      As Long
    Description As String
End Type

'-----------------------------------------------------------------------------
' Fields
'-----------------------------------------------------------------------------
Private con_       As Object          ' ADODB.Connection
Private cmd_       As Object          ' ADODB.Command
Private rec_       As Object          ' ADODB.Recordset
Private prmParser_ As Object          ' VBScript.RegExp
Private statement_ As String          ' Query statement
Private params_    As Object          ' Parameter container
Private prmMap_    As Collection      ' Parameter map
Private affected_  As Long            ' Affected rows
Private stmtClass_ As StatementClass  ' Statement class
Private trans_     As Long            ' Transaction level
Private exception_ As Exception       ' Process result


'*****************************************************************************
' Properties
'*****************************************************************************
' Version
Public Property Get Version() As String
    Version = VERSION_MAJOR & "." & VERSION_MINOR & "." & VERSION_PATCH
End Property

' QueryStatement
Public Property Let QueryStatement(Prop As String)
    If IsConnect Then
        statement_ = Trim(Prop)
        stmtClass_ = ClassifyStatement(statement_)
        Set cmd_ = CreateCommand(statement_)
    End If
End Property
Public Property Get QueryStatement() As String
    QueryStatement = statement_
End Property

' Params
Public Property Get Params() As Object
    Set Params = params_
End Property

' IsConnect
Public Property Get IsConnect() As Boolean
    If con_ Is Nothing Then
        IsConnect = False
    Else
        IsConnect = (con_.State = ObjectState.adStateOpen)
    End If
End Property

' HasRecordset
Public Property Get HasRecordset() As Boolean
    If rec_ Is Nothing Then
        HasRecordset = False
    Else
        HasRecordset = (rec_.State = ObjectState.adStateOpen)
    End If
End Property

' IsTrans
Public Property Get IsTrans() As Boolean
    IsTrans = (trans_ > 0)
End Property

' Recordset
Public Property Get Recordset() As Object
    Set Recordset = rec_
End Property

' Recordmap
Public Property Get Recordmap() As Variant
    Dim recmap() As String  '
    Dim recRows  As Long    '
    Dim recCols  As Long    '
    Dim r        As Long    '
    Dim c        As Long    '

    If rec_ Is Nothing Then
        Recordmap = Empty
        Exit Function
    End If

    If rec_.State = ObjectState.adStateClosed Then
        Recordmap = Empty
        Exit Function
    End If

    If rec_.RecordCount <= 0 Then
        Recordmap = Empty
        Exit Function
    End If

    recRows = rec_.RecordCount - 1
    recCols = rec_.Fields.Count - 1

    ReDim recmap(recRows, recCols) As String
    r = 0
    Do Until rec_.EOF
        For c = 0 To recCols
            recmap(r, c) = rec_.Fields(c) & ""
        Next c
        r = r + 1
        Call rec_.MoveNext
    Loop

    Recordmap = recmap

End Property

' AffectedRows
Public Property Get AffectedRows() As Long
    AffectedRows = affected_
End Property


'*****************************************************************************
' Class_Initialize()
'*****************************************************************************
Private Sub Class_Initialize()
    Set prmParser_ = CreateObject("VBScript.RegExp")
    prmParser_.Global = True
    prmParser_.IgnoreCase = False
    prmParser_.Pattern = ":(\w+)(?=(([^'#]*['#]){2})*[^'#]*$)"
End Sub


'*****************************************************************************
' Class_Terminate()
'*****************************************************************************
Private Sub Class_Terminate()
    Call Me.Break
End Sub


'*****************************************************************************
' [+] Connect()
'*****************************************************************************
Public Function Connect(Connection As ConnectionType, Source As String, Optional UserId As String = "", Optional Password As String = "", Optional CustomString As String = "") As Quetzal
    Dim conStr As String

    Select Case Connection
        Case ConnectionType.qzJetOLEDB
            conStr = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                   & "Data Source=" & Source & ";" _
                   & IIf(Len(UserId) > 0, "User ID=" & UserId & ";", "") _
                   & IIf(Len(Password) > 0, "Password=" & Password & ";", "")

        Case ConnectionType.qzAceOLEDB
            conStr = "Provider=Microsoft.Ace.OLEDB.12.0;" _
                   & "Data Source=" & Source & ";" _
                   & IIf(Len(Password) > 0, "Jet OLEDB:Database Password=" & Password & ";", "")

        Case ConnectionType.qzExcelDriver
            conStr = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" _
                   & "Dbq=" & Source & ";" _
                   & "ReadOnly=0"

        Case ConnectionType.qzTextDriver
            conStr = "Driver={Microsoft Text Driver (*.txt; *.csv)};" _
                   & "Dbq=" & Source & ";" _
                   & "ReadOnly=0"

        Case ConnectionType.qzOraOLEDB
            conStr = "Provider=OraOLEDB.Oracle;" _
                   & "Data Source=" & Source & ";" _
                   & "User ID=" & UserId & ";" _
                   & "Password=" & Password & ";"

        Case ConnectionType.qzMSDAORA
            conStr = "Provider=MSDAORA;" _
                   & "Data Source=" & Source & ";" _
                   & "User ID=" & UserId & ";" _
                   & "Password=" & Password & ";"

        Case ConnectionType.qzODBCOracle
            conStr = "Driver={Microsoft ODBC for Oracle};" _
                   & "Server=" & Source & ";" _
                   & "Uid=" & UserId & ";" _
                   & "Pwd=" & Password & ";"

        Case ConnectionType.qzCustom
            conStr = CustomString

    End Select

    On Error Resume Next
    Set con_ = CreateObject("ADODB.Connection")
    Call con_.Open(conStr)
    exception_.Title = "Connect Database Result"
    exception_.Number = Err.Number
    exception_.Description = Err.Description
    On Error GoTo 0

    Set Connect = Me

End Function


'*****************************************************************************
' [+] Query()
'*****************************************************************************
Public Function Query(Optional QueryStatement As String, Optional WithCommit As Boolean = False) As Quetzal
    If Not Me.IsConnect Then
        GoTo Suspend
    End If

    If Len(QueryStatement) > 0 Then
        Me.QueryStatement = QueryStatement
    End If

    Select Case stmtClass_
        Case StatementClass.dmlSelect
            Set rec_ = CreateRecordset()
        Case StatementClass.dmlInsert, StatementClass.dmlUpdate, StatementClass.dmlDelete
            affected_ = ExecuteCommand(WithCommit)
    End Select

Suspend:
    Set Query = Me

End Function


'*****************************************************************************
' [+] Break()
'*****************************************************************************
Public Sub Break(Optional WithCommit As Boolean = False)
    If Me.IsConnect Then
        If WithCommit And Me.IsTrans Then
            Call Me.Commit
        End If
        Call con_.Close
    End If
End Sub


'*****************************************************************************
' [+] Commit()
'*****************************************************************************
Public Function Commit() As Boolean
    If Me.IsConnect And Me.IsTrans Then
        Commit = con_.CommitTrans
        trans_ = 0
    End If
End Function


'*****************************************************************************
' [+] Rollback()
'*****************************************************************************
Public Function Rollback() As Boolean
    If Me.IsConnect And Me.IsTrans Then
        Rollback = con_.RollbackTrans
        trans_ = 0
    End If
End Function


'*****************************************************************************
' [-] CreateRecordset()
'*****************************************************************************
Private Function CreateRecordset() As Object
    Dim recset As Object

    On Error Resume Next
    Set recset = CreateObject("ADODB.Recordset")
    recset.CursorLocation = CursorLocation.adUseClientBatch
    recset.CursorType = CursorType.adOpenStatic
    recset.LockType = LockType.adLockBatchOptimistic

    Call recset.Open(cmd_)
    If recset.State = ObjectState.adStateOpen Then
        Set recset.ActiveConnection = Nothing
    End If

    exception_.Title = "Create Recordset Result"
    exception_.Number = Err.Number
    exception_.Description = Err.Description
    On Error GoTo 0

    Set CreateRecordset = recset

End Function


'*****************************************************************************
' [-] ExecuteCommand()
'*****************************************************************************
Private Function ExecuteCommand(WithCommit As Boolean) As Long
    Dim afct As Long

    If Not Me.IsTrans Then
        trans_ = con_.BeginTrans
    End If

    On Error Resume Next
    Call cmd_.Execute(afct, CreateParameters)
    exception_.Title = "Execute Command Result"
    exception_.Number = Err.Number
    exception_.Description = Err.Description
    On Error GoTo 0

    If WithCommit Then
        Call Me.Commit
    End If

    ExecuteCommand = afct

End Function


'*****************************************************************************
' [-] ClassifyStatement()
'*****************************************************************************
Private Function ClassifyStatement(Statement As String) As StatementClass
    Dim stmt As String

    stmt = UCase(Statement)
    Select Case True
        Case stmt Like "SELECT *": ClassifyStatement = StatementClass.dmlSelect
        Case stmt Like "INSERT *": ClassifyStatement = StatementClass.dmlInsert
        Case stmt Like "UPDATE *": ClassifyStatement = StatementClass.dmlUpdate
        Case stmt Like "DELETE *": ClassifyStatement = StatementClass.dmlDelete
        Case Else: ClassifyStatement = StatementClass.Undefined
    End Select
End Function


'*****************************************************************************
' [-] CreateCommand()
'*****************************************************************************
Private Function CreateCommand(Statement As String) As Object
    Dim cmd As Object

    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandType = CommandType.adCmdText
    cmd.ActiveConnection = con_
    cmd.CommandText = CreateCommandText(Statement)
    cmd.Prepared = params_.Count > 0

    Set CreateCommand = cmd

End Function


'*****************************************************************************
' [-] CreateCommandText()
'*****************************************************************************
Private Function CreateCommandText(Statement As String) As String
    Dim matches As Object  ' RegExp.MatchingCollection
    Dim mch     As Object  ' RegExp.Matche
    Dim pName   As String  ' Parameter name

    Set params_ = CreateObject("Scripting.Dictionary")
    Set prmMap_ = New Collection

    Set matches = prmParser_.Execute(Statement)
    For Each mch In matches
        pName = mch.SubMatches(0)
        If pName = "" Then
            pName = Format(params_.Count + 1, "P000")
        End If

        params_(pName) = ""
        Call prmMap_.Add(pName)
    Next

    CreateCommandText = prmParser_.Replace(Statement, "?")

End Function


'*****************************************************************************
' [-] CreateParameters()
'*****************************************************************************
Private Function CreateParameters() As Variant
    Dim prms    As Variant
    Dim cntPrms As Long
    Dim i       As Long

    cntPrms = prmMap_.Count
    If cntPrms <= 0 Then
        CreateParameters = Null
        Exit Function
    End If

    ReDim prms(cntPrms - 1) As Variant
    For i = 1 To cntPrms
        prms(i - 1) = params_(prmMap_(i))
    Next i

    CreateParameters = prms
End Function
