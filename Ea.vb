Public Class Ea

    Public Const MaxTries As Integer = 10
    Public Const WaitMilliSeconds As Integer = 1000

    Public Const esShowErrMessage As Boolean = True
    Public Const esNoErrorMessage As Boolean = False
    Public Const esAskToFixIfErr As Boolean = True
    Public Const esNoAskToFixIfErr As Boolean = False

    Public Const sqlTrue As SByte = 1
    Public Const sqlFalse As SByte = 0

#Region " Create Database Errors "

    Public Const deCouldNotCreate As Integer = -60
    Public Const deDatabaseNotFound As Integer = -61
    Public Const deNoColumnName As Integer = -62
    Public Const deNoTextLength As Integer = -63
    Public Const deInvalidColumnType As Integer = -64
    Public Const deCreateTable As Integer = -65
    Public Const deCreatePrimaryKey As Integer = -66
    Public Const deCreateIndexes As Integer = -67
    Public Const deVerifyTableStruct As Integer = -68
    Public Const deNoDataAdapt As Integer = -69
    Public Const deVerifyTableData As Integer = -70
    Public Const deCreateQuery As Integer = -71
    Public Const deSortQuery As Integer = -72
    Public Const deVerifyQuery As Integer = -73
    Public Const deVerifyTableCount As Integer = -74
    Public Const deNoQueryName As Integer = -75
    Public Const deNoQueryText As Integer = -76
    Public Const deNoDaoDbEngine As Integer = -77
    Public Const deNoDaoWorkspace As Integer = -78
    Public Const deNoCheckbox As Integer = -79

#End Region

#Region " Sql Errors "

    Public Const sqlNoErrors As Integer = 0
    Public Const sqlErr As Integer = -500
    Public Const sqlNoConnection As Integer = -501
    Public Const sqlCouldNotOpen As Integer = -502
    Public Const sqlCouldNotClose As Integer = -503
    Public Const sqlInvalidConnStr As Integer = -504
    Public Const sqlInvalidCommand As Integer = -505
    Public Const sqlNoConvertToStr As Integer = -506
    Public Const sqlDeleteErr As Integer = -507
    Public Const sqlDropErr As Integer = -508
    Public Const sqlCreateErr As Integer = -509
    Public Const sqlAlterErr As Integer = -510
    Public Const sqlInvalidAlter As Integer = -511
    Public Const sqlNoDatabase As Integer = -512
    Public Const sqlNoTableName As Integer = -513
    Public Const sqlNoColumnName As Integer = -514
    Public Const sqlNoKeyName As Integer = -515
    Public Const sqlNoIndexName As Integer = -516
    Public Const sqlNoColumn As Integer = -517
    Public Const sqlNoUser As Integer = -518
    Public Const sqlInsertErr As Integer = -519
    Public Const sqlUpdateErr As Integer = -520
    Public Const sqlMakeTableErr As Integer = -521
    Public Const sqlRenameTableErr As Integer = -522
    Public Const sqlCreateIndexErr As Integer = -523
    Public Const sqlRenameColumnErr As Integer = -524
    Public Const sqlNoCommandText As Integer = -525
    Public Const sqlNoQueryName As Integer = -526
    Public Const sqlNoQueryType As Integer = -527
    Public Const sqlNoProcName As Integer = -528
    Public Const sqlNoViewName As Integer = -529
    Public Const sqlTableNotFound As Integer = -530
    Public Const sqlCreatePrimaryKey As Integer = -531
    Public Const sqlKeyWord As Integer = -532
    Public Const sqlEmptyTableErr As Integer = -533
    Public Const sqlNoSort As Integer = -534
    Public Const sqlInvalidSort As Integer = -535
    Public Const sqlAlreadyExists As Integer = -536
    Public Const sqlNoSettings As Integer = -537
    Public Const sqlConnectionStringErr As Integer = -538
    Public Const sqlInvalidWhere As Integer = -539
    Public Const sqlSelectAlreadyHasWhere As Integer = -540
    Public Const sqlQueryNotFound As Integer = -541
    Public Const sqlInvalidParameters As Integer = -542
    Public Const sqlOverMaxTries As Integer = -543
    Public Const sqlInvalidColumnType As Integer = -544

    Public Const sqlCannotConnectToServer As Integer = -550
    Public Const sqlInvalidUserNamePassword As Integer = -551

    Public Const sqlOtherErr As Integer = -599

#End Region

End Class
