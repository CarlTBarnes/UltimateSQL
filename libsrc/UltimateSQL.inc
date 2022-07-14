!region Begin Comments

!!>Version         2022.2.8.1
!!>Template        UltimateSQL.tpl
!!>Revisions       Created Help file
!!>                Added comments to methods
!!>Created         DEC  2,2020,  8:35:39am
!!>Modified        JAN  9,2022,  9:14:45am
!endregion End Comments

                                PRAGMA('link(ODBC32.LIB)')   

    INCLUDE('UltimateSQLDirect.INC'),ONCE  

! RRS = Rick Smith
! RRS 01/22/19 Added to support large tables.
!
    INCLUDE('ctFieldQ.INC'),ONCE
!
! RRS 01/22/19 End.
!
Sybase                          EQUATE(0)
MS_SQL                          EQUATE(1)
Postgres                        EQUATE(2)
My_SQL                          EQUATE(3) 
ODBC_Direct                     EQUATE(4)  
SQLite                          EQUATE(5)  

ULSDriver_Sybase                EQUATE(0)
ULSDriver_MSSQL                 EQUATE(1)
ULSDriver_Postgres              EQUATE(2)
ULSDriver_MySQL                 EQUATE(3) 
ULSDriver_ODBC                  EQUATE(4)  
ULSDriver_SQLite                EQUATE(5)

BusyHandlingDoNothing           EQUATE(1)
BusyHandlingOneConnectionPerThread      EQUATE(2)
BusyHandlingRetryOnBusy         EQUATE(3)
BusyHandlingConnectionLocking   EQUATE(4) 

SQL_HANDLE_ENV                  EQUATE(1)
SQL_HANDLE_DBC                  EQUATE(2)
SQL_HANDLE_STMT                 EQUATE(3)
SQL_HANDLE_DESC                 EQUATE(4)

SQL_SUCCESS                     EQUATE(0)
SQL_SUCCESS_WITH_INFO           EQUATE(1)
SQL_ERROR                       EQUATE(-1)
SQL_INVALID_HANDLE              EQUATE(-2)

SQL_NULL_HANDLE                 EQUATE(0)

SQL_ATTR_ODBC_VERSION           EQUATE(200)

SQL_OV_ODBC2                    EQUATE(2)
SQL_OV_ODBC3                    EQUATE(3)

SQL_COPT_SS_INTEGRATED_SECURITY         EQUATE(1203)    

us:ForJsonAuto                  EQUATE(1)
us:ForJsonPath                  EQUATE(2)


LoginConnections                Queue,Type
Name                                STRING(512)
                                End      
!!!
QueryMethodDummyFile            EQUATE(1)
QueryMethodODBC                 EQUATE(2)

qExtendedPropertiesType         QUEUE,TYPE 
Object                              STRING(200)
Path                                STRING(200)
Name                                STRING(100)
Value                               STRING(10000)
                                END      

qSQLServerListType              QUEUE,TYPE
Name                                STRING(512)                           
                                END 

                                  
qErrorsType                     QUEUE,TYPE  
Thread                              LONG
ErrorCode                           STRING(200)
ErrorDescription                    STRING(500)
                                END
                     
qStatementsType                 QUEUE,TYPE   
ID                                  LONG(0)
Statement                           STRING(100000)
                                END
          
qBatchErrorsType                QUEUE,TYPE   
ID                                  LONG(0)
Statement                           STRING(100000) 
Error                               STRING(1000)
                                END

!--------------------------------------------------------------------------------
UltimateSQL                     CLASS(UltimateSQLDirect),TYPE,MODULE('UltimateSQL.CLW'),LINK('UltimateSQL.CLW')     
!
!Properties
!----------  

qExtendedProperties                 &qExtendedPropertiesType
qSQLServerList                      &qSQLServerListType     
qErrors                             &qErrorsType    
qStatements                         &qStatementsType 
qBatchErrors                        &qBatchErrorsType

LoginConnections                    &LoginConnections,PROTECTED  
USQLCriticalSection                 &ICriticalSection

QueryTableName                      STRING(200),PRIVATE  
InitHasRun                          BYTE

Server                              STRING(200)
Database                            STRING(200)
User                                STRING(200)
Password                            STRING(200)
Catalog                             STRING(200)  
ConnectionString                    STRING(200)  
FullConnectionString                STRING(200)   
ConnectionStringWithProvider        STRING(200)     
ApplicationName                     STRING(200)
ODBCDriver                          STRING(200)   
WSID                                STRING(200)
USQLDriverOptions                   STRING(1000)     

ShowODBCResults                     BYTE(0)

SQLError                            STRING(1800) 
SendToDebug                         BYTE(0)

SQLiteInMemory                      BYTE(0)

AllDebugOff                         BYTE(0)

AutoFill                            BYTE(0)
AutoFillINIFile                     STRING(260)
AutoFillSection                     STRING(60)

_Driver                             LONG,PROTECTED            !1 = MS SQL 
 
QueryResultsShowInPopUp             BYTE(FALSE) ! Set to TRUE to view your Query Results in a Pop-Up, FALSE to turn it off
ShowQueryInDebugView                BYTE(FALSE) ! Set to TRUE to send your Query to DebugView, FALSE to turn it off
AddQueryToClipboard                 BYTE(FALSE) ! Set to TRUE to send your Query to the Clipboard, FALSE to turn it off
AppendQueryToClipboard              BYTE(FALSE) ! Set to TRUE to send your Query to the Clipboard, FALSE to turn it off 

HINT                                STRING(220)
LOGONSCREEN                         BYTE(TRUE)
SAVESTOREDPROC                      BYTE(TRUE)
TRUSTEDCONNECTION                   BYTE(FALSE)  
MULTIPLEACTIVERESULTSETS            BYTE(TRUE)
ALLOWDETAILS                        BYTE(TRUE)
APPENDBUFFER                        LONG
AUTOINC                             STRING(100)
BINDCOLORDER                        BYTE(0)
BINDCONSTANTS                       BYTE(TRUE)
BUSYHANDLING                        BYTE(3)
BUSYMESSAGE                         STRING(200)
BUSYRETRIES                         BYTE(20)
CLIPSTRINGS                         BYTE(TRUE)
FASTCOLUMNFETCH                     BYTE(TRUE)
FORCEUPPERCASE                      BYTE(FALSE)
GATHERATOPEN                        BYTE(FALSE)
IGNORETRUNCATION                    BYTE(FALSE)
ISOLATIONLEVEL                      BYTE(1)
LOGFILEPATH                         STRING(220)
LOGFILEMESSAGE                      STRING(220)
NESTING                             BYTE(TRUE)
ODBCCALL                            BYTE(TRUE)
ORDERINSELECT                       BYTE(FALSE)
PREAUTOINC                          BYTE(TRUE)
TURBOSQL                            BYTE(FALSE)
USEINNERJOIN                        BYTE(TRUE)
VERIFYVIASELECT                     BYTE(FALSE)
WHERE                               STRING(220)
ZEROISNULL                          BYTE(TRUE)      

QueryMethod                         BYTE(1) 
AlwaysODBC                          BYTE(0)

RequireNativeClient                 BYTE(0) 
NativeClient                        STRING(1000)
NoNativeClientMessage               STRING(500)
NoODBCDriverMessage                 STRING(500)   
ODBCSilentInstall                   BYTE(0)
ODBCDriverMSIFileLocation           STRING(255)

CheckForAndRemoveClarionPrefixes    BYTE(0)
FieldSeparatorReplacement           STRING(6)
   
vBusyHandling                       STRING(50) 
vBusyRetries                        STRING(50) 
vGatherAtOpen                       STRING(50)   
vIgnoreTruncation                   STRING(50)   
vLogonScreen                        STRING(50)   
vMultipleActiveResultSets           STRING(50)    
vSaveStoredProcedure                STRING(50)    
vTurboSQL                           STRING(50)    
vVerifyViaSelect                    STRING(50)   

CustomConnectionString              STRING(1000)
                  
Provider                            STRING(20)
StatementID                         LONG(0)
StatementGUID                       STRING(36)
!
!Methods
!-------
      
AddColumn                           PROCEDURE(STRING pTable,STRING pColumn,STRING pType,<STRING pLength>,<STRING pOptions>),LONG,PROC,VIRTUAL
AlterColumn                         PROCEDURE(STRING pTable,STRING pColumn,STRING pType,<STRING pLength>,<STRING pOptions>),LONG,PROC,VIRTUAL  
Bracket                             PROCEDURE(STRING pValue),STRING,VIRTUAL
ColumnExists                        PROCEDURE(STRING pTable,STRING pColumn,<STRING pCatalog>,<STRING pSchema>),LONG,VIRTUAL
Connect                             PROCEDURE(*STRING pServer,*STRING pUser,*STRING pPassword,*STRING pDatabase,*BYTE pTrusted,<BYTE pLoginNamePasswordOnly>,BYTE pForce=0,<STRING pInstructions>),STRING,PROC,VIRTUAL  
CreateDatabase                      PROCEDURE(String Server, String USR, String PWD, String Database, <Byte Trusted>),BYTE,PROC,VIRTUAL  
CreateColumn                        PROCEDURE(STRING pTable,STRING pColumn,STRING pType,<STRING pLength>,<STRING pOptions>),LONG,PROC,VIRTUAL
DatabaseExists                      PROCEDURE(STRING pDatabase),BYTE,VIRTUAL                   
DisableTableConstraints             PROCEDURE(STRING pTable),PROC,VIRTUAL    
EnableTableConstraints              PROCEDURE(STRING pTable),PROC,VIRTUAL
DropColumn                          PROCEDURE(STRING pTable,STRING pColumn),LONG,PROC,VIRTUAL
DropDatabase                        PROCEDURE(STRING pDatabase),LONG,PROC,VIRTUAL
DropDependencies                    PROCEDURE(STRING pTable,STRING pColumn),LONG,PROC,VIRTUAL  
DropFunction                        PROCEDURE(STRING pFunction),LONG,PROC,VIRTUAL 
DropProcedure                       PROCEDURE(STRING pProcedure),LONG,PROC,VIRTUAL 
DropTable                           PROCEDURE(FILE pFile),LONG,PROC,VIRTUAL 
DropTable                           PROCEDURE(STRING pFile),LONG,PROC,VIRTUAL 
DropTrigger                         PROCEDURE(STRING pTrigger),LONG,PROC,VIRTUAL 
DropView                            PROCEDURE(STRING pView),LONG,PROC,VIRTUAL 
Empty                               PROCEDURE(FILE pFile),LONG,PROC,VIRTUAL 
ExecuteScriptFromFile               PROCEDURE(STRING pFileName),BYTE,PROC,VIRTUAL
ExecuteScriptFromBlob               PROCEDURE(*BLOB pBlob),BYTE,PROC,VIRTUAL 
ExecuteScriptsFromQUEUE             PROCEDURE(*QUEUE pQueue),VIRTUAL
Get                                 PROCEDURE(*FILE pFile, *KEY pKey, <STRING pSelect>), BYTE, PROC,VIRTUAL
Get                                 PROCEDURE(*KEY pKey, <STRING pSelect>), BYTE, PROC,VIRTUAL
GetColumnLength                     PROCEDURE(STRING pTable,STRING pColumn),LONG,PROC,VIRTUAL
GetConnectionInformation            PROCEDURE(STRING pConnectStr,*LoginConnections pConnectionList,BYTE pTrusted=0),LONG,PROC,PRIVATE  
GetFieldList                        PROCEDURE(*FILE pTbl),String,VIRTUAL
GetSQLNativeClientDriver            PROCEDURE(),STRING,VIRTUAL
HandleError                         PROCEDURE(LONG pErrorCode,STRING pError,LONG pFileErrorCode,STRING pFileError),VIRTUAL
ProcessScript                       PROCEDURE(STRING pScript),BYTE,PROC,VIRTUAL
Quote                               PROCEDURE(STRING pText,BYTE pEscape=0),STRING,VIRTUAL     
          
SetPRAGMA                           PROCEDURE(STRING pPragma),VIRTUAL

CreateTableList                     PROCEDURE(*FILE pTbl,<STRING pViewName>),VIRTUAL

Records                             PROCEDURE(*FILE pFile,<STRING pFilter>),LONG,VIRTUAL
RenameColumn                        PROCEDURE(STRING pTable,STRING pOldColumn,STRING pNewColumn),LONG,PROC,VIRTUAL

RemoveIllegalCharacters             PROCEDURE(String pString),STRING,VIRTUAL
Set                                 PROCEDURE(*KEY pKey, <STRING pSelect>, BYTE pReverse=False),VIRTUAL
Set                                 PROCEDURE(*KEY pKeyIgnored, *KEY pKey, <STRING pSelect>, BYTE pReverse=False),VIRTUAL
SetCatalog                          PROCEDURE(STRING pCatalog),VIRTUAL   
SetQueryConnection                  PROCEDURE(STRING pConnectionString,<STRING pQueryTableName>),VIRTUAL

TableExists                         PROCEDURE(FILE pFile,<STRING pCatalog>,<STRING pSchema>),BYTE,VIRTUAL
TableExists                         PROCEDURE(STRING pTable,<STRING pCatalog>,<STRING pSchema>),BYTE,VIRTUAL
ObjectExists                        PROCEDURE(STRING pObjectName,<STRING pSchema>),LONG,VIRTUAL
TestConnection                      PROCEDURE(STRING Server,STRING Database, STRING USR, STRING PWD, <Byte Trusted>, <*STRING ErrorOut>),BYTE,PROC,VIRTUAL
Trace                               PROCEDURE(*FILE pTbl, <STRING pLogfile>),VIRTUAL  
Truncate                            PROCEDURE(FILE pFile),LONG,PROC,VIRTUAL 
Truncate                            PROCEDURE(STRING pFileName),LONG,PROC,VIRTUAL 
      
PrepStatement                       PROCEDURE(STRING pStatement,BYTE pSingleQuote=0,BYTE pBracket=0),STRING,VIRTUAL

Insert                              PROCEDURE(*FILE pTbl, BYTE pIncludePK=False, BYTE pGetIdentity=FALSE), LONG, PROC,VIRTUAL
Update                              PROCEDURE(*FILE pTbl), BYTE, PROC,VIRTUAL
Delete                              PROCEDURE(*FILE pTbl), BYTE, PROC,VIRTUAL
GetInsertStatement                  PROCEDURE(*FILE pTbl), STRING, PROC,VIRTUAL 
GetUpdateStatement                  PROCEDURE(*FILE pTbl), STRING, PROC,VIRTUAL 
GetDeleteStatement                  PROCEDURE(*FILE pTbl), STRING, PROC,VIRTUAL 

Construct                           PROCEDURE()
Destruct                            PROCEDURE()       

Init                                PROCEDURE(LONG pDriver = ULSDriver_MSSQL,<STRING pProperties>,BYTE pUseRegistryToRetrieveServers=1),VIRTUAL
Kill                                PROCEDURE(),VIRTUAL

Query                               PROCEDURE (STRING pQuery, <*QUEUE pQ>, <*? pC1>, <*? pC2>, <*? pC3>, <*? pC4>, <*? pC5>, <*? pC6>, <*? pC7>, <*? pC8>, <*? pC9>, <*? pC10>, <*? pC11>, <*? pC12>, <*? pC13>, <*? pC14>, <*? pC15>, <*? pC16>, <*? pC17>,<*? pC18>, <*? pC19>, <*? pC20>, <*? pC21>, <*? pC22>, <*? pC23>, <*? pC24>, <*? pC25>, <*? pC26>, <*? pC27>, <*? pC28>, <*? pC29>, <*? pC30>, <*? pC31>, <*? pC32>, <*? pC33>, <*? pC34>, <*? pC35>, <*? pC36>, <*? pC37>, <*? pC38>, <*? pC39>, <*? pC40>, <*? pC41>, <*? pC42>),STRING,PROC,VIRTUAL
QueryDummy                          PROCEDURE (STRING pQuery, <*QUEUE pQ>, <*? pC1>, <*? pC2>, <*? pC3>, <*? pC4>, <*? pC5>, <*? pC6>, <*? pC7>, <*? pC8>, <*? pC9>, <*? pC10>, <*? pC11>, <*? pC12>, <*? pC13>, <*? pC14>, <*? pC15>, <*? pC16>, <*? pC17>,<*? pC18>, <*? pC19>, <*? pC20>, <*? pC21>, <*? pC22>, <*? pC23>, <*? pC24>, <*? pC25>, <*? pC26>, <*? pC27>, <*? pC28>, <*? pC29>, <*? pC30>, <*? pC31>, <*? pC32>, <*? pC33>, <*? pC34>, <*? pC35>, <*? pC36>, <*? pC37>, <*? pC38>, <*? pC39>, <*? pC40>, <*? pC41>, <*? pC42>),BYTE,PROC,VIRTUAL
QueryODBC                           PROCEDURE (STRING pQuery, <*QUEUE pQ>, <*? pC1>, <*? pC2>, <*? pC3>, <*? pC4>, <*? pC5>, <*? pC6>, <*? pC7>, <*? pC8>, <*? pC9>, <*? pC10>, <*? pC11>, <*? pC12>, <*? pC13>, <*? pC14>, <*? pC15>, <*? pC16>, <*? pC17>,<*? pC18>, <*? pC19>, <*? pC20>, <*? pC21>, <*? pC22>, <*? pC23>, <*? pC24>, <*? pC25>, <*? pC26>, <*? pC27>, <*? pC28>, <*? pC29>, <*? pC30>, <*? pC31>, <*? pC32>, <*? pC33>, <*? pC34>, <*? pC35>, <*? pC36>, <*? pC37>, <*? pC38>, <*? pC39>, <*? pC40>, <*? pC41>, <*? pC42>),BYTE,PROC,VIRTUAL
QueryResult                         PROCEDURE (STRING pQuery),STRING,VIRTUAL 
QueryToJson                         PROCEDURE (STRING pQuery,BYTE pJsonForOption = us:ForJsonAuto),STRING,VIRTUAL

! RRS 01/22/19 Added to support large tables.
QueryCT                             PROCEDURE (STRING pQuery, <*QUEUE pQ>, <*ctFieldQ pFields>), BYTE, PROC,VIRTUAL
! RRS 01/22/19 End.

ExtendedProperty_Insert             PROCEDURE(STRING pObjectName,STRING pPropertyName,STRING pPropertyValue), LONG, PROC,VIRTUAL
ExtendedProperty_Update             PROCEDURE(STRING pObjectName,STRING pPropertyName,STRING pPropertyValue), LONG, PROC,VIRTUAL
ExtendedProperty_Delete             PROCEDURE(STRING pObjectName,STRING pPropertyName), LONG, PROC,VIRTUAL
ExtendedProperty_GetValue           PROCEDURE(STRING pObjectName,STRING pPropertyName), STRING,VIRTUAL
ExtendedProperty_Exists             PROCEDURE(STRING pObjectName,STRING pPropertyName), BYTE,VIRTUAL
GetAllExtendedProperties            PROCEDURE(STRING pObjectName),VIRTUAL

AppendToDriverOptions               PROCEDURE()   ,VIRTUAL  

SetCustomConnectionString           PROCEDURE(STRING pCustomConnectionString),VIRTUAL
SetDriverOptionDefaults             PROCEDURE(),VIRTUAL
SetBusyHandling                     PROCEDURE(BYTE pBusyHandling) ,VIRTUAL
SetBusyRetries                      PROCEDURE(LONG pBusyRetries),VIRTUAL 
SetGatherAtOpen                     PROCEDURE(BYTE pTrueFalse),VIRTUAL 
SetIgnoreTruncation                 PROCEDURE(BYTE pTrueFalse),VIRTUAL 
SetLogonScreen                      PROCEDURE(BYTE pTrueFalse),VIRTUAL 
SetMultipleActiveResultSets         PROCEDURE(BYTE pTrueFalse),VIRTUAL 
SetSaveStoredProcedure              PROCEDURE(BYTE pTrueFalse),VIRTUAL   
SetTurboSQL                         PROCEDURE(BYTE pTrueFalse),VIRTUAL 
SetVerifyViaSelect                  PROCEDURE(BYTE pTrueFalse),VIRTUAL  

SendDriverString                    PROCEDURE(STRING pMessage),STRING,PROC ,VIRTUAL

SetDriver                           PROCEDURE(LONG pDriver = ULSDriver_MSSQL),VIRTUAL

SetAllConnectionStrings             PROCEDURE(STRING pServer,STRING TheDatabase,STRING TheUserName,STRING ThePassword,BYTE IsTrusted),STRING,PROC,VIRTUAL

GetAutoFill                         PROCEDURE(*CSTRING pServer,*CSTRING pOwnerName,*BYTE pWindowsAuthentication),VIRTUAL
SaveAutoFill                        PROCEDURE(*CSTRING pServer,*CSTRING pOwnerName,*BYTE pWindowsAuthentication),VIRTUAL

StripSchema                         PROCEDURE(STRING pSchema,STRING pTable),STRING   ,VIRTUAL

CheckClarionTablePrefixes           PROCEDURE (STRING pQuery),STRING,VIRTUAL
CheckForODBCDriver                  PROCEDURE(),BYTE,VIRTUAL   

CreateTemporaryProcedure            PROCEDURE (STRING pQuery),VIRTUAL   
DropTemporaryProcedure              PROCEDURE (STRING pTemporaryProcedureName),VIRTUAL   

GetLocalSQLServers                  PROCEDURE(),VIRTUAL
          
ClearErrors                         PROCEDURE(),VIRTUAL
SetError                            PROCEDURE(LONG pErrorCode,STRING pErrorDescription),VIRTUAL 
GetError                            PROCEDURE(<STRING pStatementID>),STRING,VIRTUAL 
       
AddStatement                        PROCEDURE(STRING pStatement),VIRTUAL
AddStatement                        PROCEDURE(LONG pStatementID,STRING pStatement),VIRTUAL
ExecuteStatements                   PROCEDURE(LONG pStatementID=0),STRING,PROC,VIRTUAL  
CreateDBErrorsTable                 PROCEDURE(STRING pBatchGUID),VIRTUAL 
DropDBErrorsTable                   PROCEDURE(STRING pBatchGUID),VIRTUAL 
RemoveBatchErrors                   PROCEDURE(STRING pBatchGUID),VIRTUAL

CheckForDebug                       PROCEDURE(STRING pQueryToCheck),STRING
       
BeginTransaction                    PROCEDURE()
EndTransaction                      PROCEDURE()

Wait                                PROCEDURE(Long pId),Virtual
Release                             PROCEDURE(Long pId),Virtual
Trace                               PROCEDURE(string pStr),Virtual      

                                END    


    
