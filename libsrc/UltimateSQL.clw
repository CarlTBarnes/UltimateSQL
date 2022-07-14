!region Begin Comments
!!>Author          Members of the Clarion Community: John Hickey, Arnold Young,
!!>                Rick Martin, Andy Wilton, Mike Hanson, Mark Goldberg, Rick
!!>                Smith, Carl Barnes
!!>Version         2022.2.8.1
!!>Template        UltimateSQL.tpl
!!>Revisions       Created Help file
!!>                Added comments to methods
!!>Created         DEC  3,2020,  5:40:11am
!!>Modified        JAN  9,2022,  1:39:18pm
!!>See end of file for additional notes
!endregion End Comments
                                MEMBER()   
                                pragma('link(C%V%MSS%X%%L%.LIB)')
                                pragma('link(C%V%ODB%X%%L%.LIB)')

    INCLUDE('EQUATES.CLW')
    INCLUDE('UltimateSQL.INC'),ONCE                                             
    INCLUDE('UltimateSQLString.INC'),ONCE
    INCLUDE('UltimateSQLScripts.INC'),ONCE   
    Include('CWSYNCHM.INC'),ONCE  

                                MAP
                                    Module('ODBC32')
                                        USQLAllocHandle(short HandleType, Long InputHandle, *Long OutputHandle), Short, Pascal ,NAME('SQLAllocHandle'),PROC
                                        USQLBrowseConnect(Long hdbc, *CString ConnectStringIn, Short CSInSize, *CString ConnectStringOut, Short BufferSize, *Short ReturnSize), Short, Pascal, Raw ,NAME('SQLBrowseConnect'), PROC
                                        USQLDisconnect(Long hdbc), Short, Pascal ,NAME('SQLDisconnect'), PROC
                                        USQLFreeHandle(Short HandleType, Long Handle), Short, Pascal ,NAME('SQLFreeHandle'), PROC
                                        USQLSetConnectAttr(Long DBCHandle, Long Attribute, Long ValuePtr, Long ValueLen), Short, Pascal ,NAME('SQLSetConnectAttr'), PROC
                                        USQLSetEnvAttr(Long EnvironmentHandle, Long Attribute, Long ValuePtr, Long ValueLen), Short, Pascal ,NAME('SQLSetEnvAttr'), PROC
                                    END
                                END
                         
SQLiteMemoryDatabase            FILE,DRIVER('SQLite','/TURBOSQL=True'),OWNER(':memory:'),PRE(SQLiteMemoryDatabase),BINDABLE,THREAD,CREATE
Record                              RECORD,PRE()
Field1                                  STRING(100)
Field2                                  STRING(100)
Field3                                  STRING(100)

                                    END   
                                END



DatabaseCheckOwnerString        STRING(200)
TestConnectionString            STRING(200)   
  
DatabaseConnectionString        STRING(200)

eVersion                        EQUATE(2)  
DriverOptions                   STRING(1000)
                             
QueryResultsTemp                FILE,DRIVER('MSSQL',DriverOptions),PRE(QueryResultsTemp),BINDABLE,THREAD
Record                              RECORD,PRE()
Temp                                    STRING(1)
                                    END   
                                END 
 

DebugInitted                    BYTE(FALSE)    

QueryFieldsGroup                GROUP,TYPE
C01                                 CSTRING(600000)
C02                                 CSTRING(8000)
C03                                 CSTRING(8000)
C04                                 CSTRING(8000)
C05                                 CSTRING(8000)
C06                                 CSTRING(8000)
C07                                 CSTRING(8000)
C08                                 CSTRING(8000)
C09                                 CSTRING(8000)
C10                                 CSTRING(8000)
C11                                 CSTRING(8000)
C12                                 CSTRING(8000)
C13                                 CSTRING(8000)
C14                                 CSTRING(8000)
C15                                 CSTRING(8000)
C16                                 CSTRING(8000)
C17                                 CSTRING(8000)
C18                                 CSTRING(8000)
C19                                 CSTRING(8000)
C20                                 CSTRING(8000)
C21                                 CSTRING(8000)
C22                                 CSTRING(8000)
C23                                 CSTRING(8000)
C24                                 CSTRING(8000)
C25                                 CSTRING(8000)
C26                                 CSTRING(8000)
C27                                 CSTRING(8000)
C28                                 CSTRING(8000)
C29                                 CSTRING(8000)
C30                                 CSTRING(8000)
C31                                 CSTRING(8000)
C32                                 CSTRING(8000)
C33                                 CSTRING(8000)
C34                                 CSTRING(8000)
C35                                 CSTRING(8000)
C36                                 CSTRING(8000)
C37                                 CSTRING(8000)
C38                                 CSTRING(8000)
C39                                 CSTRING(8000)
C40                                 CSTRING(8000)
C41                                 CSTRING(8000)
C42                                 CSTRING(8000)
                                END    


! -----------------------------------------------------------------------
UltimateSQL.CheckForODBCDriver          PROCEDURE()
! -----------------------------------------------------------------------
        
Result                                      BYTE(0)
       
DIWindow                                    WINDOW,AT(,,260,64),CENTER,GRAY,FONT('Segoe UI',12), |
                                                    COLOR(COLOR:White),DOUBLE
                                                STRING('Installing required ODBC driver, one moment please...'),AT(46,27), |
                                                        USE(?STRING1)
                                            END

    CODE
   
    Result  =  TRUE
    
    LOOP
        IF SELF._Driver = ULSDriver_ODBC ! AND ~SELF.RequireNativeClient 
            SELF.NativeClient  =  SELF.ODBCDriver
            IF GETREG(REG_LOCAL_MACHINE,'SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers',CLIP(SELF.ODBCDriver)) <> 'Installed'  
                IF SELF.ODBCSilentInstall   
                    IF SELF.ODBCSilentInstall = 2      
                        OPEN(DIWindow) 
                        DISPLAY()
                    END
                    
                    RUN('msiexec /quiet /passive /qn /i ' & CLIP(SELF.ODBCDriverMSIFileLocation) & ' IACCEPTMSODBCSQLLICENSETERMS=YES',1)  
                    
                    IF SELF.ODBCSilentInstall = 2      
                        CLOSE(DIWindow)
                    END    
                    
                    SELF.ODBCSilentInstall  =  FALSE  
                    CYCLE 
                    
                ELSE
                    IF SELF.NoODBCDriverMessage = ''
                        SELF.NoODBCDriverMessage  =  'The ODBC Driver ' & CLIP(SELF.ODBCDriver) & ' has not been installed.|This is required.'
                    END
                    MESSAGE(SELF.NoODBCDriverMessage)
                    Result  =  FALSE
                    BREAK
                END
                
            ELSE
                BREAK
                
            END
        ELSE
            Result  =  TRUE
            BREAK
            
        END
        
    END
         
    RETURN Result
    
    
! -----------------------------------------------------------------------
!!! <summary>Prompts for SQL Connection Information.</summary>
!!! <param name="*STRING">Server Name</param>
!!! <param name="*STRING">User Name</param>
!!! <param name="*STRING">Password</param>
!!! <param name="*STRING">Database</param>
!!! <param name="*BYTE">*BYTE pTrusted</param>
!!! <param name="BYTE"><Optional> True is a Trusted Connection, False if not.</param>
!!! <param name="BYTE"><Optional> LoginNamePasswordOnly.  You will be prompted only for name and password.  You should pass server and database information.</param>
!!! <param name="STRING"><Optional> Instructions that will be displayed on the Login screen if displayed.</param>
!!! <returns>Returns the connection string</returns>
!!! <remarks>Parameters are passed in by pointer, and will be filled with entered values when the procedure exits.</remarks>
UltimateSQL.Connect             PROCEDURE(*STRING pServer,*STRING pUserName,*STRING pPassword,*STRING pDatabase,*BYTE pTrusted,<BYTE pLoginNamePasswordOnly>,BYTE pForce=0,<STRING pInstructions>)  ! ,STRING

TheServer                           CSTRING(200)
TheUserName                         CSTRING(200)
ThePassword                         CSTRING(200) 
TheDatabase                         CSTRING(200)
Trusted                             BYTE(0)

TheResult                           STRING(800)

ConnectStr                          STRING(100)  
TestResult                          STRING(800)   

SQLServers                          QUEUE,PRE(SQLServers)               
Name                                    STRING(512)                           
                                    END
SQLDatabases                        QUEUE,PRE(SQLDatabases)               
Name                                    STRING(512)                           
                                    END
                                                  
Scripts                             UltimateSQLScripts
NativeClient                        STRING(50) 

LocalCount                          LONG
LocalOrNetwork                      STRING(10)

                                    INCLUDE('UltimateSQLConnectWindow.clw')
    
    CODE 
    
    TheResult          =  ''            
    SELF.NativeClient  =  ''
    
    SELF.NativeClient  =  SELF.GetSQLNativeClientDriver()
    IF SELF.NativeClient = ''  
        IF SELF.NoNativeClientMessage = ''
            SELF.NoNativeClientMessage  =  'No SQL Native Client Driver is installed on this machine.|This is necessary for the program to run properly.'
        END
        MESSAGE(SELF.NoNativeClientMessage,'Error',ICON:Hand)
        DO ProcedureReturn
        
    END                                                                                                                       
    
    IF SELF.ODBCDriver
        IF ~SELF.CheckForODBCDriver() 
            DO ProcedureReturn
        END 
        
    END
    
    TheServer    =  CLIP(pServer)
    TheDatabase  =  CLIP(pDatabase)
    TheUserName  =  CLIP(pUserName)
    ThePassword  =  CLIP(pPassword)
    Trusted      =  pTrusted
     
    LocalOrNetwork = 'Local'
    
    SELF.GetAutoFill(TheServer,TheUserName,Trusted)
    
    IF RECORDS(SELF.qSQLServerList) 
        FREE(SQLServers)
        LOOP LocalCount = 1 TO RECORDS(SELF.qSQLServerList)
            GET(SELF.qSQLServerList,LocalCount)
            SQLServers  =  SELF.qSQLServerList
            ADD(SQLServers)
        END
        
    END
    
    IF ((TheServer AND TheDatabase AND Trusted) OR (TheServer AND TheDatabase AND ThePassword AND TheUserName AND ~Trusted)) AND ~pForce       
        IF SELF.TestConnection(TheServer,TheDatabase,TheUserName,ThePassword,Trusted)

            TheResult  =  SELF.SetAllConnectionStrings(TheServer,TheDatabase,TheUserName,ThePassword,Trusted)
            DO ProcedureReturn
            
        END 
        
    END  
    
    OPEN(Window)              
    0{PROP:Hide        }         =  TRUE
    ?SHEET{PROP:Wizard}          =  TRUE  
    ?SHEET{PROP:NoSheet}         =  TRUE
    ?TheServer{PROP:LineHeight}  =  12 
    
    IF pInstructions
        ?PromptInstructions{PROP:Text}  =  pInstructions
    END
    
    IF pLoginNamePasswordOnly
        ?TheDatabase{PROP:Drop}  =  0
    END
    
    DISPLAY() 
    
    SELECT(?SHEET,2)
    0{PROP:Hide}  =  FALSE
    EXECUTE Trusted + 1
        ?LISTAuthentication{PROP:Selected}  =  2   
        ?LISTAuthentication{PROP:Selected}  =  1
    END  
    
    DO SetUserPasswordFields

    ACCEPT
        CASE FIELD() 
        OF ?TheServer
            CASE EVENT()
            OF EVENT:DroppingDown
!                IF RECORDS(SQLServers) < 2
!                    ConnectStr  =  'Driver=SQL Server;'
!                    SELF.GetConnectionInformation(ConnectStr, SQLServers,CHOOSE(?LISTAuthentication{PROP:Selected}=1,1,0))
!                    DISPLAY()
!                END
                CLEAR(SQLDatabases)
                FREE(SQLDatabases) 
                
            OF EVENT:Accepted
                CLEAR(SQLDatabases)
                FREE(SQLDatabases)
                IF TheServer = '<Browse for more...>'   
                    ConnectStr  =  'Driver=SQL Server;'
                    SELF.GetConnectionInformation(ConnectStr, SQLServers,CHOOSE(?LISTAuthentication{PROP:Selected}=1,1,0))  
                    TheServer = '' 
                    UPDATE()
                    DISPLAY()
                    CLEAR(SQLDatabases)
                    FREE(SQLDatabases) 
                END
                
            END 
        
        OF ?LocalOrNetwork
            CASE EVENT()
            OF EVENT:Accepted   
                UPDATE()
                DISPLAY()
                CLEAR(SQLDatabases)
                FREE(SQLDatabases) 
                IF LocalOrNetwork = 'Local' 
                    FREE(SQLServers)
                    LOOP LocalCount = 1 TO RECORDS(SELF.qSQLServerList)
                        GET(SELF.qSQLServerList,LocalCount)
                        SQLServers  =  SELF.qSQLServerList
                        ADD(SQLServers)
                    END
                    
                ELSE  
                    ConnectStr  =  'Driver=SQL Server;'
                    SELF.GetConnectionInformation(ConnectStr, SQLServers,CHOOSE(?LISTAuthentication{PROP:Selected}=1,1,0))  
                    TheServer = '' 
                    
                END
            END
            
        OF ?LISTAuthentication
            CASE EVENT()
            OF EVENT:Accepted
                DO SetUserPasswordFields  
                
            END
            
        OF ?TheDatabase
            CASE EVENT()
            OF EVENT:DroppingDown
                If TheServer AND ~pLoginNamePasswordOnly
                    ConnectStr  =  CLIP(TheServer) & ',master,' & CLIP(TheUserName) & ',' & CLIP(ThePassword) 
                    IF ?LISTAuthentication{PROP:Selected}=1
                        ConnectStr  =  CLIP(TheServer) & ',master;TRUSTED_CONNECTION=Yes'   
                    END
                    SELF.ConnectionString  =  ConnectStr  
                    SELF.SetAllConnectionStrings(TheServer,TheDatabase,TheUserName,ThePassword,CHOOSE(?LISTAuthentication{PROP:Selected}=1,1,0))
                    SELF.QueryMethod  =  QueryMethodODBC
                    SELF.Query('SELECT name FROM master.dbo.sysdatabases Order By Name',SQLDatabases,SQLDatabases.Name)
                    Display()
                END   
                
            END  
            
        OF ?ButtonTest  
            CASE EVENT()
            OF EVENT:Accepted
                TheResult  =  SELF.TestConnection(TheServer,|
                        CHOOSE(pLoginNamePasswordOnly,'master',TheDatabase),|
                        TheUserName,|
                        ThePassword,|
                        CHOOSE(?LISTAuthentication{PROP:Selected}=1,1,0),TestResult)  
                MESSAGE(CLIP(TestResult))
                CYCLE 
                
            END
            
        OF ?OkButton  
            CASE EVENT()
            OF EVENT:Accepted
                IF ?LISTAuthentication{PROP:Selected}=2 AND (TheUserName = '' OR ThePassword = '')
                    MESSAGE('Authentication is turned on, but you have not supplied a User Name and/or Password.')
                    CYCLE
                END
                
                TheResult  =  CLIP(TheServer) & ',' & CLIP(TheDatabase) & ',' & CLIP(TheUserName) & ',' & CLIP(ThePassword) 
                IF ?LISTAuthentication{PROP:Selected}=1
                    TheResult  =  CLIP(TheServer) & ',' & CLIP(TheDatabase) & ';TRUSTED_CONNECTION=Yes'   
                END   
!!!                IF SELF.ApplicationName
!!!                    TheResult = CLIP(TheResult) & ';APP=' & CLIP(SELF.ApplicationName)
!!!                END
                
                BREAK
                
            END
            
        OF ?CancelButton
            CASE EVENT()
            OF EVENT:Accepted
                TheResult  =  '0'
                BREAK
            END
            
        END
        
    END 
    
    Trusted  =  0
    IF ?LISTAuthentication{PROP:Selected}=1
        Trusted  =  1 
        
    END
    
    DO ProcedureReturn
    
    
ProcedureReturn                 ROUTINE  
    
    DATA

TheTestResult   BYTE(0)
IsTrusted       BYTE(0)


    CODE
    
    IF ~TheResult
    ELSE
        SELF.Catalog  =  TheDatabase 
        
        pServer                 =  TheServer
        pDatabase               =  TheDatabase
        pUserName               =  TheUserName
        pPassword               =  ThePassword   
        SELF.Server             =  TheServer
        SELF.Database           =  TheDatabase
        SELF.User               =  TheUserName
        SELF.Password           =  ThePassword
        IsTrusted               =  Trusted
        SELF.TRUSTEDCONNECTION  =  IsTrusted
        pTrusted                =  IsTrusted
        TheTestResult           =  SELF.TestConnection(TheServer,CHOOSE(pLoginNamePasswordOnly,'master',TheDatabase),TheUserName,ThePassword,IsTrusted,TestResult)
        IF ~TheTestResult
            TheResult  =  ''
        ELSE
            SELF.ConnectionString  =  TheResult  
            SELF.SetAllConnectionStrings(pServer,CHOOSE(pLoginNamePasswordOnly,'master',TheDatabase),TheUserName,ThePassword,IsTrusted)
            
            
        END 
        
        SELF.SaveAutoFill(TheServer,TheUserName,IsTrusted)
        
    END
    
        
    RETURN TheResult
    
SetUserPasswordFields           ROUTINE
    
    ?TheUserName{PROP:Disable   } =  CHOOSE(?LISTAuthentication{PROP:Selected}=1,1,0)
    ?TheUserName{PROP:Background}  =  CHOOSE(?LISTAuthentication{PROP:Selected}=1,COLOR:BTNFACE,COLOR:NONE)
    ?ThePassword{PROP:Disable   } =  CHOOSE(?LISTAuthentication{PROP:Selected}=1,1,0)        
    ?ThePassword{PROP:Background}  =  CHOOSE(?LISTAuthentication{PROP:Selected}=1,COLOR:BTNFACE,COLOR:NONE)
    

! -----------------------------------------------------------------------        
UltimateSQL.SetAllConnectionStrings     PROCEDURE(STRING TheServer,STRING TheDatabase,STRING TheUserName,STRING ThePassword,BYTE IsTrusted)
! -----------------------------------------------------------------------        

Provider                                    STRING(20)

    CODE        
    
    IF ~SELF.InitHasRun
        SELF.Init()
    END
    
    IF SELF.Provider = 'SQL Native Client'
        Provider  =  'SQLNCLI.1' 
    ELSIF SELF.Provider = 'None'
        SELF.NativeClient  =  ''
        
    END
    
    IF IsTrusted
        SELF.FullConnectionString  =  'Server=' & CLIP(TheServer) & ';Database=' & CLIP(TheDatabase) & ';TRUSTED_CONNECTION=Yes' & |
                CHOOSE(SELF.NativeClient='','',';Driver={{' & CLIP(SELF.NativeClient) & '}') & |
                CHOOSE(SELF.ApplicationName = '','',';app=' & CLIP(SELF.ApplicationName) ) & |
                CHOOSE(SELF.WSID = '','',';WSID=' & CLIP(SELF.WSID)) & ';'
        
        SELF.ConnectionStringWithProvider  =  'Provider=' & CLIP(Provider) & ';Persist Security Info=True;;TRUSTED_CONNECTION=Yes;Initial Catalog=' & CLIP(TheDatabase) & |
                ';Data Source=' & CLIP(TheServer) & |
                CHOOSE(SELF.ApplicationName = '','',';app=' & CLIP(SELF.ApplicationName)) & |
                CHOOSE(SELF.WSID = '','',';WSID=' & CLIP(SELF.WSID)) & ';'    
        
        SELF.ConnectionString  =  CLIP(TheServer) & ',' & CLIP(TheDatabase) & ';TRUSTED_CONNECTION=Yes' & |
                CHOOSE(SELF.NativeClient='','',';Driver={{' & CLIP(SELF.NativeClient) & '}') & |
                CHOOSE(SELF.ApplicationName = '','',';app=' & CLIP(SELF.ApplicationName)) & |
                CHOOSE(SELF.WSID = '','',';WSID=' & CLIP(SELF.WSID)) & ';'   
        
    ELSE  
        SELF.FullConnectionString  =  'Server=' & CLIP(TheServer) & ';Database=' & CLIP(TheDatabase) & ';Uid=' & CLIP(TheUserName) & ';Pwd=' & CLIP(ThePassword) & |
                CHOOSE(SELF.NativeClient='','',';Driver={{' & CLIP(SELF.NativeClient) & '}') & |
                CHOOSE(SELF.ApplicationName = '','',';app=' & CLIP(SELF.ApplicationName)) & |
                CHOOSE(SELF.WSID = '','',';WSID=' & CLIP(SELF.WSID)) & ';'   
        
        SELF.ConnectionStringWithProvider  =  'Provider=' & CLIP(Provider) & ';Password=' & CLIP(ThePassword) & ';Persist Security Info=True;User ID=' & CLIP(TheUserName) & |
                ';Initial Catalog=' & CLIP(TheDatabase) & ';Data Source=' & CLIP(TheServer) & |
                CHOOSE(SELF.NativeClient='','',';Driver={{' & CLIP(SELF.NativeClient) & '}') & |
                CHOOSE(SELF.ApplicationName = '','','app=' & CLIP(SELF.ApplicationName)) & |
                CHOOSE(SELF.WSID = '','',';WSID=' & CLIP(SELF.WSID)) & ';' 
        
        SELF.ConnectionString  =  CLIP(TheServer) & ',' & CLIP(TheDatabase) & ',' & CLIP(TheUserName) & ',' & CLIP(ThePassword) & |
                CHOOSE(SELF.NativeClient='','',';Driver={{' & CLIP(SELF.NativeClient) & '}') & |
                CHOOSE(SELF.ApplicationName = '','',';app=' & CLIP(SELF.ApplicationName)) & |
                CHOOSE(SELF.WSID = '','',';WSID=' & CLIP(SELF.WSID)) & ';'
            
    END                                                    
    
    IF SELF._Driver = ULSDriver_ODBC
        SELF.ConnectionString  =  SELF.FullConnectionString
    END
    
    SELF.Server             =  TheServer
    SELF.Database           =  TheDatabase
    SELF.User               =  TheUserName
    SELF.Password           =  ThePassword
    SELF.TRUSTEDCONNECTION  =  IsTrusted   
    
    SELF.Catalog  =  TheDatabase                                                                                
    
    RETURN SELF.ConnectionString    

    
! -----------------------------------------------------------------------
!!! <summary>Gets a list of MSSQL Servers or Databases</summary>           
!!! <param name="ConnectStr">The Connection string</param>
!!! <param name="LoginConnections">The LoginConnections data type, a Queue to hold Server or Database names</param>        
!!! <param name="Trusted">Whether or not this is a Trusted connection.  TRUE = Trusted</param>        
! -----------------------------------------------------------------------        
UltimateSQL.GetConnectionInformation    PROCEDURE(STRING pConnectStr,*LoginConnections pConnectionList,BYTE pTrusted=0)
                                        

SQLReturn                                   Long
EnvHandle                                   Long
DBCHandle                                   Long
ConnectStrIn                                CString(2000)
ConnectStrOut                               CString(2000)
ReturnSize                                  Short
TestStringIN                                CString(50)
TestString                                  CString(50)
StartPos                                    Short
EndPos                                      Short

    CODE                                                     ! Begin processed code
                                              
    SetCursor(CURSOR:Wait)
    ConnectStrIn  =  pConnectStr
    SQLReturn     =  USQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, EnvHandle)
        
    If SQLReturn >= 0 Then !If No Error
        SQLReturn  =  USQLSetEnvAttr(EnvHandle, SQL_ATTR_ODBC_VERSION, SQL_OV_ODBC3, 0)
        If SQLReturn >= 0 Then   !If No Error
            SQLReturn  =  USQLAllocHandle(SQL_HANDLE_DBC, EnvHandle, DBCHandle)
            If SQLReturn >= 0 Then !If No Error
                If pTrusted Then
                    SQLReturn  =  USQLSetConnectAttr(DBCHandle, SQL_COPT_SS_INTEGRATED_SECURITY, 1, 0)
                Else
                    SQLReturn  =  USQLSetConnectAttr(DBCHandle, SQL_COPT_SS_INTEGRATED_SECURITY, 0, 0)
                End
                If SQLReturn >= 0 Then !If No Error
                    SQLReturn  =  USQLBrowseConnect(DBCHandle, ConnectStrIn, Size(ConnectStrIn), ConnectStrOut, Size(ConnectStrOut), ReturnSize)
                    If SQLReturn >= 0 Then !If Error Browsing Network
                        TestStringIn  =  'SERVER='
                        StartPos      =  Instring(TestStringIn, UPPER(ConnectStrIn), 1, 1)
                        If StartPOS = 0 Then
                            TestString  =  'SERVER={{'
                        Else
                            TestString  =  'DATABASE={{'
                        End

                        StartPos  =  Instring(TestString, UPPER(ConnectStrOut), 1, 1)

                        If StartPOS > 0 Then
                            StartPOS      +=  Len(TestString)
                            ConnectStrOut  =  Sub(ConnectStrOut, StartPos, Len(ConnectStrOut))
                            EndPOS         =  Instring('}', UPPER(ConnectStrOut), 1, 1) - 1
                            ConnectStrOut  =  Sub(ConnectStrOut, 1, EndPOS)

                            Free(pConnectionList)

                            Loop Until Len(Clip(ConnectStrOut)) <= 0
                                I# += 1
                                Clear(pConnectionList)

                                EndPOS  =  Instring(',', UPPER(ConnectStrOut), 1, 1)
                                If EndPOS > 0 Then
                                    pConnectionList.Name  =  Clip(Sub(ConnectStrOut, 1, EndPOS - 1))
                                    ConnectStrOut         =  Sub(ConnectStrOut, EndPOS + 1, Len(ConnectStrOut))

                                Else
                                    pConnectionList.Name  =  Clip(ConnectStrOut)
                                    ConnectStrOut         =  ''
                                End
                                Add(pConnectionList, pConnectionList.Name)
                            End  
                        End  
                    End
                    USQLDisconnect(DBCHandle)
                End
            End
            USQLFreeHandle(SQL_HANDLE_DBC, DBCHandle)
        End  
        USQLFreeHandle(SQL_HANDLE_ENV, ENVHandle)
    End 
    
    SETCURSOR()
    SORT(pConnectionList,pConnectionList.Name)
    RETURN SQLReturn
   
          
UltimateSQL.SetQueryConnection          PROCEDURE(STRING pConnectionString,<STRING pQueryTableName>) !Initialize the Connection String, optionally set the name of the Query Table in your SQL database 

szSQL                                       CSTRING(501)

    CODE	
    
    IF pConnectionString = ''
        RETURN
    END                                     
    
    SELF.ConnectionString  =  pConnectionString  
    SELF.ConnectionString  =  pConnectionString
    IF pQueryTableName = ''
        pQueryTableName  =  'Queries'
    END
    SELF.QueryTableName  =  pQueryTableName          

    
UltimateSQL.GetSQLNativeClientDriver    PROCEDURE()   !,STRING
 
DriverString                                UltimateString

    CODE                                   
              
    
    DriverString.Assign(GetReg(REG_CLASSES_ROOT,'SQLNCLI13')) 
    
    if not DriverString.Length()
        DriverString.Assign(GetReg(REG_CLASSES_ROOT,'SQLNCLI12'))
    end
    
    if not DriverString.Length()
        DriverString.Assign(GetReg(REG_CLASSES_ROOT,'SQLNCLI11'))
    end
    
    if not DriverString.Length()
        DriverString.Assign(GetReg(REG_CLASSES_ROOT,'SQLNCLI10'))
    end 
    
    if not DriverString.Length()
        DriverString.Assign(GetReg(REG_CLASSES_ROOT,'SQLNCLI'))
    end
  
    return DriverString.Get()
    

! -----------------------------------------------------------------------
!!! <summary>Tests to make sure an MSSQL connection is valid/summary>           
!!! <param name="Server">Server Name</param>
!!! <param name="Database">Database Name</param>
!!! <param name="Username">User Name</param>        
!!! <param name="Password">Password</param>   
!!! <param name="Trusted">Whether or not this is a Trusted connection.  TRUE = Trusted</param>        
!!! <param name="ErrorOut">If included, the results of the test are returned in the passed String</param>        
!!! <param name="Return Value">Returns TRUE if successful, FALSE if unsuccessful</param>        
! -----------------------------------------------------------------------        
UltimateSQL.TestConnection      PROCEDURE(STRING Server,STRING Database, STRING USR, STRING PWD, <BYTE Trusted>, <*STRING ErrorOut>) ! BYTE

StatusString                        STRING(2000)
FileErr                             STRING(512)
SQLStr                              STRING(1024)
DBCount                             LONG
ReturnVal                           BYTE

SysDatabases                        FILE,DRIVER('MSSQL'), Name('SysDatabases')
Record                                  RECORD,PRE()
name                                        CSTRING(129)
dbid                                        SHORT
sid                                         STRING(85)
mode                                        SHORT
status                                      LONG
status2                                     LONG
crdate                                      STRING(8)
crdate_GROUP                                GROUP,OVER(crdate)
crdate_DATE                                     DATE
crdate_TIME                                     TIME
                                            END
reserved                                    STRING(8)
reserved_GROUP                              GROUP,OVER(reserved)
reserved_DATE                                   DATE
reserved_TIME                                   TIME
                                            END
category                                    LONG
cmptlevel                                   BYTE
filename                                    CSTRING(261)
version                                     LONG
                                        END
                                    END   
TurboSQLTable                       FILE,DRIVER('MSSQL','/TURBOSQL=True'), pre(TurboSQL) 
Record                                  RECORD 
Variable                                    LONG 
                                        END
                                    END
    CODE                                                      
    
    !True   =   Connection Successful
    !False  =   Connection Failed      
     
    StatusString  =  'BEGIN TEST...<13><10><13><10>'
    If Omitted(4) Then
        Trusted  =  False
    End

    ReturnVal  =  True

    If Len(Clip(Server)) <= 0 Then
        ReturnVal  =  False
    End

    If (Len(Clip(USR)) <= 0) AND (~Trusted) Then
        ReturnVal  =  False
    End

    If ReturnVal And Trusted Then
        Send(SysDatabases, '/TRUSTEDCONNECTION = TRUE')
    Else
        Send(SysDatabases, '/TRUSTEDCONNECTION = FALSE')
    End

    StatusString  =  Clip(StatusString) & 'Testing Provided Values...'

    If ReturnVal Then
        StatusString  =  Clip(StatusString) & 'SUCCESS!<13><10>'
        StatusString  =  Clip(StatusString) & 'Attempting Connection...'
        TurboSQLTable{PROP:Owner     } =  Clip(Server) & ',master,' & Clip(USR) & ',' & Clip(PWD)
        SysDatabases{PROP:Owner     } =  Clip(Server) & ',master,' & Clip(USR) & ',' & Clip(PWD)
        SysDatabases{Prop:Logonscreen}  =  False 
        Open(SysDatabases)

        If Error() Then
            FileErr  =  'Error['
            If FileError() Then
                FileErr  =  Clip(FileErr) & FileErrorCode() & ']: ' & FileError()
            Else
                FileErr  =  Clip(FileErr) & ErrorCode() & ']: ' & Error()
            End
            StatusString  =  Clip(StatusString) & 'FAILED!<13><10>' & Clip(FileErr) & '<13><10>'
            ReturnVal     =  False
        End

        If ReturnVal Then
            StatusString  =  Clip(StatusString) & 'SUCCESS!<13><10>'
            StatusString  =  Clip(StatusString) & 'Running Simple Query...'
                                                         
            IF Database
                SQLStr  =  'SELECT COUNT(*) FROM SYSDATABASES Where name = ' & SELF.Quote(Database)
            ELSE
                SQLStr  =  'SELECT COUNT(*) FROM SYSDATABASES'
            END
            Open(TurboSQLTable) 
            TurboSQLTable{Prop:SQL}  =  SQLStr
            If Error() Then
                FileErr  =  'Error['
                If FileError() Then
                    FileErr  =  Clip(FileErr) & FileErrorCode() & ']: ' & FileError()
                Else
                    FileErr  =  Clip(FileErr) & ErrorCode() & ']: ' & Error()
                End
                StatusString  =  Clip(StatusString) & 'FAILED!<13><10>' & Clip(FileErr) & '<13><10>'
                
                ReturnVal  =  False
            Else 
                NEXT(TurboSQLTable)
                IF ERROR() 
                    If FileError() Then
                        FileErr  =  Clip(FileErr) & FileErrorCode() & ']: ' & FileError()
                    Else
                        FileErr  =  Clip(FileErr) & ErrorCode() & ']: ' & Error()
                    End  
                    StatusString  =  Clip(StatusString) & 'FAILED!<13><10>' & Clip(FileErr) & '<13><10>'
                    ReturnVal     =  False
                ELSIF TurboSQL:Variable = 0
                    StatusString  =  Clip(StatusString) & 'FAILED!<13><10>Database does not exist.' & '<13><10>'
                    ReturnVal     =  False
                ELSE
                    
                    StatusString  =  Clip(StatusString) & 'SUCCESS!<13><10>'
                END
                
            End
            Close(TurboSQLTable)
            Close(SysDatabases)
            SysDatabases{Prop:Disconnect}
        End
    Else
        StatusString  =  Clip(StatusString) & 'FAILED!<13><10>'
    End

    IF ReturnVal Then
        StatusString  =  Clip(StatusString) & '<13><10>Test Succeeded!<13><10>'   
        ReturnVal     =  TRUE
    Else
        StatusString  =  Clip(StatusString) & '<13><10>Test FAILED! Please Correct Problems and Try Again!<13><10>'
    End

    If ~Omitted(ErrorOut) Then
        ErrorOut  =  StatusString   
    End

    Return ReturnVal     
        
        
UltimateSQL.SetDriver           PROCEDURE(LONG pDriver = ULSDriver_MSSQL)  

    CODE
    
    SELF._Driver  =  pDriver


! -----------------------------------------------------------------------
!!! <summary>Initializes the Debug class</summary>                   
! -----------------------------------------------------------------------        
UltimateSQL.Init                PROCEDURE(LONG pDriver = ULSDriver_MSSQL,<STRING pProperties>,BYTE pUseRegistryToRetrieveServers = 1)
     
    CODE
    
    SELF.ShowODBCResults  =  FALSE
    
    SELF._Driver  =  pDriver
    IF pDriver = ULSDriver_ODBC     
        IF pProperties
            SELF.NativeClient  =  pProperties 
            
        END
        
    END 
    
    IF pDriver = ULSDriver_SQLite
        SELF.ConnectionString  =  pProperties 
    ELSE
        SELF.ODBCDriver  =  pProperties
        SELF.GetLocalSQLServers()
        
    END
   
    SELF.InitHasRun  =  TRUE


    
!-----------------------------------
UltimateSQL.Kill                PROCEDURE()
!-----------------------------------

    CODE

    RETURN


UltimateSQL.QueryResult         PROCEDURE(STRING pQuery)   !,STRING 

Result                              STRING(4000000)

    CODE
        
    Result  =  ''
    SELF.Query(pQuery,,Result)        
     

    RETURN CLIP(Result)
        

UltimateSQL.SetCatalog          PROCEDURE(STRING pCatalog)   !,STRING 

Result                              STRING(500)

    CODE
    
    IF pCatalog
        SELF.Catalog  =  pCatalog
    END
    

    RETURN       
          
    
UltimateSQL.BeginTransaction                PROCEDURE()

    CODE
    
    SELF.QueryDummy('BEGIN TRANSACTION;')


UltimateSQL.EndTransaction                  PROCEDURE()    

    CODE
    
    SELF.QueryDummy('COMMIT;')

    
UltimateSQL.Query               PROCEDURE(STRING pQuery, <*QUEUE pQ>, <*? pC1>, <*? pC2>, <*? pC3>, <*? pC4>, <*? pC5>, <*? pC6>, <*? pC7>, <*? pC8>, <*? pC9>, <*? pC10>, <*? pC11>, <*? pC12>, <*? pC13>, <*? pC14>, <*? pC15>, <*? pC16>, <*? pC17>,<*? pC18>, <*? pC19>, <*? pC20>, <*? pC21>, <*? pC22>, <*? pC23>, <*? pC24>, <*? pC25>, <*? pC26>, <*? pC27>, <*? pC28>, <*? pC29>, <*? pC30>, <*? pC31>, <*? pC32>, <*? pC33>, <*? pC34>, <*? pC35>, <*? pC36>, <*? pC37>, <*? pC38>, <*? pC39>, <*? pC40>, <*? pC41>, <*? pC42>)  !,BYTE,PROC

result                              long

    CODE        
    
    SELF.Wait(1)  

    SELF.AppendToDriverOptions()
    
    IF LEN(CLIP(pQuery)) < 6   
        SELF.Release(1)
        RETURN ''
    END      
    
    IF SELF.CheckForAndRemoveClarionPrefixes
        pQuery  =  SELF.CheckClarionTablePrefixes(pQuery) 
        
    END
    IF UPPER(pQuery[1:4]) = 'EXEC' OR UPPER(pQuery[1:5]) = '*EXEC' 
        result  =  SELF.QueryODBC(pQuery,pQ,pC1,pC2,pC3,pC4,pC5,pC6,pC7,pC8,pC9,pC10,pC11,pC12,pC13,pC14,pC15,pC16,pC17,pC18,pC19,pC20,pC21,pC22,pC23,pC24,pC25,pC26,pC27,pC28,pC29,pC30,pC31,pC32,pC33,pC34,pC35,pC36,pC37,pC38,pC39,pC40,pC41,pC42)  !,BYTE,PROC
    
    ELSIF (UPPER(pQuery[1:6]) = 'SELECT' OR UPPER(pQuery[1:7]) = '*SELECT' |
            OR UPPER(pQuery[1:4]) = 'WITH' OR UPPER(pQuery[1:5]) = '*WITH'|
            OR UPPER(pQuery[1:4]) = 'RESTORE' OR UPPER(pQuery[1:5]) = '*RESTORE'|
            OR UPPER(pQuery[1:4]) = 'DBCC' OR UPPER(pQuery[1:5]) = '*DBCC'|
            OR UPPER(pQuery[1:7]) = 'DECLARE' OR UPPER(pQuery[1:8]) = '*DECLARE')|
            AND SELF.NativeClient 
        
        IF SELF._Driver  = ULSDriver_SQLite  
            result  =  SELF.QueryDummy(pQuery,pQ,pC1,pC2,pC3,pC4,pC5,pC6,pC7,pC8,pC9,pC10,pC11,pC12,pC13,pC14,pC15,pC16,pC17,pC18,pC19,pC20,pC21,pC22,pC23,pC24,pC25,pC26,pC27,pC28,pC29,pC30,pC31,pC32,pC33,pC34,pC35,pC36,pC37,pC38,pC39,pC40,pC41,pC42)  !,BYTE,PROC
            
        ELSE       
            result  =  SELF.QueryODBC(pQuery,pQ,pC1,pC2,pC3,pC4,pC5,pC6,pC7,pC8,pC9,pC10,pC11,pC12,pC13,pC14,pC15,pC16,pC17,pC18,pC19,pC20,pC21,pC22,pC23,pC24,pC25,pC26,pC27,pC28,pC29,pC30,pC31,pC32,pC33,pC34,pC35,pC36,pC37,pC38,pC39,pC40,pC41,pC42)  !,BYTE,PROC
            
        END
          
    ELSIF ~SELF.NativeClient
        result  =  SELF.QueryDummy(pQuery,pQ,pC1,pC2,pC3,pC4,pC5,pC6,pC7,pC8,pC9,pC10,pC11,pC12,pC13,pC14,pC15,pC16,pC17,pC18,pC19,pC20,pC21,pC22,pC23,pC24,pC25,pC26,pC27,pC28,pC29,pC30,pC31,pC32,pC33,pC34,pC35,pC36,pC37,pC38,pC39,pC40,pC41,pC42)  !,BYTE,PROC
        
    ELSE
        result  =  SELF.QueryDummy(pQuery,pQ,pC1,pC2,pC3,pC4,pC5,pC6,pC7,pC8,pC9,pC10,pC11,pC12,pC13,pC14,pC15,pC16,pC17,pC18,pC19,pC20,pC21,pC22,pC23,pC24,pC25,pC26,pC27,pC28,pC29,pC30,pC31,pC32,pC33,pC34,pC35,pC36,pC37,pC38,pC39,pC40,pC41,pC42)  !,BYTE,PROC
        
    END 
    SELF.Release(1)
        
    RETURN result
     
    
UltimateSQL.QueryToJson                     PROCEDURE (STRING pQuery,BYTE pJsonForOption = us:ForJsonAuto)  !,STRING,VIRTUAL
    
qJson                               QUEUE,PRE(qJson)
JsonString                              STRING(5000)
                                    END
JsonReturnString                                UltimateString
lc                                              LONG


    CODE
     
    SELF.Query(CLIP(pQuery) & ' FOR JSON ' & CHOOSE(pJsonForOption = us:ForJsonAuto,'AUTO','PATH'),qJson,qJson.JsonString)
    LOOP lc = 1 TO RECORDS(qJson)
        GET(qJson,lc)
        JsonReturnString.Append(CLIP(qJson.JsonString))
    END
    
    RETURN JsonReturnString.Get()
        
! -----------------------------------------------------------------------
!!! <summary>Sends Queries to the SQL database</summary>           
!!! <param name="Query">The actual Query to be sent to the SQL database</param>
!!! <param name="Q">A Queue to receive Query results.  This is optional if you are only receiving a single row.</param>        
!!! <param name="C1...C42">Variables belonging to the passed Queue, or stand-alone variables to reeive a single result.</param>        
! -----------------------------------------------------------------------
UltimateSQL.QueryDummy          PROCEDURE(STRING pQuery, <*QUEUE pQ>, <*? pC1>, <*? pC2>, <*? pC3>, <*? pC4>, <*? pC5>, <*? pC6>, <*? pC7>, <*? pC8>, <*? pC9>, <*? pC10>, <*? pC11>, <*? pC12>, <*? pC13>, <*? pC14>, <*? pC15>, <*? pC16>, <*? pC17>,<*? pC18>, <*? pC19>, <*? pC20>, <*? pC21>, <*? pC22>, <*? pC23>, <*? pC24>, <*? pC25>, <*? pC26>, <*? pC27>, <*? pC28>, <*? pC29>, <*? pC30>, <*? pC31>, <*? pC32>, <*? pC33>, <*? pC34>, <*? pC35>, <*? pC36>, <*? pC37>, <*? pC38>, <*? pC39>, <*? pC40>, <*? pC41>, <*? pC42>)  !,BYTE,PROC
            
QueryResults                        FILE,DRIVER('ODBC',DriverOptions),PRE(QueryResults),BINDABLE,THREAD
Record                                  RECORD,PRE()
QueryFields                                 LIKE(QueryFieldsGroup)
                                        END   
                                    END  

QueryResultsMSSQL                   FILE,DRIVER('MSSQL',DriverOptions),PRE(QueryResultsMSSQL),BINDABLE,THREAD
Record                                  RECORD,PRE()
QueryFields                                 LIKE(QueryFieldsGroup)
                                        END   
                                    END

QueryResultsSQLite                  FILE,DRIVER('SQLite',DriverOptions),PRE(QueryResultsSQLite),BINDABLE,THREAD
Record                                  RECORD,PRE()
QueryFields                                 LIKE(QueryFieldsGroup)
                                        END   
                                    END  

QueryResultsSQLiteInMem             FILE,DRIVER('SQLite','/TURBOSQL=True'),OWNER(':memory:'),PRE(QueryResultsSQLiteInMem),BINDABLE,THREAD,CREATE
Record                                  RECORD,PRE()
QueryFields                                 LIKE(QueryFieldsGroup)
                                        END   
                                    END 

QueryFile                           &FILE
QueryGroup                          &GROUP

QueryView                           VIEW(QueryResults)
                                    END

QueryViewMSSQL                      VIEW(QueryResultsMSSQL)  
                                    END

QueryViewSQLite                     VIEW(QueryResultsSQLite)
                                    END  

ExecOK                              BYTE(0)
ResultQ                             &QUEUE
Recs                                ULONG(0)

QString                             CSTRING(LEN(pQuery)+1)  !8192)   ! 8K Limit ???  !MG

NoRetVal                            BYTE(0)         ! NO Return Values (True/False)
BindVars                            BYTE(0)         ! Binded Variables Exist (True/False)
BindVarQ                            QUEUE           ! Binded Variables
No                                      BYTE          ! No
Name                                    STRING(18)    ! Name
                                    END     
UsingExec                           BYTE(0)  

StoredPrefix                        STRING(20)

QueueField                          ANY
FCount                              LONG  
qCount                              LONG

TestField                           ANY
locString                           UltimateString
locQuery                            UltimateString
Prefix                              STRING(30)
lc                                  LONG
TotalQCount                         LONG
      
sQuery                              UltimateString

    CODE 
     
    PUSHBIND()        
    
    ExecOK  =  False  
    SELF.ClearErrors()
    FREE(BindVarQ) 
    SELF.SQLError  =  ''
    BindVars       =  False 
    NoRetVal       =  False 
    StoredPrefix   =  SELF.DebugPrefix
    IF SELF.QueryTableName = ''
        SELF.QueryTableName  =  'dbo.Queries'
    END 
       
    IF UPPER(pQuery[1:4]) = 'EXEC' OR UPPER(pQuery[1:5]) = '*EXEC'
        UsingEXEC  =  TRUE
    END                            
    
    SELF.AppendToDriverOptions()
         
    IF SELF.SQLiteInMemory = TRUE  
        QueryGroup  &=  QueryResultsSQLiteInMem:Record  
        
    ELSE
        IF SELF._Driver  = ULSDriver_ODBC
            QueryFile   &=  QueryResults
            QueryGroup  &=  QueryResults:Record
        
        ELSIF SELF._Driver  = ULSDriver_MSSQL
            QueryFile   &=  QueryResultsMSSQL
            QueryGroup  &=  QueryResultsMSSQL:Record   
        
        ELSIF SELF._Driver  = ULSDriver_SQLite  
            QueryFile   &=  QueryResultsSQLite
            QueryGroup  &=  QueryResultsSQLite:Record
        
        END 
        
    END
    
    QueryFile{PROP:Owner}  =  SELF.ConnectionString
    QueryFile{PROP:Name}   =  SELF.QueryTableName   
    ExecOK                 =  False  
    FREE(BindVarQ) ; BindVars = False ; NoRetVal = False  
    
    IF (OMITTED(pQ) AND OMITTED(pC1) AND OMITTED(pC2)) THEN NoRetVal = True.  ! No Return Values - Possible an UPDATE/DELETE statement
    
    IF NOT OMITTED(pQ) THEN ResultQ &= pQ END  ! If Result Queue Exists - Reference Queue  
                      
    SELF.DebugPrefix  =  '[SQL] ' & CLIP(QueryFile{PROP:Owner})
    sQuery.Assign(SELF.CheckForDebug(pQuery))
    
    IF pQuery = ''
        BEEP !; MESSAGE('Missing Query Statement')
        
    ELSE     
        !~! Parse Query String for EMBEDDED Variables to be BOUND
        IF INSTRING('CALL ',UPPER(sQuery.Get()),1,1) OR INSTRING('EXEC ',UPPER(sQuery.Get()),1,1) ! Check if Stored Procedure is Called  
            QString  =  CLIP(sQuery.Get())
            S# = 0 ;  L# = LEN(CLIP(QString))
            LOOP C# = 1 TO L#
                IF S# AND INLIST(QString[C#],',',' ',')') 
                    BindVarQ.No    =  RECORDS(BindVarQ) + 1
                    BindVarQ.Name  =  QString[(S#+1) : (C#-1)]
                    IF NOT OMITTED(3+BindVarQ.No) ! Bound Variables MUST be at the Beginning of OUTPUT Variables   
                        EXECUTE BindVarQ.No
                            BIND(CLIP(BindVarQ.Name),pC1)
                            BIND(CLIP(BindVarQ.Name),pC2)
                            BIND(CLIP(BindVarQ.Name),pC3)
                            BIND(CLIP(BindVarQ.Name),pC4)
                            BIND(CLIP(BindVarQ.Name),pC5)
                            BIND(CLIP(BindVarQ.Name),pC6)
                            BIND(CLIP(BindVarQ.Name),pC7)
                            BIND(CLIP(BindVarQ.Name),pC8)
                            BIND(CLIP(BindVarQ.Name),pC9)
                            BIND(CLIP(BindVarQ.Name),pC10)
                            BIND(CLIP(BindVarQ.Name),pC11)
                            BIND(CLIP(BindVarQ.Name),pC12)
                            BIND(CLIP(BindVarQ.Name),pC13)
                            BIND(CLIP(BindVarQ.Name),pC14)
                            BIND(CLIP(BindVarQ.Name),pC15)
                            BIND(CLIP(BindVarQ.Name),pC16)
                            BIND(CLIP(BindVarQ.Name),pC17)
                            BIND(CLIP(BindVarQ.Name),pC18)
                            BIND(CLIP(BindVarQ.Name),pC19)
                            BIND(CLIP(BindVarQ.Name),pC20)
                            BIND(CLIP(BindVarQ.Name),pC21)
                            BIND(CLIP(BindVarQ.Name),pC22)
                            BIND(CLIP(BindVarQ.Name),pC23)
                            BIND(CLIP(BindVarQ.Name),pC24)
                            BIND(CLIP(BindVarQ.Name),pC25)
                            BIND(CLIP(BindVarQ.Name),pC26)
                            BIND(CLIP(BindVarQ.Name),pC27)
                            BIND(CLIP(BindVarQ.Name),pC28)
                            BIND(CLIP(BindVarQ.Name),pC29)
                            BIND(CLIP(BindVarQ.Name),pC30)
                            BIND(CLIP(BindVarQ.Name),pC31)
                            BIND(CLIP(BindVarQ.Name),pC32)
                            BIND(CLIP(BindVarQ.Name),pC33)
                            BIND(CLIP(BindVarQ.Name),pC34)
                            BIND(CLIP(BindVarQ.Name),pC35)
                            BIND(CLIP(BindVarQ.Name),pC36)
                            BIND(CLIP(BindVarQ.Name),pC37)
                            BIND(CLIP(BindVarQ.Name),pC38)
                            BIND(CLIP(BindVarQ.Name),pC39)
                            BIND(CLIP(BindVarQ.Name),pC40)
                            BIND(CLIP(BindVarQ.Name),pC41)
                            BIND(CLIP(BindVarQ.Name),pC42)
                        END
                        
                    END

                    ADD(BindVarQ,+BindVarQ.No)
                    IF ERRORCODE()
                        BEEP 
                        SELF.Debug('BindVarQ : ' & ERROR()) 
                        BREAK
                        
                    END  
                    
                    S# = 0  
                    
                END  
                
                IF QString[C#] = '&'                      
                    IF S#
                        BEEP
                        SELF.Debug('Improper Use of BINDED Variables')
                        BREAK 
                        
                    ELSE
                        S# = C#  
                        
                    END 
                    
                END
                
            END

            BindVars  =  RECORDS(BindVarQ)
        END   
        
        IF UsingExec
        ELSE   
            IF SELF.SQLiteInMemory   
                OPEN(QueryResultsSQLiteInMem) 
                CLEAR(QueryResultsSQLiteInMem) 
                BUFFER(QueryResultsSQLiteInMem,100)       
                
            ELSE
                IF ~STATUS(QueryFile)  
                    SELF.AppendToDriverOptions()  
                    OPEN(QueryFile) 
                    CLEAR(QueryFile) 
                
                END     
            
                CLEAR(QueryGroup)
                BUFFER(QueryFile,100)       
                
            END
            
        END 
        
        IF SELF.ErrorCode()
        ELSE
            IF UsingEXEC
                SELF.AppendToDriverOptions()   
                IF SELF._Driver  = ULSDriver_ODBC   
                    QueryResults{PROP:Owner}  =  SELF.ConnectionString
                    QueryResults{PROP:Name}   =  SELF.QueryTableName  
                    OPEN(QueryResults)
                    OPEN(QueryView)
                    BUFFER(QueryView,100)       
        
                ELSIF SELF._Driver  = ULSDriver_MSSQL  
                    QueryViewMSSQL{PROP:Owner}  =  SELF.ConnectionString
                    QueryViewMSSQL{PROP:Name}   =  SELF.QueryTableName 
                    OPEN(QueryResultsMSSQL)
                    OPEN(QueryViewMSSQL)
                    
                ELSIF SELF._Driver  = ULSDriver_SQLite 
                    IF SELF.SQLiteInMemory   
                        OPEN(QueryResultsSQLiteInMem) 
                        BUFFER(QueryResultsSQLiteInMem,100) 
                        
                    ELSE
                        QueryViewSQLite{PROP:Owner}  =  SELF.ConnectionString
                        QueryViewSQLite{PROP:Name}   =  SELF.QueryTableName 
                        OPEN(QueryResultsSQLite) 
                        OPEN(QueryViewSQLite)
                        BUFFER(QueryResultsSQLite,100)  
                        
                    END
                    
                END 
                
            END 
            
            IF SELF.ERRORCODE()
                
            ELSE     
                IF LEN(CLIP(sQuery.Get())) > 7
!                    IF SELF._Driver = ULSDriver_SQLite AND |
!                            UPPER(pQuery[1:6]) <> 'SELECT' AND |
!                            UPPER(pQuery[1:7]) <> '*SELECT' AND |
!                            UPPER(pQuery[1:6]) <> 'PRAGMA' AND |
!                            UPPER(pQuery[1:7]) <> '*PRAGMA'  
!                        locString.Assign('BEGIN TRANSACTION;')  
!                        DO ProcessQuery 
!                        
!                    END   
                    
                    locString.Assign(sQuery.Get())
                    DO ProcessQuery 
                    
!                    IF SELF._Driver = ULSDriver_SQLite AND |
!                            UPPER(pQuery[1:6]) <> 'SELECT' AND |
!                            UPPER(pQuery[1:7]) <> '*SELECT' AND |
!                            UPPER(pQuery[1:6]) <> 'PRAGMA' AND |
!                            UPPER(pQuery[1:7]) <> '*PRAGMA'
!                        locString.Assign('COMMIT;')
!                        DO ProcessQuery 
!                        
!                    END
                    
                END
                
            END
            
        END
        
    END  
    
    IF SELF.SQLiteInMemory
    ELSE
        BUFFER(QueryFile,1)       
    END
    
    IF UsingEXEC
        IF SELF._Driver  = ULSDriver_ODBC
            CLOSE(QueryView)
            CLOSE(QueryResults)  
            
        ELSIF SELF._Driver  = ULSDriver_MSSQL
            CLOSE(QueryViewMSSQL)
            CLOSE(QueryResultsMSSQL)
                                    
        ELSIF SELF._Driver  = ULSDriver_SQLite  
            IF SELF.SQLiteInMemory
                CLOSE(QueryResultsSQLiteInMem)
            ELSE
                CLOSE(QueryViewSQLite)
                CLOSE(QueryResultsSQLite)
                
            END
            
        END
       
    ELSE   
        IF SELF._Driver  = ULSDriver_SQLite AND SELF.SQLiteInMemory
            CLOSE(QueryResultsSQLiteInMem)
            
        ELSE
            IF STATUS(QueryFile)  
                CLOSE(QueryFile) 
                SELF.Error('Closing Queryfile') 
            
            END 
            
        END
        
    END
    
    IF BindVars
        LOOP C# = 1 TO BindVars
            GET(BindVarQ, C#)
            IF BindVarQ.Name THEN UNBIND(CLIP(BindVarQ.Name)).   ! Use pushbind popbind??
            
        END
        
    END
    FREE(BindVarQ) 
    
    IF (SELF.SendToDebug OR SELF.ShowQueryInDebugView OR SELF.AddQueryToClipboard OR SELF.AppendQueryToClipboard) AND SELF.GetError()
        SELF.DebugPrefix  =  '[ERR] '
        SELF.SendToDebug  =  FALSE

        IF ~SELF.AllDebugOff    
            
            IF SELF.ShowQueryInDebugView OR SELF.SendToDebug
                SELF.Debug(SELF.GetError())  
            END
            
            IF SELF.AppendQueryToClipboard
                SETCLIPBOARD(CLIP(CLIPBOARD()) & '<13,10,13,10>' & CLIP(SELF.GetError()))
                
            ELSIF SELF.AddQueryToClipboard
                SETCLIPBOARD(CLIP(SELF.GetError())) 
                
            END
        END
            
    END     
    
    SELF.DebugPrefix  =  StoredPrefix
        
    POPBIND()       
    
    RETURN ExecOK 
    
    
ProcessQuery                    ROUTINE
    
    IF UsingEXEC
        IF SELF._Driver  = ULSDriver_ODBC
            QueryView{PROP:SQL}  =  CLIP(locString.Get())
        
        ELSIF SELF._Driver  = ULSDriver_MSSQL
            QueryViewMSSQL{PROP:SQL}  =  CLIP(locString.Get()) 
                        
        ELSIF SELF._Driver  = ULSDriver_SQLite 
            IF SELF.SQLiteInMemory
                QueryResultsSQLiteInMem{PROP:SQL}  =  CLIP(locString.Get())
            ELSE
                QueryViewSQLite{PROP:SQL}  =  CLIP(locString.Get())
            END
            
        END
                    
    ELSE    
        IF SELF._Driver  = ULSDriver_SQLite AND SELF.SQLiteInMemory
            QueryResultsSQLiteInMem{PROP:SQL}  =  CLIP(locString.Get())
            
        ELSE
            QueryFile{PROP:SQL}  =  CLIP(locString.Get()) 
            
        END
        
    END 
    
    IF ERRORCODE() 
        IF ERRORCODE() <> 33
            SELF.SetError(ERRORCODE(),FILEERROR())
                        
        END
                    
    ELSE
        IF BindVars OR NoRetVal
            ExecOK  =  TRUE   
            
        ELSE
            Recs  =  0 
            
            LOOP 
                IF UsingEXEC   
                    IF SELF._Driver  = ULSDriver_ODBC
                        NEXT(QueryView)
        
                    ELSIF SELF._Driver  = ULSDriver_MSSQL
                        NEXT(QueryViewMSSQL)
                                    
                    ELSIF SELF._Driver  = ULSDriver_SQLite  
                        IF SELF.SQLiteInMemory
                            NEXT(QueryResultsSQLiteInMem) 
                            
                        ELSE
                            NEXT(QueryViewSQLite)
                            
                        END
                        
                    END
                                
                ELSE      
                    IF SELF._Driver  = ULSDriver_SQLite AND SELF.SQLiteInMemory 
                        NEXT(QueryResultsSQLiteInMem)
                        
                    ELSE
                        NEXT(QueryFile)  
                        
                    END
                    
                END     
                            
                Recs +=  1    
!                IF SELF._Driver = ULSDriver_SQLite AND |
!                        UPPER(pQuery[1:6]) <> 'SELECT' AND |
!                        UPPER(pQuery[1:7]) <> '*SELECT' AND |
!                        UPPER(pQuery[1:6]) <> 'PRAGMA' AND |
!                        UPPER(pQuery[1:7]) <> '*PRAGMA'  
!                 
!                    RETURN
!                        
!                END
                IF NOT ERRORCODE()
                    ExecOK  =  True 
                                
                    IF NOT OMITTED(pQ) AND ~UsingExec AND SELF._Driver <> ULSDriver_SQLite  
                        CLEAR(ResultQ)
                        IF OMITTED(pC1)  
                            FCount  =  0
                            LOOP 
                                FCount        +=  1
                                QueueField    &=  WHAT(ResultQ,FCount) 
                                IF QueueField &= NULL THEN BREAK.
                                QueueField     =  WHAT(QueryGroup,FCount+1)
                                            
                            END      
                            
                            ADD(ResultQ)
                            CYCLE
                                        
                        END
                                    
                    END
  
                    IF NOT OMITTED(pC1)  THEN pC1  = WHAT(QueryGroup,2).
                    IF NOT OMITTED(pC2)  THEN pC2  = WHAT(QueryGroup,3).
                    IF NOT OMITTED(pC3)  THEN pC3  = WHAT(QueryGroup,4).
                    IF NOT OMITTED(pC4)  THEN pC4  = WHAT(QueryGroup,5).
                    IF NOT OMITTED(pC5)  THEN pC5  = WHAT(QueryGroup,6).
                    IF NOT OMITTED(pC6)  THEN pC6  = WHAT(QueryGroup,7).
                    IF NOT OMITTED(pC7)  THEN pC7  = WHAT(QueryGroup,8).
                    IF NOT OMITTED(pC8)  THEN pC8  = WHAT(QueryGroup,9).
                    IF NOT OMITTED(pC9)  THEN pC9  = WHAT(QueryGroup,10).
                    IF NOT OMITTED(pC10) THEN pC10 = WHAT(QueryGroup,11).
                    IF NOT OMITTED(pC11) THEN pC11 = WHAT(QueryGroup,12).
                    IF NOT OMITTED(pC12) THEN pC12 = WHAT(QueryGroup,13).
                    IF NOT OMITTED(pC13) THEN pC13 = WHAT(QueryGroup,14).
                    IF NOT OMITTED(pC14) THEN pC14 = WHAT(QueryGroup,15).
                    IF NOT OMITTED(pC15) THEN pC15 = WHAT(QueryGroup,16).
                    IF NOT OMITTED(pC16) THEN pC16 = WHAT(QueryGroup,17).
                    IF NOT OMITTED(pC17) THEN pC17 = WHAT(QueryGroup,18).
                    IF NOT OMITTED(pC18) THEN pC18 = WHAT(QueryGroup,19).
                    IF NOT OMITTED(pC19) THEN pC19 = WHAT(QueryGroup,20).
                    IF NOT OMITTED(pC20) THEN pC20 = WHAT(QueryGroup,21).
                    IF NOT OMITTED(pC21) THEN pC21 = WHAT(QueryGroup,22).
                    IF NOT OMITTED(pC22) THEN pC22 = WHAT(QueryGroup,23).
                    IF NOT OMITTED(pC23) THEN pC23 = WHAT(QueryGroup,24).
                    IF NOT OMITTED(pC24) THEN pC24 = WHAT(QueryGroup,25).
                    IF NOT OMITTED(pC25) THEN pC25 = WHAT(QueryGroup,26).
                    IF NOT OMITTED(pC26) THEN pC26 = WHAT(QueryGroup,27).
                    IF NOT OMITTED(pC27) THEN pC27 = WHAT(QueryGroup,28).
                    IF NOT OMITTED(pC28) THEN pC28 = WHAT(QueryGroup,29).
                    IF NOT OMITTED(pC29) THEN pC29 = WHAT(QueryGroup,30).
                    IF NOT OMITTED(pC30) THEN pC30 = WHAT(QueryGroup,31).
                    IF NOT OMITTED(pC31) THEN pC31 = WHAT(QueryGroup,32).
                    IF NOT OMITTED(pC32) THEN pC32 = WHAT(QueryGroup,33).
                    IF NOT OMITTED(pC33) THEN pC33 = WHAT(QueryGroup,34).
                    IF NOT OMITTED(pC34) THEN pC34 = WHAT(QueryGroup,35).
                    IF NOT OMITTED(pC35) THEN pC35 = WHAT(QueryGroup,36).
                    IF NOT OMITTED(pC36) THEN pC36 = WHAT(QueryGroup,37).
                    IF NOT OMITTED(pC37) THEN pC37 = WHAT(QueryGroup,38).
                    IF NOT OMITTED(pC38) THEN pC38 = WHAT(QueryGroup,39).
                    IF NOT OMITTED(pC39) THEN pC39 = WHAT(QueryGroup,40).
                    IF NOT OMITTED(pC40) THEN pC40 = WHAT(QueryGroup,41).
                    IF NOT OMITTED(pC41) THEN pC41 = WHAT(QueryGroup,42).
                    IF NOT OMITTED(pC42) THEN pC42 = WHAT(QueryGroup,43).  
                    IF NOT OMITTED(pq) ! Result Queue  
                        ADD(ResultQ)
                    END 
                                
                ELSE   
                    IF OMITTED(pq) ! NO Result Queue
                        IF ERRORCODE() <> 33
                            IF ERRORCODE() = 90
                                CASE FILEERRORCODE()
                                OF ''          
                                    ! Ignore NO File Error
                                OF '24000'     
                                    ! Ignore Cursor State Error - 
                                    ! Statement Executes BUT Error Returned
                                OF 'S1010'     
                                    ! Ignore Function Sequencing Error - 
                                    ! Statement Executes BUT Error Returned
                                ELSE 
                                    IF FILEERROR()
                                        SELF.SetError(ERRORCODE(),FILEERROR())
                                    END
                                                
!!                                                SELF.Debug('Error : ' & ERROR() &|
!!                                                        ' [' & ERRORCODE() & ']|' & 'File Error : ' &|
!!                                                        FILEERROR() & ' [' & FILEERRORCODE() &|
!!                                                        ']||' & pQuery,'NEXT VIEW')
                                END
                                            
                            ELSE
                                IF FILEERROR()
                                    SELF.SetError(ERRORCODE(),FILEERROR())
                                END
                                            
!!                                            SELF.Debug('Error : ' & ERROR() & ' [' &|
!!                                                    ERRORCODE() & ']|' &|
!!                                                    'File : ' & ERRORFILE() & '||' & pQuery,'NEXT VIEW')
                            END
                        ELSE
!!                                        IF FILEERROR()
!!                                            SELF.SetError(ERRORCODE(),FILEERROR())
!!                                        END
                                        
                        END
                                    
                    ELSE
                        IF FILEERROR()
                            IF ERRORCODE() <> 33
                                SELF.SetError(ERRORCODE(),FILEERROR())
                            END
                                        
                        END
                                    
                        BREAK
                                    
                    END
                                
                END
                
                IF OMITTED(pq) THEN BREAK. ! NO Result Queue   
                            
            END
            
        END
                    
    END 
    
        
UltimateSQL.QueryODBC           PROCEDURE(STRING pQuery, <*QUEUE pQ>, <*? pC1>, <*? pC2>, <*? pC3>, <*? pC4>, <*? pC5>, <*? pC6>, <*? pC7>, <*? pC8>, <*? pC9>, <*? pC10>, <*? pC11>, <*? pC12>, <*? pC13>, <*? pC14>, <*? pC15>, <*? pC16>, <*? pC17>,<*? pC18>, <*? pC19>, <*? pC20>, <*? pC21>, <*? pC22>, <*? pC23>, <*? pC24>, <*? pC25>, <*? pC26>, <*? pC27>, <*? pC28>, <*? pC29>, <*? pC30>, <*? pC31>, <*? pC32>, <*? pC33>, <*? pC34>, <*? pC35>, <*? pC36>, <*? pC37>, <*? pC38>, <*? pC39>, <*? pC40>, <*? pC41>, <*? pC42>)  !,BYTE,PROC

ExecOK                              BYTE(0)
lCnt                                long
lRowCnt                             long
lColCnt                             long  
ReturnValue                         LONG                        

TheOwnerString                      UltimateSQLString
        
DirectODBC                          UltimateSQLDirect    
ViewResults                         UltimateSQLResultsViewClass

StoredPrefix                        STRING(20)
ErrCnt                              LONG
       
QueryFile                           &FILE
QueryGroup                          &GROUP

QueueField                          ANY
FCount                              LONG
      
qNamesAndPosition                   QUEUE,PRE(qNamesAndPosition)
Name                                    STRING(100)
Position                                LONG
                                    END     

QueryODBC                           FILE,DRIVER('ODBC','/TURBOSQL=True /LOGONSCREEN=FALSE'),PRE(QueryODBC),BINDABLE,THREAD
Record                                  RECORD,PRE()
AnyField                                    STRING(200)
                                        END   
                                    END 
 
sQuery                              UltimateString

    CODE     
    
    IF ~pQuery
        RETURN Level:Fatal
    END            
             
    SELF.AppendToDriverOptions()
    
    SELF.SQLError  =  '' 
    SELF.ClearErrors()
    
    StoredPrefix  =  SELF.DebugPrefix
    
    
    sQuery.Assign(SELF.CheckForDebug(pQuery))    
    
    IF SELF.QueryTableName = ''
        SELF.QueryTableName  =  'dbo.Queries'
    END                       
    
    IF NOT OMITTED(pQ) AND OMITTED(pC1) 
        FCount  =  0
        LOOP 
            FCount      +=  1
            QueueField  &=  WHAT(pQ,FCount) 
            IF QueueField &= NULL THEN BREAK.
              
            qNamesAndPosition.Name      =  UPPER(WHO(pQ,FCount)  )
            qNamesAndPosition.Position  =  FCount
            ADD(qNamesAndPosition)
            
        END 
        
    END 
    
    ExecOK       =  False                           
    ReturnValue  =  DirectODBC.OpenConnection(SELF.FullConnectionString)   
    IF DirectODBC.ExecDirect(CLIP(sQuery.Get())) = UltimateSQL_Success  
        LOOP lCnt = 1 To Records(DirectODBC.ResultSets)
            IF DirectODBC.AssignCurrentResultSet(lCnt) = Level:Benign
                LOOP lRowCnt = 1 To Records(DirectODBC.CurrentResultSet)  
                    IF NOT OMITTED(pQ)    
                        CLEAR(pQ)
                        IF OMITTED(pC1)  
                            FCount  =  0  
                            LOOP FCount = 1 TO DirectODBC.SQLColCount  
                                qNamesAndPosition.Name  =  UPPER(DirectODBC.GetColumnName(FCount)) 
                                GET(qNamesAndPosition,qNamesAndPosition.Name)
                                IF ~ERROR()
                                    QueueField  &=  WHAT(pQ,qNamesAndPosition.Position) 
                                    IF QueueField &= NULL THEN BREAK.
                                    QueueField  =  DirectODBC.GetColumnValue(lRowCnt,FCount) 
                                END 
                                
                            END 
                                    
                            ADD(pQ)
                            CYCLE
                        END
                        
                    END    
                    
                    IF NOT OMITTED(PC1);pC1 = DirectODBC.GetColumnValue(lRowCnt,1) END
                    IF NOT OMITTED(PC2);pC2 = DirectODBC.GetColumnValue(lRowCnt,2) END
                    IF NOT OMITTED(PC3);pC3 = DirectODBC.GetColumnValue(lRowCnt,3) END
                    IF NOT OMITTED(PC4);pC4 = DirectODBC.GetColumnValue(lRowCnt,4) END
                    IF NOT OMITTED(PC5);pC5 = DirectODBC.GetColumnValue(lRowCnt,5) END
                    IF NOT OMITTED(PC6);pC6 = DirectODBC.GetColumnValue(lRowCnt,6) END
                    IF NOT OMITTED(PC7);pC7 = DirectODBC.GetColumnValue(lRowCnt,7) END
                    IF NOT OMITTED(PC8);pC8 = DirectODBC.GetColumnValue(lRowCnt,8) END
                    IF NOT OMITTED(PC9);pC9 = DirectODBC.GetColumnValue(lRowCnt,9) END
                    IF NOT OMITTED(PC10);pC10 = DirectODBC.GetColumnValue(lRowCnt,10) END
                    IF NOT OMITTED(PC11);pC11 = DirectODBC.GetColumnValue(lRowCnt,11) END
                    IF NOT OMITTED(PC12);pC12 = DirectODBC.GetColumnValue(lRowCnt,12) END
                    IF NOT OMITTED(PC13);pC13 = DirectODBC.GetColumnValue(lRowCnt,13) END
                    IF NOT OMITTED(PC14);pC14 = DirectODBC.GetColumnValue(lRowCnt,14) END
                    IF NOT OMITTED(PC15);pC15 = DirectODBC.GetColumnValue(lRowCnt,15) END
                    IF NOT OMITTED(PC16);pC16 = DirectODBC.GetColumnValue(lRowCnt,16) END
                    IF NOT OMITTED(PC17);pC17 = DirectODBC.GetColumnValue(lRowCnt,17) END
                    IF NOT OMITTED(PC18);pC18 = DirectODBC.GetColumnValue(lRowCnt,18) END
                    IF NOT OMITTED(PC19);pC19 = DirectODBC.GetColumnValue(lRowCnt,19) END
                    IF NOT OMITTED(PC20);pC20 = DirectODBC.GetColumnValue(lRowCnt,20) END
                    IF NOT OMITTED(PC21);pC21 = DirectODBC.GetColumnValue(lRowCnt,21) END
                    IF NOT OMITTED(PC22);pC22 = DirectODBC.GetColumnValue(lRowCnt,22) END
                    IF NOT OMITTED(PC23);pC23 = DirectODBC.GetColumnValue(lRowCnt,23) END
                    IF NOT OMITTED(PC24);pC24 = DirectODBC.GetColumnValue(lRowCnt,24) END
                    IF NOT OMITTED(PC25);pC25 = DirectODBC.GetColumnValue(lRowCnt,25) END
                    IF NOT OMITTED(PC26);pC26 = DirectODBC.GetColumnValue(lRowCnt,26) END
                    IF NOT OMITTED(PC27);pC27 = DirectODBC.GetColumnValue(lRowCnt,27) END
                    IF NOT OMITTED(PC28);pC28 = DirectODBC.GetColumnValue(lRowCnt,28) END
                    IF NOT OMITTED(PC29);pC29 = DirectODBC.GetColumnValue(lRowCnt,29) END
                    IF NOT OMITTED(PC30);pC30 = DirectODBC.GetColumnValue(lRowCnt,30) END
                    IF NOT OMITTED(PC31);pC31 = DirectODBC.GetColumnValue(lRowCnt,31) END
                    IF NOT OMITTED(PC32);pC32 = DirectODBC.GetColumnValue(lRowCnt,32) END
                    IF NOT OMITTED(PC33);pC33 = DirectODBC.GetColumnValue(lRowCnt,33) END
                    IF NOT OMITTED(PC34);pC34 = DirectODBC.GetColumnValue(lRowCnt,34) END
                    IF NOT OMITTED(PC35);pC35 = DirectODBC.GetColumnValue(lRowCnt,35) END
                    IF NOT OMITTED(PC36);pC36 = DirectODBC.GetColumnValue(lRowCnt,36) END
                    IF NOT OMITTED(PC37);pC37 = DirectODBC.GetColumnValue(lRowCnt,37) END
                    IF NOT OMITTED(PC38);pC38 = DirectODBC.GetColumnValue(lRowCnt,38) END
                    IF NOT OMITTED(PC39);pC39 = DirectODBC.GetColumnValue(lRowCnt,39) END
                    IF NOT OMITTED(PC40);pC40 = DirectODBC.GetColumnValue(lRowCnt,40) END
                    IF NOT OMITTED(PC41);pC41 = DirectODBC.GetColumnValue(lRowCnt,41) END
                    IF NOT OMITTED(PC42);pC42 = DirectODBC.GetColumnValue(lRowCnt,42) END 
                    
                    IF NOT OMITTED(pq)
                        ADD(pQ)
                        
                    ELSE 
                        BREAK ! NO Result Queue 
                        
                    END 
                    
                END     
                    
            END 
                
        END 
            
    END     
        
    IF SELF.ShowODBCResults   
        ViewResults.DisplayResults(DirectODBC,'ODBC Results',SELF.ConnectionString)  
        SELF.ShowODBCResults  =  FALSE
            
    END 
        
    DirectODBC.CloseConnection()
 
    FREE(SELF.ErrorMsg)
    SELF.SQLError  =  ''
    LOOP ErrCnt = 1 TO RECORDS(DirectODBC.ErrorMsg)
        GET(DirectODBC.ErrorMsg,ErrCnt)        
        IF INSTRING('Errorcode 33',DirectODBC.ErrorMsg.ErrorMsg,1,1);CYCLE.
        IF INSTRING('Errorcode 2',DirectODBC.ErrorMsg.ErrorMsg,1,1);CYCLE.
        SELF.ErrorMsg.ErrorMsg    =  DirectODBC.ErrorMsg.ErrorMsg
        SELF.ErrorMsg.ErrorState  =  DirectODBC.ErrorMsg.ErrorState
        ADD(SELF.ErrorMsg)  
        SELF.SetError(DirectODBC.ErrorMsg.ErrorState,DirectODBC.ErrorMsg.ErrorMsg)
        
    END
    
    IF ~RECORDS(SELF.ErrorMsg)
        ExecOK  =  TRUE
       
    END
       
    IF SELF.SendToDebug
        IF ~RECORDS(SELF.ErrorMsg)
        ELSE
            SELF.DebugPrefix  =  '[ERR] '
            IF ~SELF.AllDebugOff    
            
                IF SELF.ShowQueryInDebugView OR SELF.SendToDebug
                    LOOP lCnt = 1 TO RECORDS(SELF.ErrorMsg) 
                        GET(SELF.ErrorMsg,lCnt)
                        SELF.Debug(SELF.ErrorMsg.ErrorMsg)  
                    END
                
                END
            END
            
        END
            
    END    
    
    SELF.DebugPrefix  =  StoredPrefix
    
    RETURN ExecOK  
   
    
UltimateSQL.CreateTemporaryProcedure    PROCEDURE (STRING pQuery) 
          
QueryTempMSSQL                              FILE,DRIVER('MSSQL',DriverOptions),PRE(QueryTempMSSQL),BINDABLE,THREAD
Record                                          RECORD,PRE()
TempField                                           STRING(1)
                                                END   
                                            END

    CODE
         
    SELF.Wait(72)
    SELF.QueryDummy(pQuery)
    SELF.Release(72)    
    RETURN
    
    
UltimateSQL.DropTemporaryProcedure      PROCEDURE (STRING pTemporaryProcedureName) 

    CODE
    
    SELF.Wait(72)
    SELF.QueryDummy('IF OBJECT_ID(<39>tempdb..' & CLIP(pTemporaryProcedureName) & '<39>) IS NOT NULL <13,10>' & |
            'BEGIN <13,10>' & |
            '    DROP PROC ' & CLIP(pTemporaryProcedureName) & ' <13,10>' & |
            'END')
    SELF.Release(72)    
    RETURN
!
! RRS 01/23/19 Added for large tables.
!
UltimateSQL.QueryCT             PROCEDURE (STRING pQuery, <*QUEUE pQ>, <*ctFieldQ pFields>)

ExecOK                              BYTE(0)
lCnt                                long
lRowCnt                             long
lColCnt                             long  
ReturnValue                         LONG                        

TheOwnerString                      UltimateSQLString
        
DirectODBC                          UltimateSQLDirect    
ViewResults                         UltimateSQLResultsViewClass

StoredPrefix                        STRING(20)
ErrCnt                              LONG

    CODE  
    
    IF ~pQuery
        RETURN Level:Fatal
    END
    
    SELF.Wait(4)
    
    SELF.SQLError  =  ''
    StoredPrefix   =  SELF.DebugPrefix  
    pQuery         =  SELF.CheckForDebug(pQuery)
    
    IF SELF.QueryTableName = ''
        SELF.QueryTableName  =  'dbo.Queries'
        
    END 
    
    ExecOK       =  False
    ReturnValue  =  DirectODBC.OpenConnection(SELF.FullConnectionString)  

    If DirectODBC.ExecDirect(CLIP(pQuery)) = UltimateSQL_Success  
        Loop lCnt = 1 To Records(DirectODBC.ResultSets)
            if DirectODBC.AssignCurrentResultSet(lCnt) = Level:Benign
                Loop lRowCnt = 1 To Records(DirectODBC.CurrentResultSet)
                    LOOP lColCnt = 1 To RECORDS(pFields.Q)
                        pFields.SetOrigField( lColCnt, DirectODBC.GetColumnValue(lRowCnt, lColCnt))  
                        
                    END  
                    
                    IF NOT OMITTED(pq)
                        ADD(pQ)
                        
                    ELSE
                        BREAK
                        
                    END 
                    
                END
                
            END
            
        END
        
        IF SELF.QueryResultsShowInPopUp                                 
            ViewResults.DisplayResults(DirectODBC,'ODBC Results',SELF.ConnectionString)  
        END
        
    END 
    
    DirectODBC.CloseConnection()
    FREE(SELF.ErrorMsg)
    
    LOOP ErrCnt = 1 TO RECORDS(DirectODBC.ErrorMsg)
        GET(DirectODBC.ErrorMsg,ErrCnt)
        SELF.ErrorMsg.ErrorMsg    =  DirectODBC.ErrorMsg.ErrorMsg
        SELF.ErrorMsg.ErrorState  =  DirectODBC.ErrorMsg.ErrorState
        ADD(SELF.ErrorMsg)
        
    END 
    
    IF SELF.SendToDebug
        IF ~RECORDS(SELF.ErrorMsg)
        ELSE
            IF ~SELF.AllDebugOff    
                IF SELF.ShowQueryInDebugView OR SELF.SendToDebug
                    LOOP lCnt = 1 TO RECORDS(SELF.ErrorMsg) 
                        GET(SELF.ErrorMsg,lCnt)
                        
                    END 
                    
                END 
                
            END 
            
        END
        
    END
    
    SELF.Release(4)
    
    RETURN ExecOK
!
! RRS 01/23/19 End.

    
!!! <summary>Adds a Column to a Table</summary>
!!! <param name=""Column"">Column Name</param>
!!! <param name=""Length"">Length of Field (if applicable
!!! <param name=""Options"">Any additional options - ex. DEFAULT 0  ex. NOT NULL</param>
!!! <param name=""Table"">Table Name</param>
!!! <param name=""Type"">Type of Field to add - ex. VCHAR</param>
!!! <remarks>Save function as AddColumn</remarks>          
UltimateSQL.CreateColumn        PROCEDURE(STRING pTable,STRING pColumn,STRING pType,<STRING pLength>,<STRING pOptions>)

    CODE
    
    RETURN SELF.AddColumn(pTable,pColumn,pType,pLength,pOptions)
    

! -----------------------------------------------------------------------
!!! <summary>Adds a Column to a Table</summary>           
!!! <param name="Table">Table Name</param>
!!! <param name="Column">Column Name</param>
!!! <param name="Type">Type of Field to add - ex. VCHAR</param>
!!! <param name="Length">Length of Field (if applicable) - ex. 20  ex. 15,2</param>
!!! <param name="Options">Any additional options - ex. DEFAULT 0  ex. NOT NULL</param>
! -----------------------------------------------------------------------        
UltimateSQL.AddColumn           PROCEDURE(STRING pTable,STRING pColumn,STRING pType,<STRING pLength>,<STRING pOptions>)

Result                              LONG
VarDefinition                       UltimateSQLString
  
USQuery                             UltimateSQLString


    CODE

    VarDefinition.Assign(pColumn)
    VarDefinition.Append(' ' & pType)
    IF pLength
        VarDefinition.Append('(' & CLIP(pLength) & ')') 
        
    END  
    
    IF pOptions
        VarDefinition.Append(' ' & CLIP(pOptions)) 
        
    END   
    
    IF SELF._Driver  = ULSDriver_SQLite
        USQuery.Append('*ALTER TABLE ' & CLIP(pTable) & ' ADD ' & CLIP(VarDefinition.Get()))
        
    ELSE  
        USQuery.Append('IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ' & SELF.Quote(pTable) & |
                ' AND COLUMN_NAME = ' & SELF.Quote(pColumn) & ') ALTER TABLE ' & CLIP(pTable) & ' ADD ' & CLIP(VarDefinition.Get()))
        
    END
    
    RETURN SELF.QUERY(USQuery.Get()) 
        

! -----------------------------------------------------------------------
!!! <summary>Alters a Column in a Table</summary>           
!!! <param name="Table">Table Name</param>
!!! <param name="Column">Column Name</param>
!!! <param name="Type">Type of Field to add - ex. VCHAR</param>
!!! <param name="Length">Length of Field (if applicable) - ex. 20  ex. 15,2</param>
!!! <param name="Options">Any additional options - ex. DEFAULT 0  ex. NOT NULL</param>
! -----------------------------------------------------------------------        
UltimateSQL.AlterColumn         PROCEDURE(STRING pTable,STRING pColumn,STRING pType,<STRING pLength>,<STRING pOptions>)

Result                              LONG
VarDefinition                       UltimateSQLString


    CODE
    IF ~SELF.ColumnExists(pTable,pColumn)
        RETURN SELF.AddColumn(pTable,pColumn,pTable,pLength,pOptions)
    ELSE
        VarDefinition.Assign(pColumn)
        VarDefinition.Append(' ' & pType)
        IF pLength
            VarDefinition.Append('(' & CLIP(pLength) & ')')
        END
        IF pOptions
            VarDefinition.Append(' ' & CLIP(pOptions))
        END
    
        SELF.DropDependencies(pTable,pColumn)
!        SELF.DisableTableConstraints(pTable)
!        SELF.EnableTableConstraints(pTable)
        
        RETURN   SELF.QUERY('ALTER TABLE ' & CLIP(pTable) & ' ALTER COLUMN ' & CLIP(VarDefinition.Get()))        
        
    END
                         
        
! -----------------------------------------------------------------------
!!! <summary>Checks to see if a Column exists</summary>           
!!! <param name="Table">Table Name</param>
!!! <param name="Column">Column Name</param>        
!!! <param name="Catalog">Catalog (Database) Name, optional</param>
!!! <param name="Schema">Schema Name, optional, default is dbo</param>
!!! <remarks>Returns TRUE if exists, FALSE if does not exist</remarks>        
! ----------------------------------------------------------------------- 
UltimateSQL.ColumnExists        PROCEDURE(STRING pTable,STRING pColumn,<STRING pCatalog>,<STRING pSchema>)  !,LONG

Schema                              STRING(200)
Catalog                             STRING(200)
Result                              LONG

    CODE
    
    Catalog  =  pCatalog
    Schema   =  pSchema
    IF Catalog = ''
        Catalog  =  SELF.Catalog
    END             
    
    IF Schema = ''
        Schema  =  'dbo'
    END
    
    RETURN SELF.QUERYResult('SELECT count(*) FROM INFORMATION_SCHEMA.COLUMNS' & |  
            ' WHERE TABLE_CATALOG = ' & CLIP(SELF.Quote(Catalog)) & |
            ' AND TABLE_SCHEMA = ' & CLIP(SELF.Quote(Schema)) & | 
            ' AND TABLE_NAME = ' & CLIP(SELF.Quote(pTable)) & |
            ' AND COLUMN_NAME = ' & CLIP(SELF.Quote(pColumn)))    
        
          
        
        
! -----------------------------------------------------------------------
!!! <summary>Alters a Column in a Table</summary>           
!!! <param name="Table">Table Name</param>
!!! <param name="Column">Column Name</param>
!!! <remarks>Dropping a Column can fail if Constraints, Indexes, and Statistics exist.
!!! This method drops all of those first, so remove column will succeed.</remarks>
! -----------------------------------------------------------------------         
UltimateSQL.DisableTableConstraints     PROCEDURE(STRING pTable) !,PROC    
! -----------------------------------------------------------------------         

    CODE
    
    SELF.Query('ALTER TABLE ' & SELF.Quote(pTable) & ' NOCHECK CONSTRAINT ALL') 

    RETURN
    

! -----------------------------------------------------------------------
!!! <summary>Alters a Column in a Table</summary>           
!!! <param name="Table">Table Name</param>
!!! <param name="Column">Column Name</param>
!!! <remarks>Dropping a Column can fail if Constraints, Indexes, and Statistics exist.
!!! This method drops all of those first, so remove column will succeed.</remarks>
! -----------------------------------------------------------------------         
UltimateSQL.EnableTableConstraints      PROCEDURE(STRING pTable)  !,PROC      
! -----------------------------------------------------------------------   

    CODE
    
    SELF.Query('ALTER TABLE ' & SELF.Quote(pTable) & ' CHECK CONSTRAINT ALL') 

    RETURN

    
!!! <summary>Alters a Column in a Table</summary>           
!!! <param name="Table">Table Name</param>
!!! <param name="Column">Column Name</param>
!!! <remarks>Dropping a Column can fail if Constraints, Indexes, and Statistics exist.
!!! This method drops all of those first, so remove column will succeed.</remarks>
! -----------------------------------------------------------------------         
UltimateSQL.DropColumn          PROCEDURE(STRING pTable,STRING pColumn)  !,LONG,PROC
Result                              LONG(0)
VarDefinition                       UltimateSQLString


    CODE
        
    IF SELF.DropDependencies(pTable,pColumn)
    END
    Result  =  SELF.QUERY('ALTER TABLE ' & CLIP(pTable) & ' DROP COLUMN ' & CLIP(pColumn))   
        
    RETURN Result

        
! -----------------------------------------------------------------------
!!! <summary>Alters a Column in a Table</summary>           
!!! <param name="Table">Table Name</param>
!!! <param name="Column">Column Name</param>
!!! <remarks>Dropping a Column can fail if Constraints, Indexes, and Statistics exist.
!!! This method drops all of those first, so remove column will succeed.</remarks>
! -----------------------------------------------------------------------         
UltimateSQL.DropDependencies    PROCEDURE(STRING pTable,STRING pColumn)  !,LONG,PROC    

Result                              LONG(0)
Scripts                             UltimateSQLScripts
ScriptToRun                         UltimateSQLString

    CODE              
        
    ScriptToRun.Assign(Scripts.DropAllDependencies())
    ScriptToRun.Replace('[PASSEDTABLE]',CLIP(SELF.RemoveIllegalCharacters(pTable)))
    ScriptToRun.Replace('[PASSEDCOLUMN]',CLIP(SELF.RemoveIllegalCharacters(pColumn)))   
        
    RETURN  SELF.Query(ScriptToRun.Get())    
        
        
! -----------------------------------------------------------------------
!!! <summary>Drops a Table from the Database</summary>           
!!! <param name="Table">Table Name</param>
! ----------------------------------------------------------------------- 
UltimateSQL.DropTable           PROCEDURE(FILE pFile)  !,LONG  !,PROC 

Result                              LONG

    CODE
        
    RETURN SELF.QUERY('IF OBJECT_ID(' & SELF.Quote(pFile{PROP:Name}) & ') IS NOT NULL DROP TABLE ' & CLIP(SELF.Bracket(NAME(pFile)))   )

    
! -----------------------------------------------------------------------
!!! <summary>Drops a Table from the Database</summary>           
!!! <param name="Table">Table Name</param>
! ----------------------------------------------------------------------- 
UltimateSQL.DropTable           PROCEDURE(STRING pFile)  !,LONG  !,PROC 

Result                              LONG

    CODE
        
    RETURN SELF.QUERY('IF OBJECT_ID(' & SELF.Quote(pFile) & ') IS NOT NULL DROP TABLE ' & SELF.Bracket(pFile)) 
    
             
! -----------------------------------------------------------------------
!!! <summary>Drops a View from the Database</summary>           
!!! <param name="View">View Name</param>
! ----------------------------------------------------------------------- 
UltimateSQL.DropView            PROCEDURE(STRING pView)  !,LONG  !,PROC 

Result                              LONG

    CODE
        
    RETURN SELF.QUERY('*IF OBJECT_ID(' & SELF.Quote(pView) & ') IS NOT NULL DROP VIEW ' & SELF.Bracket(pView))   
    
    
! -----------------------------------------------------------------------
!!! <summary>Drops a Procedure from the Database</summary>           
!!! <param name="View">Procedure Name</param>
! ----------------------------------------------------------------------- 
UltimateSQL.DropProcedure       PROCEDURE(STRING pProcedure)  !,LONG  !,PROC 

Result                              LONG

    CODE
        
    RETURN SELF.QUERY('*IF OBJECT_ID(' & UPPER(SELF.Quote(pProcedure)) & ') IS NOT NULL DROP PROCEDURE ' & SELF.Bracket(pProcedure))  
     
    
! -----------------------------------------------------------------------
!!! <summary>Drops a Routine from the Database</summary>           
!!! <param name="View">Routine Name</param>
! ----------------------------------------------------------------------- 
UltimateSQL.DropFunction        PROCEDURE(STRING pFunction)  !,LONG  !,PROC 

Result                              LONG

    CODE
        
    RETURN SELF.QUERY('IF OBJECT_ID(' & UPPER(SELF.Quote(pFunction)) & ') IS NOT NULL DROP FUNCTION ' & SELF.Bracket(pFunction))  
    
    
! -----------------------------------------------------------------------
!!! <summary>Drops a Routine from the Database</summary>           
!!! <param name="View">Routine Name</param>
! ----------------------------------------------------------------------- 
UltimateSQL.DropTrigger         PROCEDURE(STRING pTrigger)  !,LONG  !,PROC 

Result                              LONG

    CODE
        
    RETURN SELF.QUERY('*IF OBJECT_ID(' & UPPER(SELF.Quote(pTrigger)) & ') IS NOT NULL DROP TRIGGER ' & SELF.Bracket(pTrigger))  
    
     
! -----------------------------------------------------------------------
!!! <summary>Truncates a Table</summary>           
!!! <param name="Table">Table Name</param>
! ----------------------------------------------------------------------- 
UltimateSQL.Empty               PROCEDURE(FILE pFile)  !,LONG  !,PROC 

Result                              LONG

    CODE
        
    RETURN SELF.Truncate(pFile)  
        
        
! -----------------------------------------------------------------------
!!! <summary>Get record by key</summary>
!!! <param name="Key">Reference to key in table</param>
!!! <param name="Select">Optional additional SQL selection criteria</param>
!!! <remarks>
!!! Returns True on success, False on failure. Use SQL to get a 
!!! record using the indicated key. Assumes that the key
!!! values have been filled into the record buffer. Example:
!!!     GLcls:ClassificationID = loc:ClassificationID
!!!     IF SQL.Get(GLcls:Key_ClassificationID) THEN ...
!!! SQL Generated:
!!!     SELECT * FROM tablename WHERE f1 = v1 [AND f2 = v2...]
!!!         [AND (pSelect)]
!!! </remarks>
! -----------------------------------------------------------------------        
UltimateSQL.Get                 PROCEDURE(*KEY pKey, <STRING pSelect>)  !, BYTE, PROC

oUltimateDB                         UltimateDB

    CODE

    RETURN  oUltimateDB.SQLFetch(pKey,pSelect)
  
    
! -----------------------------------------------------------------------
UltimateSQL.Get                 PROCEDURE(*FILE pFile, *KEY pKey, <STRING pSelect>)  !, BYTE, PROC

oUltimateDB                         UltimateDB

    CODE

    RETURN  oUltimateDB.SQLFetch(pKey,pSelect)    
    

! -----------------------------------------------------------------------    
UltimateSQL.GetFieldList        PROCEDURE(*FILE pTbl)  !,String

oUltimateDB                         UltimateDB
    
    CODE
    
    RETURN oUltimateDB.GetFieldList(pTbl)
                 
    
! -----------------------------------------------------------------------
UltimateSQL.CreateTableList     PROCEDURE(*FILE pTbl,<STRING pViewName>)   

    CODE
    

! -----------------------------------------------------------------------
UltimateSQL.Records             PROCEDURE(*FILE pFile,<STRING pFilter>) !,LONG

LocalFilter                         UltimateString

    CODE
 
    IF pFilter
        LocalFilter.Assign(' Where ' & pFilter)
    END
        
    RETURN SELF.QueryResult('Select COUNT(*) FROM ' & NAME(pFile) & CLIP(LocalFilter.Get()))

        
! -----------------------------------------------------------------------
!!! <summary>Fetch records by key</summary>
!!! <param name="Key">Reference to key in table</param>
!!! <param name="Select">Optional additional selection criteria</param>
!!! <param name="Reverse">T=Reverse order of key, [F]=Use key order</param>
!!! <remarks>Uses SQL to fetch a list of child records using the indicated key.
!!! All selection is done through the optional select argument.
!!! Example:
!!!     SQL.Set(INtac:Key_APInvoiceIDActivityDate)
!!!     LOOP
!!!       NEXT(INtac)
!!!       IF ERRORCODE() THEN BREAK.
!!!         ...
!!!     END
!!! SQL Generated:
!!!     SELECT * FROM tablename [WHERE pSelect] ORDER BY f1, ...</remarks>
! -----------------------------------------------------------------------        
UltimateSQL.Set                 PROCEDURE(*KEY pKey, <STRING pSelect>, BYTE pReverse=False)

oUltimateDB                         UltimateDB

    CODE

    oUltimateDB.SQLFetchChildren(pKey,pSelect)     
        
    RETURN
    

! -----------------------------------------------------------------------        
UltimateSQL.Set                 PROCEDURE(*KEY pKeyIgnored,*KEY pKey, <STRING pSelect>, BYTE pReverse=False)

oUltimateDB                         UltimateDB

    CODE

    oUltimateDB.SQLFetchChildren(pKey,pSelect)   
        
    RETURN      


! -----------------------------------------------------------------------    
UltimateSQL.RenameColumn        PROCEDURE(STRING pTable,STRING pOldColumn,STRING pNewColumn)          !,LONG,PROC  

    CODE
    
    RETURN SELF.Query('sp_RENAME ' & SELF.Quote(CLIP(pTable) & '.' & CLIP(pOldColumn)) & ', ' & SELF.Quote(pNewColumn) & ', <39>COLUMN<39>')
    
    
        
! -----------------------------------------------------------------------
!!! <summary>Truncates a Table</summary>           
!!! <param name="Table">Table Name</param>
! ----------------------------------------------------------------------- 
UltimateSQL.Truncate            PROCEDURE(FILE pFile)  !,LONG  !,PROC 

Result                              LONG

    CODE
    
    IF SELF._Driver = ULSDriver_SQLite
        RETURN SELF.QUERY('DELETE FROM ' & CLIP(NAME(pFile)))   
        
    ELSE
        RETURN SELF.QUERY('TRUNCATE TABLE ' & CLIP(NAME(pFile)))   
        
    END
    
        
! -----------------------------------------------------------------------
UltimateSQL.Truncate            PROCEDURE(STRING pFileName)  !,LONG,PROC 

Result                              LONG

    CODE
      
    IF SELF._Driver = ULSDriver_SQLite
        RETURN SELF.QUERY('DELETE FROM ' & CLIP(pFileName))   
        
    ELSE
        RETURN SELF.QUERY('TRUNCATE TABLE ' & CLIP(pFileName))  
        
    END
    
! -----------------------------------------------------------------------
!!! <summary>Checks to see if a Database exists</summary>           
!!! <param name="Database">Database Name</param>
!!! <remarks>Returns TRUE if exists, FALSE if does not exist</remarks>        
! -----------------------------------------------------------------------         
UltimateSQL.DatabaseExists      PROCEDURE(STRING pDatabase)   !,BYTE

    CODE              
        
    RETURN SELF.QueryResult('Select count(*) from master.dbo.sysdatabases where name = ' & SELF.Quote(pDatabase))  
        
        
! -----------------------------------------------------------------------
!!! <summary>Creates a new Database. Connection string information must be passed in.</summary>
!!! <param name=""String"">Server</param>
!!! <param name=""String"">User Name</param>
!!! <param name=""String"">Password</param>
!!! <param name=""String"">Database Name</param>
!!! <param name=""Byte""><Optional> Byte Trusted</param>
!!! <returns>Returns status of Database creation (see Class for return valuesUltimateSQL.            
UltimateSQL.CreateDatabase      PROCEDURE(String Server, String USR, String PWD, String Database, <Byte Trusted>) ! Declare Procedure  

SQLStr                              String(1024)
FileErr                             String(20)
DBCount                             Long
ReturnVal                           Long

SysObjects                          FILE,DRIVER('MSSQL'),OWNER(DatabaseCheckOwnerString), Name('SysObjects')
Record                                  RECORD,PRE()
                                        END
                                    END
SysDatabases                        FILE,DRIVER('MSSQL'),OWNER(DatabaseCheckOwnerString), Name('SysDatabases')
Record                                  RECORD,PRE()
name                                        CSTRING(129)
dbid                                        SHORT
sid                                         STRING(85)
mode                                        SHORT
status                                      LONG
status2                                     LONG
crdate                                      STRING(8)
crdate_GROUP                                GROUP,OVER(crdate)
crdate_DATE                                     DATE
crdate_TIME                                     TIME
                                            END
reserved                                    STRING(8)
reserved_GROUP                              GROUP,OVER(reserved)
reserved_DATE                                   DATE
reserved_TIME                                   TIME
                                            END
category                                    LONG
cmptlevel                                   BYTE
filename                                    CSTRING(261)
version                                     LONG
                                        END
                                    END

    CODE                                                     ! Begin processed code
    !Return Values

    !-11    =   Tables Exist must be blank database
    !-10    =   Do not have permission to complete operation
    !-9     =   Login Failed
    !-8     =   Server Not Found
    !-7     =   Error Creating Database Object
    !-6     =   Error Reteiving Record
    !-5     =   SQL Query Error
    !-4     =   Could Not Open
    !-3     =   User Name Invalid
    !-2     =   Database Name Invalid
    !-1     =   Server Name Invalid
    !0      =   DB Created Successfully
    !1      =   DB Already Exists

    !Select Count(*) From sysobjects Where xtype='U'
    If Omitted(5) Then
        Trusted  =  False
    End

    ReturnVal  =  0

    If Len(Clip(Server)) <= 0 Then
        ReturnVal  =  -1
    End

    If Len(Clip(Database)) <= 0 Then
        ReturnVal  =  -2
    End
    If (Len(Clip(USR)) <= 0) AND (~Trusted) Then
        ReturnVal  =  -3
    End

    If ~ReturnVal And Trusted Then
        Send(SysDatabases, '/TRUSTEDCONNECTION = TRUE')
    Else
        Send(SysDatabases, '/TRUSTEDCONNECTION = FALSE')
    End

    If ~ReturnVal Then
        DatabaseCheckOwnerString        =  Clip(Server) & ',master,' & Clip(USR) & ',' & Clip(PWD)
        SysDatabases{Prop:Logonscreen}  =  False
        
        Open(SysDatabases)

        If Error() Then
            If FileError() Then
                FileErr  =  FileErrorCode()
            Else
                FileErr  =  ErrorCode()
            End
            ReturnVal  =  -4
        End

        If ~ReturnVal Then
            !Query for Database in SQL Server
            SQLStr                  =  'SELECT COUNT(*) FROM SYSDATABASES WHERE name=N''' & Database & ''';'
            SysDatabases{Prop:SQL}  =  SQLStr
            If Error() Then
                If FileError() Then
                    FileErr  =  FileErrorCode()
                Else
                    FileErr  =  ErrorCode()
                End
                ReturnVal  =  -5
            End

            If ~ReturnVal Then
                !Check Query
                Next(SysDatabases)
                If Error() Then
                    If FileError() Then
                        FileErr  =  FileErrorCode()
                    Else
                        FileErr  =  ErrorCode()
                    End
                    ReturnVal  =  -6
                Else
                    DBCount  =  SysDatabases.Name
                    If DBCount > 0 Then
                        !Database already exists do not create
                        ReturnVal  =  1
                    End
                End

                If ~ReturnVal Then
                    !Create Database
                    SQLStr                  =  'CREATE DATABASE [' & Clip(Database) & ']'
                    SysDatabases{Prop:SQL}  =  SQLStr
                    If Error() Then
                        If FileError() Then
                            FileErr  =  FileErrorCode()
                            !Message(FileErrorCode() & ' = ' & FileError())
                        Else
                            FileErr  =  ErrorCode()
                        End
                        ReturnVal  =  -7
                    End
                End
            End
            
            Close(SysDatabases)
            SysDatabases{Prop:Disconnect}
        End
    End

    !Set Specific Error Codes
    IF Len(Clip(FileErr)) > 0 Then
        Case Clip(FileErr)
        Of  '08001'
            !Server Not Found
            ReturnVal  =  -8
        Of  '28000'
            !Login Failed
            ReturnVal  =  -9
        Of  '37000'
            !Do not have permission to complete operation
            ReturnVal  =  -10
        End
    End
    Return ReturnVal



!----------------------------------------
UltimateSQL.ObjectExists        PROCEDURE(STRING pObjectName,<STRING pSchema>) !,LONG
!----------------------------------------
Schema                              STRING(200)
ScriptName                          STRING(200)

    CODE     
    
    Schema  =  pSchema
    
    IF Schema = ''
        Schema  =  'dbo'
    END
    
    ScriptName  =  SELF.PrepStatement(Schema,,TRUE) & '.' & SELF.PrepStatement(pObjectName,,TRUE)    
    

    RETURN   SELF.QueryResult('*SELECT COUNT(*) FROM sys.objects' & |  
            ' WHERE object_id = object_id(N' & SELF.PrepStatement(ScriptName,TRUE) & |
            ') AND type <> <39>U<39>')+0
    
    
! -----------------------------------------------------------------------
!!! <summary>Checks to see if a Table exists</summary>           
!!! <param name="Table">Table Name</param>
!!! <param name="Catalog">Catalog (Database) Name, optional</param>
!!! <param name="Schema">Schema Name, optional, default is dbo</param>
!!! <remarks>Returns TRUE if exists, FALSE if does not exist</remarks>        
! -----------------------------------------------------------------------         
UltimateSQL.TableExists         PROCEDURE(STRING pTable,<STRING pCatalog>,<STRING pSchema>)  !,BYTE
        
Schema                              STRING(200)
Catalog                             STRING(200)
qMethod                             BYTE
Result                              BYTE
TableName                           STRING(200)

    CODE     
    
    Catalog  =  pCatalog
    Schema   =  pSchema
    IF Catalog = ''
        Catalog  =  SELF.Database
    END             
    IF Schema = ''
        Schema  =  'dbo'
    END
    
    TableName  =  SELF.StripSchema(Schema,pTable)    

    qMethod           =  SELF.QueryMethod
    SELF.QueryMethod  =  QueryMethodODBC
        
    Result  =  SELF.QUERYResult('*SELECT count(*) FROM INFORMATION_SCHEMA.TABLES' & |  
            ' WHERE TABLE_CATALOG = ' & SELF.Quote(Catalog) & |
            ' AND TABLE_SCHEMA = ' & SELF.Quote(Schema) & | 
            ' AND TABLE_NAME = ' & SELF.Quote(TableName))+0
    
    SELF.QueryMethod  =  qMethod
    
    RETURN   Result
    
        
UltimateSQL.StripSchema         PROCEDURE(STRING pSchema,STRING pTable) !,STRING

us                                  UltimateString

    CODE
    
    us.Assign(UPPER(pTable))
    us.Replace(UPPER(CLIP(pSchema)) & '.','')   
    
    RETURN us.Get()
    
    
    
    
    
    
UltimateSQL.TableExists         PROCEDURE(FILE pFile,<STRING pCatalog>,<STRING pSchema>)  !,BYTE     

    CODE
        
    RETURN SELF.TableExists(NAME(pFile),pCatalog,pSchema)
        
! -----------------------------------------------------------------------
!!! <summary>Drops a Database</summary>           
!!! <param name="Table">Database name</param>
!!! <remarks>Well this can certainly have some ramifications!
!!! Be careful here.</remarks>
! ----------------------------------------------------------------------- 
UltimateSQL.DropDatabase        PROCEDURE(STRING pDatabase)   !,LONG,PROC         

Result                              LONG

    CODE
    SELF.QUERY('ALTER DATABASE ' & CLIP(pDatabase) & ' SET SINGLE_USER WITH ROLLBACK IMMEDIATE')
    RETURN SELF.QUERY('DROP DATABASE ' & CLIP(pDatabase))   
        
        
! -----------------------------------------------------------------------
!!! <summary>Reads a script from a file in to memory and executes it.</summary>           
!!! <param name="FileName">Name of the file (with full path) containing the script.</param>
! -----------------------------------------------------------------------
UltimateSQL.ExecuteScriptFromFile       PROCEDURE(STRING pFileName)   !,BYTE

FetchScript                                 UltimateSQLString                        
QueryStatement                              UltimateSQLString


    CODE
        
    QueryStatement.Assign(CLIP(FetchScript.ReadFile(pFileName))) 
         
    RETURN SELF.ProcessScript(QueryStatement.Get())  
        
    
! -----------------------------------------------------------------------
!!! <summary>Executes an SQL Script from a BLOB</summary>           
!!! <param name="BLOB">A BLOB field</param>
!!! <remarks>Useful for storing scripts in an encrypted file such as a TPS file.</remarks>
! -----------------------------------------------------------------------        
UltimateSQL.ExecuteScriptFromBlob       PROCEDURE(*BLOB pBlob)  !,BYTE,PROC  

QueryStatement                              UltimateSQLString    
BlobSize                                    LONG


    CODE
    
    BlobSize  =  pBlob{PROP:Size}
    
    IF BLOBSIZE = 0
        RETURN Level:Cancel
    END
    
    QueryStatement.Assign(pBlob[0:BlobSize])  
    RETURN SELF.ProcessScript(QueryStatement.Get())
    
    
! -----------------------------------------------------------------------
!!! <summary>Executes an SQL Script from a BLOB</summary>           
!!! <param name="BLOB">A BLOB field</param>
!!! <remarks>Useful for storing scripts in an encrypted file such as a TPS file.</remarks> 
! -----------------------------------------------------------------------        
UltimateSQL.ExecuteScriptsFromQUEUE     PROCEDURE(*QUEUE pQueue)
! -----------------------------------------------------------------------        
    
lc                                          LONG
QueryGroup                                  &GROUP
CurrentField                                ANY
Status                                      Any

locString                                   UltimateString

    CODE

    locString.Assign('')  
    
    IF RECORDS(pQueue)
        IF SELF._Driver = ULSDriver_SQLite
            SELF.Query('BEGIN TRANSACTION')
        END
    
        LOOP lc = 1 TO RECORDS(pQueue)
            GET(pQueue,lc)
            QueryGroup  &=  pQueue 
        
            CurrentField  &=  WHAT(QueryGroup,1)    
        
            IF SELF._Driver = ULSDriver_SQLite
                SELF.Query(CLIP(CurrentField))
            
            ELSE
                locString.Append(CLIP(CurrentField) & '<13,10>')
        
                IF locString.Length(1) > 20000  
                    SELF.Query(locString.Get())
                    locString.Assign('')
            
                END 
            
            END
            
        END
     
        IF SELF._Driver = ULSDriver_SQLite
            SELF.Query('COMMIT')
        
        ELSE
            IF locString.Length(1) > 0
                SELF.Query(locString.Get())
                locString.Assign('')
            
            END
        
        END
        
    END
    
    
UltimateSQL.AddStatement        PROCEDURE(STRING pStatement)    

    CODE
    
    SELF.AddStatement(-1,pStatement)


UltimateSQL.AddStatement        PROCEDURE(LONG pID,STRING pStatement)    

    CODE
    
    SELF.Wait(22)
    SELF.qStatements.ID         =  pID
    SELF.qStatements.Statement  =  pStatement
    ADD(SELF.qStatements)
    SELF.Release(22)
                                 
    
UltimateSQL.ExecuteStatements   PROCEDURE(LONG pStatementID=0)
         
lc                                  LONG
lcc                                 LONG
ErrorLine                           LONG
StatementGUID                       STRING(36)
TotalQCount                         LONG(0)    
ErrorCount                          LONG(0)
TotalQOffset                        LONG(0)
locString                           UltimateString
locQuery                            UltimateString
                                MAP
ProcessThroughQueue                 PROCEDURE()
ProcessBatch                        PROCEDURE(),LONG,PROC 
                                END

    CODE 
                   
    SELF.Wait(23) 
    IF RECORDS(SELF.qStatements)    
        SELF.StatementID    =  pStatementID 
        SELF.StatementGUID  =  SELF.QueryResult('SELECT NEWID()')  
        StatementGUID       =  SELF.StatementGUID
        TotalQCount         =  0     
        IF SELF.DatabaseExists('##DB_Errors')
            SELF.Query('Delete ##DB_Errors Where StatementID = ' & SELF.StatementID)
        ELSE
            SELF.CreateDBErrorsTable(SELF.StatementGUID) 
        END
        
        ProcessThroughQueue()
       
        !Clear all records for this batch
        LOOP lc = RECORDS(SELF.qStatements) TO 1 BY -1
            GET(SELF.qStatements,lc)
            IF SELF.qStatements.ID = SELF.StatementID
                DELETE(SELF.qStatements)
            END
            
        END
        CLEAR(SELF.qStatements)
            
    END
    
    SELF.Release(23)
     
    RETURN StatementGUID
 
    
ProcessThroughQueue             PROCEDURE()
       

Prefix                              STRING(30)
qCount                              LONG
EndOfQ                              LONG
TotalRecords                        LONG

    CODE
    
    LOOP  
        IF qCount = -1;BREAK.
        
        locString.Assign('')  
        locQuery.Assign('') 
        Prefix        =  ''
        qCount        =  0 
        TotalQCount   =  0   
        EndOfQ        =  FALSE
        TotalRecords  =  RECORDS(SELF.qStatements)
        
        LOOP  
            qCount       +=  1
            TotalQCount  +=  1     
            IF qCount > TotalRecords;BREAK.
            
            GET(SELF.qStatements,qCount + 1)
            IF ERROR()
                EndOfQ  =  TRUE
            END
            
            GET(SELF.qStatements,qCount) 
            IF ERROR()
                qCount  =  -1
                BREAK
            END
            
            IF SELF.qStatements.Statement = '';CYCLE.       
            
            IF SELF.qStatements.ID = SELF.StatementID OR SELF.qStatements.ID = -1
                IF SELF.qStatements.Statement[1:1] = '*' OR SELF.ShowQueryInDebugView OR SELF.SendToDebug
                    IF SELF.qStatements.Statement[1:1] = '*'
                        SELF.qStatements.Statement  =  SUB(SELF.qStatements.Statement,2,128000) 
                    END
                    
                END                                         
                        
                IF SELF._Driver = ULSDriver_SQLite 
                    SELF.Query(SELF.qStatements.Statement)
                    locString.Assign('')  
                            
                ELSE       
                    locQuery.Append(CLIP(SELF.qStatements.Statement) & ';<13,10>') 
                    IF Prefix = ''
                        Prefix  =  ' + CHAR(13) + CHAR(10) + '
                    END
                    IF locQuery.Length(1) > 999995000 OR EndOfQ  
                        IF ProcessBatch() = 2  
                            qCount  =  -1
                        ELSE
                            TotalQOffset  +=  TotalQCount
                            LOOP lc = qCount TO 1 BY -1
                                GET(SELF.qStatements,lc)
                                IF SELF.qStatements.ID = SELF.StatementID OR SELF.qStatements.ID = -1
                                    DELETE(SELF.qStatements)
                                END
            
                            END
                            
                        END
                        
                        BREAK  
                        
                    END 
                            
                END
                            
            END
                        
        END    
        
    END     
    
        
ProcessBatch                    PROCEDURE()   
   
Result                              LONG(0)  

    CODE
    IF SELF._Driver = ULSDriver_SQLite
        SELF.Query(locQuery.Get())
        
    ELSE    
        SELF.QueryDummy(locQuery.Get())
        IF SELF.GetError()                                                                                                                                       
            SELF.Query('INSERT INTO ##DB_Errors VALUES (' & SELF.Quote(SELF.StatementGUID) & |
                    ',' & SELF.StatementID & ',' & SELF.Quote(SELF.GetError()) & ',0,' & TotalQCount & ',' & SELF.Quote(locQuery.Get()) & ',GETDATE()); <13,10>')
        END
        
!        locString.Append('DECLARE @SQL NVARCHAR(MAX); <13,10>')
!        locString.Append('BEGIN TRY <13,10>BEGIN TRANSACTION <13,10>')  
!        locString.Append(locQuery.Get() & ' <13,10>')
!        locString.Append('COMMIT TRANSACTION <13,10>')  
!        locString.Append('END TRY <13,10>')  
!        locString.Append('BEGIN CATCH <13,10>')  
!        locString.Append('IF @@TRANCOUNT > 0 <13,10>')  
!        locString.Append('ROLLBACK TRAN <13,10>') 
!        locString.Append('INSERT INTO ##DB_Errors VALUES (' & SELF.Quote(SELF.StatementGUID) & |
!                ',' & SELF.StatementID & ',ERROR_MESSAGE(),ERROR_LINE(),' & TotalQCount & ',<39><39>,GETDATE()); <13,10>') 
!        locString.Append('END CATCH <13,10>')
!        SELF.QueryDummy(locString.Get())
        
        RETURN 2
        
        locString.Append('DECLARE @SQL NVARCHAR(MAX); <13,10>')
        locString.Append('BEGIN TRY <13,10>BEGIN TRANSACTION <13,10>')  
        locString.Append('SET @SQL = ' & locQuery.Get())
        locString.Append('EXECUTE SP_EXECUTESQL @SQL;')
        locString.Append('COMMIT TRANSACTION <13,10>')  
        locString.Append('END TRY <13,10>')  
        locString.Append('BEGIN CATCH <13,10>')  
        locString.Append('IF @@TRANCOUNT > 0 <13,10>')  
        locString.Append('ROLLBACK TRAN <13,10>') 
        locString.Append('INSERT INTO ##DB_Errors VALUES (' & SELF.Quote(SELF.StatementGUID) & |
                ',' & SELF.StatementID & ',ERROR_MESSAGE(),ERROR_LINE(),' & TotalQCount & ',<39><39>,GETDATE()); <13,10>') 
        locString.Append('END CATCH <13,10>')  
        SELF.QueryDummy(locString.Get())
       
        IF SELF.QueryResult('*Select COUNT(*) From ##DB_Errors Where GUID = ' & SELF.Quote(SELF.StatementGUID) & ' AND Statement = <39><39>')+0  
            ErrorCount += 1 
            ErrorLine  =  SELF.QueryResult('Select Line From ##DB_Errors Where GUID = ' & SELF.Quote(SELF.StatementGUID) & ' AND Statement = <39><39>')
            GET(SELF.qStatements,ErrorLine)  
            SELF.Query('Update ##DB_Errors Set Statement = ' & SELF.Quote(SELF.qStatements.Statement) & ',BatchLine = ' & (ErrorLine + TotalQOffset + ErrorCount) & |
                    ' Where GUID = ' & SELF.Quote(SELF.StatementGUID) & ' AND Line = ' & ErrorLine & ' AND Statement = <39><39>' ) 
            DELETE(SELF.qStatements)                          
            Result = 2 
            
        END
        
    END
    
    RETURN  Result
    
    
UltimateSQL.DropDBErrorsTable   PROCEDURE(STRING pBatchGUID)  

    CODE
       
    SELF.DropTable('##DB_Errors') 
    
    
UltimateSQL.RemoveBatchErrors               PROCEDURE(STRING pBatchGUID)     
  
    CODE
    
    SELF.Query('Delete ##DB_Errors Where GUID = ' & SELF.Quote(pBatchGUID))

    
UltimateSQL.CreateDBErrorsTable              PROCEDURE(STRING pBatchGUID)
    
    CODE
    
    SELF.Query('CREATE TABLE ##DB_Errors <13,10>' & |
            '(GUID  CHAR(36), <13,10>' & |
            'StatementID INT, <13,10>' & |
            'Message VARCHAR(MAX), <13,10>' & |
            'Line INT, <13,10>' & |
            'BatchLine INT, <13,10>' & |
            'Statement VARCHAR(MAX),' & |
            'DateTime DATETIME)')
    
    
UltimateSQL.ProcessScript       PROCEDURE(STRING pScript) !,BYTE,PROC

QueryStatement                      UltimateSQLString    
ScriptToExecute                     UltimateSQLString
Count                               LONG
QueryRecords                        LONG

    CODE               
    
    QueryStatement.Assign(pScript)
    QueryStatement.Split()
    QueryRecords  =  QueryStatement.Records()
    LOOP Count = 1 TO QueryRecords
        IF  CLIP(LEFT(QueryStatement.GetLine(Count))) = 'GO'
            IF ScriptToExecute.Length()
                SELF.Query(ScriptToExecute.Get()) 
            END
            ScriptToExecute.Assign()    
            CYCLE
        END 
!!            IF Count = QueryRecords AND ScriptToExecute.Contains('GO');BREAK.
        ScriptToExecute.Append(QueryStatement.GetLine(Count) & '<13,10>')
    END
    IF ScriptToExecute.Length(TRUE)
        SELF.Query(ScriptToExecute.Get()) 
            
    END
    ScriptToExecute.Assign()        

    RETURN TRUE


! -----------------------------------------------------------------------
!!! <summary>Inserts an Extended Property with a Value into an SQL object</summary>           
!!! <param name="ObjectName">The Object the Extended Property will be added to.  The Object must be fully-formed.  See notes below.</param>
!!! <param name="PropertyName">The Extended Property name.</param>
!!! <param name="PropertyValue">The Value for the Extended Property.</param>
!!! <remarks>The Object must be formatted using the following sequence: schema.object1.object2
!!! To add a Property to your Database, The ObjectName should be blank ''
!!! To add a Property to a Table, the ObjectName should be schema.tablename (ex. dbo.customer)
!!! To add a Property to a Column, the ObjectName should be schema.tablename.column (ex. dbo.customer.firstname)
!!! TO add a Property to a Stored Procedure, the ObjectName should be schema.StoredProcedureName (ex. dbo.MyStoredProcedure)</remarks>
! ----------------------------------------------------------------------- 
UltimateSQL.ExtendedProperty_Insert     PROCEDURE(STRING pObjectName,STRING pPropertyName,STRING pPropertyValue) !, LONG, PROC  

Result                                      STRING(400)
ScriptToRun                                 UltimateSQLString
Scripts                                     UltimateSQLScripts

    CODE        
        
    ScriptToRun.Assign(Scripts.InsertExtendedProperty())
    ScriptToRun.Replace('[OBJECTNAME]',CLIP(SELF.RemoveIllegalCharacters(pObjectName)))
    ScriptToRun.Replace('[PROPERTYNAME]',CLIP(SELF.RemoveIllegalCharacters(pPropertyName)))
    ScriptToRun.Replace('[PROPERTYVALUE]',CLIP(SELF.RemoveIllegalCharacters(pPropertyValue)))   
    RETURN  SELF.Query(ScriptToRun.Get())
        
          
! -----------------------------------------------------------------------
!!! <summary>Updates an Extended Property with a Value into an SQL object</summary>           
!!! <param name="ObjectName">The Object that the Extended Property will be updated from.  The Object must be fully-formed.  See notes below.</param>
!!! <param name="PropertyName">The Extended Property name.</param>
!!! <param name="PropertyValue">The Value for the Extended Property.</param>
!!! <remarks>The Object must be formatted using the following sequence: schema.object1.object2
!!! To add a Property to your Database, The ObjectName should be blank ''
!!! To add a Property to a Table, the ObjectName should be schema.tablename (ex. dbo.customer)
!!! To add a Property to a Column, the ObjectName should be schema.tablename.column (ex. dbo.customer.firstname)
!!! TO add a Property to a Stored Procedure, the ObjectName should be schema.StoredProcedureName (ex. dbo.MyStoredProcedure)</remarks>
! -----------------------------------------------------------------------         
UltimateSQL.ExtendedProperty_Update     PROCEDURE(STRING pObjectName,STRING pPropertyName,STRING pPropertyValue) !, LONG, PROC

Result                                      STRING(400)
ScriptToRun                                 UltimateSQLString
Scripts                                     UltimateSQLScripts

    CODE        
        
    IF SELF.ExtendedProperty_Exists(pObjectName,pPropertyName)
        ScriptToRun.Assign(Scripts.UpdateExtendedProperty())
        ScriptToRun.Replace('[OBJECTNAME]',CLIP(SELF.RemoveIllegalCharacters(pObjectName)))
        ScriptToRun.Replace('[PROPERTYNAME]',CLIP(SELF.RemoveIllegalCharacters(pPropertyName)))
        ScriptToRun.Replace('[PROPERTYVALUE]',CLIP(SELF.RemoveIllegalCharacters(pPropertyValue)))
        RETURN  SELF.Query('*' & ScriptToRun.Get())  
    ELSE
        RETURN SELF.ExtendedProperty_Insert(pObjectName,pPropertyName,pPropertyValue)
    END
        
        
! -----------------------------------------------------------------------
!!! <summary>Deletes an Extended Property from an SQL object</summary>           
!!! <param name="ObjectName">The Object the Extended Property will be deleted from.  The Object must be fully-formed.  See notes below.</param>
!!! <param name="PropertyName">The Extended Property name.</param>
!!! <remarks>The Object must be formatted using the following sequence: schema.object1.object2
!!! To delete a Property from your Database, The ObjectName should be blank ''
!!! To delete a Property from a Table, the ObjectName should be schema.tablename (ex. dbo.customer)
!!! To delete a Property from a Column, the ObjectName should be schema.tablename.column (ex. dbo.customer.firstname)
!!! TO delete a Property from a Stored Procedure, the ObjectName should be schema.StoredProcedureName (ex. dbo.MyStoredProcedure)</remarks>
! -----------------------------------------------------------------------         
UltimateSQL.ExtendedProperty_Delete     PROCEDURE(STRING pObjectName,STRING pPropertyName) !, LONG, PROC  

Result                                      STRING(400)
ScriptToRun                                 UltimateSQLString
Scripts                                     UltimateSQLScripts

    CODE       
        
    IF SELF.ExtendedProperty_Exists(pObjectName,pPropertyName)
        ScriptToRun.Assign(Scripts.UpdateExtendedProperty())        
        ScriptToRun.Replace('[OBJECTNAME]',CLIP(SELF.RemoveIllegalCharacters(pObjectName)))
        ScriptToRun.Replace('[PROPERTYNAME]',CLIP(SELF.RemoveIllegalCharacters(pPropertyName)))
        RETURN  SELF.Query(ScriptToRun.Get()) 
    ELSE
        RETURN ''
    END
        
        
! -----------------------------------------------------------------------
!!! <summary>Gets the Value of an Extended Property with a from an SQL object</summary>           
!!! <param name="ObjectName">The Object the Extended Property will be retrieved from.  The Object must be fully-formed.  See notes below.</param>
!!! <param name="PropertyName">The Extended Property name.</param>
!!! <remarks>The Object must be formatted using the following sequence: schema.object1.object2
!!! To get a Property of your Database, The ObjectName should be blank ''
!!! To get a Property of a Table, the ObjectName should be schema.tablename (ex. dbo.customer)
!!! To get a Property of a Column, the ObjectName should be schema.tablename.column (ex. dbo.customer.firstname)
!!! To get a Property of a Stored Procedure, the ObjectName should be schema.StoredProcedureName (ex. dbo.MyStoredProcedure)
!!! The Method returns the actual Value of the Property</remarks>
! -----------------------------------------------------------------------        
UltimateSQL.ExtendedProperty_GetValue   PROCEDURE(STRING pObjectName,STRING pPropertyName) !, STRING

Result                                      STRING(400)
ScriptToRun                                 UltimateSQLString
Scripts                                     UltimateSQLScripts            
QueryMethod                                 BYTE(0)

    CODE           
        
    ScriptToRun.Assign(Scripts.GetExtendedProperty())        
    ScriptToRun.Replace('[OBJECTNAME]',CLIP(SELF.RemoveIllegalCharacters(pObjectName)))
    ScriptToRun.Replace('[PROPERTYNAME]',CLIP(SELF.RemoveIllegalCharacters(pPropertyName)))
!!    ScriptToRun.Replace('[SELECTOPERATION]','Value')

    QueryMethod       =  SELF.QueryMethod
    SELF.QueryMethod  =  QueryMethodODBC
    Result            =  SELF.QueryResult(ScriptToRun.Get())
    SELF.QueryMethod  =  QueryMethod    
    RETURN Result
        
! -----------------------------------------------------------------------
!!! <summary>Determines whether the the Property exists in an SQL object</summary>           
!!! <param name="ObjectName">The Object where the Extended Property will be checked.  The Object must be fully-formed.  See notes below.</param>
!!! <param name="PropertyName">The Extended Property name.</param>
!!! <remarks>The Object must be formatted using the following sequence: schema.object1.object2
!!! To determine if a Property exists in your Database, The ObjectName should be blank ''
!!! To determine if a Property exists in a Table, the ObjectName should be schema.tablename (ex. dbo.customer)
!!! To determine if a Property exists in a Column, the ObjectName should be schema.tablename.column (ex. dbo.customer.firstname)
!!! To determine if a Property exists in a Stored Procedure, the ObjectName should be schema.StoredProcedureName (ex. dbo.MyStoredProcedure)
!!! The Method returns TRUE or FALSE</remarks>
! -----------------------------------------------------------------------        
UltimateSQL.ExtendedProperty_Exists     PROCEDURE(STRING pObjectName,STRING pPropertyName) !, BYTE

Result                                      LONG
ScriptToRun                                 UltimateSQLString
Scripts                                     UltimateSQLScripts            

    CODE           
        
    ScriptToRun.Assign(Scripts.GetExtendedProperty())        
    ScriptToRun.Replace('[OBJECTNAME]',CLIP(SELF.RemoveIllegalCharacters(pObjectName)))
    ScriptToRun.Replace('[PROPERTYNAME]',CLIP(SELF.RemoveIllegalCharacters(pPropertyName)))
    ScriptToRun.Replace('[SELECTOPERATION]','COUNT(*)')
    SELF.Query(ScriptToRun.Get(),,Result)
        
    RETURN Result 

 
!----------------------------------------
UltimateSQL.GetAllExtendedProperties    PROCEDURE(STRING pObjectName)
!----------------------------------------

    CODE
    
    FREE(SELF.qExtendedProperties) 
    
    SELF.Query(' SELECT <13,10>' & |
            ' CASE WHEN ob.parent_object_id>0 <13,10>' & |
            ' THEN OBJECT_SCHEMA_NAME(ob.parent_object_id) <13,10>' & |
            ' + <39>.<39>+OBJECT_NAME(ob.parent_object_id)+<39>.<39>+ob.name <13,10>' & |
            ' ELSE OBJECT_SCHEMA_NAME(ob.object_id)+<39>.<39>+ob.name END <13,10>' & |
            ' + CASE WHEN ep.minor_id>0 THEN <39>.<39>+col.name ELSE <39><39> END AS [Object], <13,10>' & |
            ' <39>schema<39>+ CASE WHEN ob.parent_object_id>0 THEN <39>/table<39>ELSE <39><39> END <13,10>' & |
            ' + <39>/<39>+ <13,10>' & |
            ' CASE WHEN ob.type IN (<39>TF<39>,<39>FN<39>,<39>IF<39>,<39>FS<39>,<39>FT<39>) THEN <39>function<39> <13,10>' & |
            ' WHEN ob.type IN (<39>P<39>, <39>PC<39>,<39>RF<39>,<39>X<39>) then <39>procedure<39> <13,10>' & |
            ' WHEN ob.type IN (<39>U<39>,<39>IT<39>) THEN <39>table<39> <13,10>' & |
            ' WHEN ob.type=<39>SQ<39> THEN <39>queue<39> <13,10>' & |
            ' ELSE LOWER(ob.type_desc) end <13,10>' & |
            ' + CASE WHEN col.column_id IS NULL THEN <39><39> ELSE <39>/column<39>END AS [Path], <13,10>' & |
            ' ep.name,value <13,10>' & |
            ' FROM sys.extended_properties ep <13,10>' & |
            ' inner join sys.objects ob ON ep.major_id=ob.OBJECT_ID AND class=1 <13,10>' & |
            ' LEFT outer join sys.columns col <13,10>' & |
            ' ON ep.major_id=col.Object_id AND class=1 <13,10>' & |
            ' AND ep.minor_id=col.column_id <13,10>' & |
            '  WHERE (SELECT CASE WHEN ob.parent_object_id>0 <13,10>' & |
            ' THEN OBJECT_SCHEMA_NAME(ob.parent_object_id) <13,10>' & |
            ' + <39>.<39>+OBJECT_NAME(ob.parent_object_id)+<39>.<39>+ob.name <13,10>' & |
            ' ELSE OBJECT_SCHEMA_NAME(ob.object_id)+<39>.<39>+ob.name END <13,10>' & |
            ' + CASE WHEN ep.minor_id>0 THEN <39>.<39>+col.name ELSE <39><39> END) = <39>' & pObjectName & '<39>',|
            SELF.qExtendedProperties,SELF.qExtendedProperties.Object,SELF.qExtendedProperties.Path,SELF.qExtendedProperties.Name,SELF.qExtendedProperties.Value)
    
        
! -----------------------------------------------------------------------
!!! <summary>Returns the length of a Column in a Table</summary>           
!!! <param name="Table">Table Name</param>
!!! <param name="Column">Column Name</param>        
! -----------------------------------------------------------------------        
UltimateSQL.GetColumnLength     PROCEDURE(STRING pTable,STRING pColumn)   ! ,LONG
       
Result                              LONG

    CODE 
        
    SELF.QUERY('SELECT Character_Maximum_Length FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ' & SELF.Quote(pTable) & ' AND COLUMN_NAME = ' & SELF.Quote(pColumn),,Result)
        
    RETURN Result

        
! -----------------------------------------------------------------------
!!! <summary>Format text value with surrounding quotes</summary>
!!! <param name="Text">Value to be quoted</param>
!!! <remarks>Format the text value for processing. Format depends
!!! on driver being used.</remarks>
! -----------------------------------------------------------------------
UltimateSQL.Quote               PROCEDURE(STRING pText,BYTE pEscape=0)

oUltimateDB                         UltimateSQLString

    CODE                            
        
    oUltimateDB.Assign(CLIP(pText))   
    oUltimateDB.Quote(SELF._Driver)  
    IF pEscape
        RETURN ' <39>' & CLIP(oUltimateDB.Get()) & '<39> ' 
    END
    
    RETURN CLIP(oUltimateDB.Get())    


! -----------------------------------------------------------------------
!!! <summary>Handles display of Query errors</summary>
!!! <param name="ErrorCode">The Errorcode()</param>
!!! <param name="Error">The Error()</param>
!!! <param name="FileErrorCode">The FileErrorcode()</param>
!!! <param name="FileError">The FileError()</param>
!!! <remarks>Formats the error message and displays it in DebugView, the Clipboard, or as a Message
!!! ErrorShowInDebugView = TRUE shows the error in DebugView        
!!! ErrorAddToClipboard = TRUE adds the error to the Clipboard       
!!! ErrorAppendToClipboard = TRUE appends the error to the Clipboard        
!!! ErrorShowAsMessage = TRUE displays the error as a Message</remarks>
! -----------------------------------------------------------------------        
UltimateSQL.HandleError         PROCEDURE(LONG pErrorCode,STRING pError,LONG pFileErrorCode,STRING pFileError)  

ErrorMessage                        STRING(300)

    CODE 

    ErrorMessage  =  'Error: ' & ERRORCODE() & '=' & ERROR() & ' | Errorcode: ' & FILEERRORCODE() & ' = ' & FILEERROR() & '.' 
    IF ~SELF.AllDebugOff    
        
        IF SELF.ErrorShowInDebugView  
            SELF.Debug(ErrorMessage)  
        END
            
        IF SELF.ErrorAppendToClipboard
            SETCLIPBOARD(CLIP(CLIPBOARD()) & '<13,10,13,10>Error - ' & FORMAT(TODAY(),@D17) & ',' & FORMAT(CLOCK(),@T7) & '<13,10>' & CLIP(ErrorMessage))
        ELSIF SELF.ErrorAddToClipboard
            SETCLIPBOARD(CLIP(ErrorMessage))
        END                
        
        IF SELF.ErrorShowAsMessage
            MESSAGE(CLIP(ErrorMessage))
        END 
    END
    
        
! -----------------------------------------------------------------------
!!! <summary>Insert record into database</summary>
!!! <param name="Tbl">Reference to table</param>
!!! <param name="IncludePK">[F]=Do not include primary key, T=Include primary key</param>
!!! <remarks>Uses SQL to build an insert into the database for a table.
!!! Returns True on success, False on failure. Assumes that all the fields
!!! of the table have been filled into the record buffer.
!!! Generated SQL:
!!!     INSERT INTO tablename (f1, ...) VALUES (v1, ...)</remarks>
! -----------------------------------------------------------------------        
UltimateSQL.Insert              PROCEDURE(*FILE pTbl, BYTE pIncludePK=FALSE, BYTE pGetIdentity=FALSE)   !, LONG, PROC   

Identity                            LONG
Result                              LONG

    CODE
    
    Result = SELF.QueryResult(SELF.GetInsertStatement(pTbl))
    IF pGetIdentity
        RETURN SELF.QueryResult('SELECT @@identity')
        
    ELSE
        RETURN Result
        
    END
    
    
! ----------------------------------------------------------------------- 
UltimateSQL.GetInsertStatement          PROCEDURE(*FILE pTbl)   !, STRING, PROC   

oUltimateDB                                 UltimateDB

    CODE
    
    RETURN oUltimateDB.SQLInsert(pTbl)

        
! -----------------------------------------------------------------------
!!! <summary>Update record in database</summary>
!!! <param name="Tbl">Reference to table</param>
!!! <remarks>Uses SQL to build an update to the database for a table.
!!! Returns True on success, False on failure. Assumes that all the fields
!!! of the table have been filled into the record buffer. Does NOT handle
!!! blobs.
!!! Generated SQL:
!!!     UPDATE tablename (f1 = v1, ...) WHERE pk = pkValue</remarks>
! -----------------------------------------------------------------------        
UltimateSQL.Update              PROCEDURE(*FILE pTbl)   !, BYTE, PROC  

    CODE
    
    RETURN SELF.QueryResult(SELF.GetUpdateStatement(pTbl))
                                           

! -----------------------------------------------------------------------     
UltimateSQL.GetUpdateStatement          PROCEDURE(*FILE pTbl)   !, STRING

oUltimateDB                                 UltimateDB

    CODE
    
    RETURN oUltimateDB.SQLUpdate(pTbl)
    
    
! -----------------------------------------------------------------------
!!! <summary>Delete record in database</summary>
!!! <param name="Tbl">Reference to table</param>
!!! <remarks>Returns True on success, False on failure. Uses SQL
!!! to delete record using primary key value. Assumes that the primary
!!! key value has been filled into the record buffer. Example:
!!!     APinv:APInvoiceID = 1
!!!     db.SQLDelete(APinv)
!!! SQL Generated:
!!!     DELETE FROM tablename WHERE pk = pkValue</remarks>
! -----------------------------------------------------------------------        
UltimateSQL.Delete              PROCEDURE(*FILE pTbl)   !, BYTE, PROC  

    CODE
    
    RETURN SELF.QueryResult(SELF.GetDeleteStatement(pTbl))


! -----------------------------------------------------------------------     
UltimateSQL.GetDeleteStatement              PROCEDURE(*FILE pTbl)   !, STRING

oUltimateDB                         UltimateDB

    CODE
    
    RETURN oUltimateDB.SQLDelete(pTbl)
    
   
! ----------------------------------------------------------------------- 
UltimateSQL.SetCustomConnectionString   PROCEDURE(STRING pCustomConnectionString)  

    CODE
    
    SELF.CustomConnectionString  =  pCustomConnectionString
    
    
! -----------------------------------------------------------------------
!!! <summary>Sends MultipleActiveResultSets Driver String</summary>           
!!! <param name="TrueFalse">Byte set to TRUE or FALSE</param>
! -----------------------------------------------------------------------                
UltimateSQL.SetMultipleActiveResultSets PROCEDURE(BYTE pTrueFalse)

    CODE
               
    SELF.vMultipleActiveResultSets  =  '/MULTIPLEACTIVERESULTSETS = ' & CHOOSE(pTrueFalse,'TRUE','FALSE')
             
        
! -----------------------------------------------------------------------
!!! <summary>Sends VerifyViaSelect Driver String</summary>           
!!! <param name="TrueFalse">Byte set to TRUE or FALSE</param>
! -----------------------------------------------------------------------
UltimateSQL.SetVerifyViaSelect          PROCEDURE(BYTE pTrueFalse) 

    CODE
    
    SELF.vVerifyViaSelect  =  '/VERIFYVIASELECT = ' & CHOOSE(pTrueFalse,'TRUE','FALSE')
                  
    
! -----------------------------------------------------------------------
!!! <summary>Sends SaveStoredProcedure Driver String</summary>           
!!! <param name="TrueFalse">Byte set to TRUE or FALSE</param>
! -----------------------------------------------------------------------
UltimateSQL.SetSaveStoredProcedure      PROCEDURE(BYTE pTrueFalse)

    CODE
                    
    SELF.vSaveStoredProcedure  =  '/SAVESTOREDPROCEDURE = ' & CHOOSE(pTrueFalse,'TRUE','FALSE')
            
        
! -----------------------------------------------------------------------
!!! <summary>Sends GatherAtOpen Driver String</summary>           
!!! <param name="TrueFalse">Byte set to TRUE or FALSE</param>
! -----------------------------------------------------------------------
UltimateSQL.SetGatherAtOpen     PROCEDURE(BYTE pTrueFalse)

    CODE
        
    SELF.vGatherAtOpen  =  '/GATHERATOPEN = ' & CHOOSE(pTrueFalse,'TRUE','FALSE')
    
    
! -----------------------------------------------------------------------
!!! <summary>Sends LogonScreen Driver String</summary>           
!!! <param name="TrueFalse">Byte set to TRUE or FALSE</param>
! -----------------------------------------------------------------------
UltimateSQL.SetLogonScreen      PROCEDURE(BYTE pTrueFalse) !,STRING,PROC
      
    CODE
         
    SELF.vLogonScreen  =  '/LOGONSCREEN = ' & CHOOSE(pTrueFalse,'TRUE','FALSE')
    
    
! -----------------------------------------------------------------------
!!! <summary>Sends TurboSQL Driver String</summary>           
!!! <param name="TrueFalse">Byte set to TRUE or FALSE</param>
! -----------------------------------------------------------------------
UltimateSQL.SetTurboSQL         PROCEDURE(BYTE pTrueFalse)

    CODE        
     
    SELF.vTurboSQL  =  '/TURBOSQL = ' & CHOOSE(pTrueFalse,'TRUE','FALSE')
    
    
! -----------------------------------------------------------------------
!!! <summary>Sends IgnoreTruncation Driver String</summary>           
!!! <param name="TrueFalse">Byte set to TRUE or FALSE</param>
! -----------------------------------------------------------------------
UltimateSQL.SetIgnoreTruncation         PROCEDURE(BYTE pTrueFalse)

    CODE
            
    SELF.vIgnoreTruncation  =  '/IGNORETRUNCATION = ' & CHOOSE(pTrueFalse,'TRUE','FALSE')

    
! -----------------------------------------------------------------------
!!! <summary>Sends BusyHandling Driver String</summary>           
!!! <param name="BusyHandling">Numeric from 1 to 4 indicating the type of BusyHandling to use</param>
! -----------------------------------------------------------------------
UltimateSQL.SetBusyHandling     PROCEDURE(BYTE pBusyHandling)        

    CODE
            
    SELF.vBusyHandling  =  '/BUSYHANDLING = ' & pBusyHandling
    
    
! -----------------------------------------------------------------------
!!! <summary>Sends BusyRetries Driver String</summary>           
!!! <param name="BusyRetries">Numeric from 1 to whatever indicating the number of Retries to use</param>
! -----------------------------------------------------------------------

UltimateSQL.SetBusyRetries      PROCEDURE(LONG pBusyRetries)

    CODE
        
    SELF.vBusyRetries  =  '/BUSYRETRIES = ' & pBusyRetries
        
           
! -----------------------------------------------------------------------
!!! <summary>Send a message to the File Driver</summary>
!!! <param name="Message">Message to send to the driver</param>
! -----------------------------------------------------------------------        
UltimateSQL.SendDriverString    PROCEDURE(STRING pMessage) !,STRING,PROC

    CODE
    
    RETURN SEND(QueryResultsTemp,CLIP(pMessage))
    
! -----------------------------------------------------------------------        
!!! Sends PRAGMA statement to SQLite table        
! -----------------------------------------------------------------------        
UltimateSQL.SetPRAGMA           PROCEDURE(STRING pPragma)
! -----------------------------------------------------------------------        
QueryResultsSQLite                  FILE,DRIVER('SQLite',DriverOptions),PRE(QueryResultsSQLite),BINDABLE,THREAD
Record                                  RECORD,PRE()
QueryFields                                 LIKE(QueryFieldsGroup)
                                        END   
                                    END  
    CODE
    
    SELF.QueryTableName  =  'PragmaTempTable'
    SELF.AppendToDriverOptions() 
    QueryResultsSQLite{PROP:Owner}  =  SELF.ConnectionString
    QueryResultsSQLite{PROP:Name}   =  SELF.QueryTableName   
    OPEN(QueryResultsSQLite)
    QueryResultsSQLite{PROP:SQL}  =  'PRAGMA ' & CLIP(pPragma)
    CLOSE(QueryResultsSQLite)
                                  
    
UltimateSQL.AppendToDriverOptions       PROCEDURE()

                                MAP
SetUSQLDriverOptions                PROCEDURE()
                                END

    CODE   
    
    SetUSQLDriverOptions()
    
    IF CLIP(SELF.USQLDriverOptions) = ''
        SELF.SetDriverOptionDefaults()
        SetUSQLDriverOptions()
        
    END
    
    IF SELF._Driver  = ULSDriver_SQLite
        SELF.USQLDriverOptions  =  '/TURBOSQL = True '
    ELSE
        SELF.USQLDriverOptions  =  '/TURBOSQL = True /LOGONSCREEN = FALSE ' & CLIP(SELF.USQLDriverOptions)   
    END
    
    DriverOptions  =  SELF.USQLDriverOptions
    
    RETURN 

    
SetUSQLDriverOptions            PROCEDURE()

    CODE
    
    SELF.USQLDriverOptions  =  CLIP(SELF.vBusyHandling)
    SELF.USQLDriverOptions  =  CLIP(SELF.USQLDriverOptions) & CHOOSE(SELF.USQLDriverOptions = '','',' ') & CLIP(SELF.vBusyRetries)
    SELF.USQLDriverOptions  =  CLIP(SELF.USQLDriverOptions) & CHOOSE(SELF.USQLDriverOptions = '','',' ') & CLIP(SELF.vGatherAtOpen)
    SELF.USQLDriverOptions  =  CLIP(SELF.USQLDriverOptions) & CHOOSE(SELF.USQLDriverOptions = '','',' ') & CLIP(SELF.vIgnoreTruncation)
    SELF.USQLDriverOptions  =  CLIP(SELF.USQLDriverOptions) & CHOOSE(SELF.USQLDriverOptions = '','',' ') & CLIP(SELF.vMultipleActiveResultSets)
    SELF.USQLDriverOptions  =  CLIP(SELF.USQLDriverOptions) & CHOOSE(SELF.USQLDriverOptions = '','',' ') & CLIP(SELF.vSaveStoredProcedure)
    SELF.USQLDriverOptions  =  CLIP(SELF.USQLDriverOptions) & CHOOSE(SELF.USQLDriverOptions = '','',' ') & CLIP(SELF.vVerifyViaSelect)
         

UltimateSQL.SetDriverOptionDefaults     PROCEDURE()

    CODE                    
                         
    SELF.SetMultipleActiveResultSets(GETINI('DriverOptions','MultipleActiveResultSets',TRUE,'.\UltimateSQL.ini'))
    SELF.SetVerifyViaSelect(GETINI('DriverOptions','VerifyViaSelect',FALSE,'.\UltimateSQL.ini'))
    SELF.SetSaveStoredProcedure(GETINI('DriverOptions','SaveStoredProcedure',TRUE,'.\UltimateSQL.ini'))
    SELF.SetGatherAtOpen(GETINI('DriverOptions','GatherAtOpen',TRUE,'.\UltimateSQL.ini'))
    SELF.SetIgnoreTruncation(GETINI('DriverOptions','IgnoreTruncation',FALSE,'.\UltimateSQL.ini'))
    SELF.SetBusyHandling(GETINI('DriverOptions','BusyHandling',2,'.\UltimateSQL.ini'))     
    SELF.SetBusyRetries(GETINI('DriverOptions','BusyRetries',300,'.\UltimateSQL.ini'))
    
    
UltimateSQL.ClearErrors         PROCEDURE()

ThreadNumber                        LONG
lc                                  LONG

    CODE
    
    SELF.Wait(4)  
        
    ThreadNumber  =  THREAD()
    
    LOOP lc = RECORDS(SELF.qErrors) TO 1 BY -1
        GET(SELF.qErrors,lc)       
        IF SELF.qErrors.Thread = ThreadNumber
            DELETE(SELF.qErrors)
        END
    END

    SELF.Release(4)

    
UltimateSQL.SetError            PROCEDURE(LONG pErrorCode,STRING pErrorDescription)  

    CODE  
    
    IF pErrorCode = 33;RETURN.
    
    SELF.Wait(5)  
    
    SELF.qErrors.Thread            =  THREAD()
    SELF.qErrors.ErrorCode         =  pErrorCode
    SELF.qErrors.ErrorDescription  =  pErrorDescription
    ADD(SELF.qErrors)
    
    SELF.Release(5)
    
    
UltimateSQL.GetError            PROCEDURE(<STRING pStatementID>) !,STRING 
        
ThreadNumber                        LONG
lc                                  LONG
em                                  UltimateString
        
qDBError                            QUEUE,PRE(qDBError)
GUID                                    STRING(36),NAME('GUID')
Message                                 STRING(500),NAME('Message')
BatchLine                               LONG,NAME('BatchLine')
Statement                               STRING(100000),NAME('Statement')
DateTime                                STRING(30),NAME('DateTime')
                                    END


    CODE
     
    SELF.Wait(6)  
    IF pStatementID
        SELF.Query('Select * From ##DB_Errors Where GUID = ' & SELF.Quote(pStatementID) & ' order by datetime',qDBError)
        LOOP lc = 1 TO RECORDS(qDBError)
            GET(qDBError,lc) 
            em.Append('Line ' & qDBError.BatchLine & ', ' & CLIP(qDBError.Message) & ', ' & CLIP(qDBError.Statement) & '<13,10>')
        END
        
    ELSE
        ThreadNumber  =  THREAD()
        LOOP lc = 1 TO RECORDS(SELF.qErrors)
            GET(SELF.qErrors,lc)
            IF SELF.qErrors.Thread = ThreadNumber
                em.Append(CLIP(SELF.qErrors.ErrorCode) & ', ' & CLIP(SELF.qErrors.ErrorDescription) & '<13,10>')
            END
        
        END 
       
        SELF.ClearErrors()
        
    END
    
    SELF.Release(6)
    
    RETURN em.Get()

    
! -----------------------------------------------------------------------
!!! <summary>Turn on/off tracing</summary>
!!! <param name="Tbl">Reference to table</param>
!!! <param name="Logfile">Name of file where trace is to occur</param>
!!! <remarks>Omitting the Logfile argument turns off tracing.</remarks>
! -----------------------------------------------------------------------        
UltimateSQL.Trace               PROCEDURE(*FILE pTbl, <STRING pLogfile>)
        
oUltimateDB                         UltimateDB

    CODE                                      
        
    oUltimateDB.SQLLog(pTbl,pLogfile)
    
    RETURN      
        
!-----------------------------------------
UltimateSQL.RemoveIllegalCharacters     PROCEDURE(String pString)
!-----------------------------------------
Result                                      String(5000)

    CODE
    Result  =  pString
    Y# = 1
    LOOP
        X# = INSTRING('<39>',Result,1,Y#)
        IF ~X#;BREAK.
        Result  =  SUB(Result,1,X#-1) & '<39><39>' & SUB(Result,X#+1,10000)
        Y# = X# + 2
    END

    RETURN Result        

    
!-----------------------------------------
!!! <summary>Adds Brackets around the passed string. Useful when using reserved words or labels with spaces</summary>
!!! <param name="STRING">STRING Value to add brackets to</param>
!!! <returns>Bracketed string</returns>
!!! <remarks>Example: Pass in CUSTOMER returns [CUSTOMER]
!!! Example: Pass in dbo.CUSTOMER returns [dbo].[CUSTOMER]</remarks>
UltimateSQL.Bracket             PROCEDURE(STRING pValue)  !,STRING
!-----------------------------------------

PeriodPosition                      LONG
Schema                              STRING(100)
Name                                STRING(200)

    CODE
    
    PeriodPosition  =  INSTRING('.',pValue,1,1)
    IF PeriodPosition = 0
        RETURN CLIP('[' & CLIP(pValue) & ']')
        
    END 

    Schema  =  SUB(pValue,1,PeriodPosition-1)
    Name    =  SUB(pValue,PeriodPosition+1,200)
    RETURN '[' & CLIP(Schema) & ']' & '.' & '[' & CLIP(Name) & ']'
        
	
!----------------------------------------
UltimateSQL.GetAutoFill         PROCEDURE(*CSTRING pServer,*CSTRING pOwnerName,*BYTE pWindowsAuthentication)
!----------------------------------------
    CODE
    
    IF SELF.AutoFill
        IF pServer = ''
            pServer  =  GETINI(SELF.AutoFillSection,'Server','',SELF.AutoFillINIFile)
        END
        IF pOwnerName = ''
            pOwnerName  =  GETINI(SELF.AutoFillSection,'User','',SELF.AutoFillINIFile)
        END
        IF pWindowsAuthentication = ''
            pWindowsAuthentication  =  GETINI(SELF.AutoFillSection,'WindowsAuthentication','',SELF.AutoFillINIFile)
        END
    END
    
    
    
!----------------------------------------
UltimateSQL.SaveAutoFill        PROCEDURE(*CSTRING pServer,*CSTRING pOwnerName,*BYTE pWindowsAuthentication)
!----------------------------------------
    CODE
    
    IF SELF.AutoFill
        PUTINI(SELF.AutoFillSection,'Server',pServer,SELF.AutoFillINIFile)
        PUTINI(SELF.AutoFillSection,'User',pOwnerName,SELF.AutoFillINIFile)
        PUTINI(SELF.AutoFillSection,'WindowsAuthentication',pWindowsAuthentication,SELF.AutoFillINIFile)
    END


!!! <summary>Checks your SQL query for Clarion prefixes (ex: CUS:) and replaces the ":" with an underscore.
!!! <param name=""STRING"">SQL Query statement</param>
!!! <returns>The modified Query</returns>
UltimateSQL.CheckClarionTablePrefixes   PROCEDURE (STRING pQuery) !,STRING
!----------------------------------------
 
DelimitorPosition                           LONG
InQuotedField                               BYTE
StringLength                                LONG
StringPosition                              LONG
 
    CODE
  
    DelimitorPosition  =  1
    InQuotedField      =  False
    StringLength       =  LEN(CLIP(pQuery))
    StringPosition     =  0
  
    IF StringLength > 2
     
        LOOP StringPosition = 1 TO StringLength
        
            IF pQuery [StringPosition] = CHR(34) OR pQuery [StringPosition] = CHR(39)
                IF InQuotedField = True
                    InQuotedField  =  False
                ELSE
                    InQuotedField  =  True
                END
            END
        
            IF InQuotedField = True THEN CYCLE.
        
            IF pQuery [StringPosition] = CHR(32) OR pQuery [StringPosition] = ',' OR pQuery [StringPosition] = '(' OR pQuery [StringPosition] = '%'
                DelimitorPosition  =  StringPosition
            END
        
            IF pQuery [StringPosition] = ':'
                IF SELF.FieldSeparatorReplacement 
                    pQuery  =  pQuery [1 : StringPosition - 1] & CLIP(SELF.FieldSeparatorReplacement) & pQuery [StringPosition + 1 : StringLength]
                   
                ELSE
                    IF DelimitorPosition = 1
                        pQuery  =  pQuery [StringPosition + 1 : StringLength]
                    ELSE
                        pQuery  =  pQuery [1 : DelimitorPosition] & pQuery [StringPosition + 1 : StringLength]
                    END 
                    
                END
           
            END
        
        END
     
    END
     
    RETURN pQuery   
    
    
!----------------------------------------
UltimateSQL.GetLocalSQLServers          PROCEDURE()
!----------------------------------------    
qRegParent                                  QUEUE,PRE(qRegParent)
ParentKey                                       STRING(100)
                                            END  

qRegChild                                   QUEUE,PRE(qRegChild)
ChildKey                                        STRING(100)
                                            END  

MachineName                                 STRING(100)

lc                                          LONG

    CODE
    
    MachineName  =  GETREG(REG_LOCAL_MACHINE,'SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName','ComputerName')   
    
    FREE(SELF.qSQLServerList)
    GETREGSUBKEYS(REG_LOCAL_MACHINE,'SOFTWARE\Microsoft\Microsoft SQL Server',qRegParent)  

    
    
    LOOP lc = 1 TO RECORDS(qRegParent)
        GET(qRegParent,lc)
        FREE(qRegChild)
        GETREGSUBKEYS(REG_LOCAL_MACHINE,'SOFTWARE\Microsoft\Microsoft SQL Server\' & CLIP(qRegParent.ParentKey),qRegChild)  
        IF qRegChild.ChildKey = 'MSSQLServer'
            SELF.qSQLServerList.Name  =  CLIP(MachineName) & '\' & qRegParent.ParentKey
            ADD(SELF.qSQLServerList)
        END
        
    END

!----------------------------------------
UltimateSQL.Construct           PROCEDURE()
!----------------------------------------    

    CODE
         
    SELF.qExtendedProperties  &=  NEW qExtendedPropertiesType
    SELF.qSQLServerList       &=  NEW qSQLServerListType  
    SELF.qErrors              &=  NEW qErrorsType
    SELF.qStatements          &=  NEW qStatementsType
    SELF.qBatchErrors         &=  NEW qBatchErrorsType
   
    SELF.USQLCriticalSection  &=  NewCriticalSection()  
    
    RETURN

		
!---------------------------------------
UltimateSQL.Destruct            PROCEDURE()
!---------------------------------------

    CODE
        
    FREE(SELF.qExtendedProperties)
    DISPOSE(SELF.qExtendedProperties)
    
    FREE(SELF.qSQLServerList)
    DISPOSE(SELF.qSQLServerList)
    
    FREE(SELF.qErrors)
    DISPOSE(SELF.qErrors)   
    
    FREE(SELF.qStatements)
    DISPOSE(SELF.qStatements)
    
    FREE(SELF.qBatchErrors)
    DISPOSE(SELF.qBatchErrors)
    
    IF ~SELF.USQLCriticalSection &= NULL
        SELF.USQLCriticalSection.Kill()
    END
    
    RETURN
    
    
!Comment Format below:

! -----------------------------------------------------------------------
!!! <summary>Description</summary>           
!!! <param name="Table">Database name</param>
!!! <remarks>Remarks.</remarks>
! -----------------------------------------------------------------------
    
!----------------------------------------
UltimateSQL.PrepStatement       PROCEDURE(STRING pStatement,BYTE pSingleQuote=0,BYTE pBracket=0)  !,STRING
!----------------------------------------
   
ReturnStatement                     CSTRING(500000)
APPosition                          LONG

    CODE

    ReturnStatement  =  pStatement
    
    LOOP
        APPosition  =  INSTRING('''',ReturnStatement,1,APPosition + 2)
        IF ~APPosition;BREAK.
        ReturnStatement  =  ReturnStatement[1 : APPosition-1] & '<39><39>' & SUB(ReturnStatement,APPosition+ 1,500000)
        
    END
    
    IF pBracket
        ReturnStatement  =  '[' & CLIP(ReturnStatement) & ']' 
    END
    
    IF pSingleQuote
        ReturnStatement  =  '<39>' & CLIP(ReturnStatement) & '<39>'
    END
    
    RETURN CLIP(ReturnStatement)               
          
             
UltimateSQL.CheckForDebug       PROCEDURE(STRING pQueryToCheck) !,STRING
      
sQuery                              UltimateString

    CODE
    
    IF pQueryToCheck[1:1] = '*' OR SELF.ShowQueryInDebugView OR SELF.AddQueryToClipboard OR SELF.AppendQueryToClipboard
        SELF.SendToDebug  =  FALSE
        IF pQueryToCheck[1:1] = '*'
            sQuery.Assign(SUB(pQueryToCheck,2,999000000))
            SELF.SendToDebug  =  TRUE   
        ELSE
            sQuery.Assign(pQueryToCheck)
            
        END
        
        IF ~SELF.AllDebugOff    
            IF SELF.ShowQueryInDebugView OR SELF.SendToDebug
                SELF.Debug(sQuery.Get())
                
            END
            
            IF SELF.AppendQueryToClipboard
                SETCLIPBOARD(CLIPBOARD() & '<13,10,13,10>' & CLIP(pQueryToCheck))
                
            ELSIF SELF.AddQueryToClipboard
                SETCLIPBOARD(CLIP(pQueryToCheck))
                
            END
            
        END 
        
    ELSE
        sQuery.Assign(pQueryToCheck)
        
    END 

    RETURN sQuery.Get()

    
UltimateSQL.Wait                Procedure(Long pId)

    code
        
    SELF.Trace('wait ' & pId)
    SELF.USQLCriticalSection.wait()
        
        
!------------------------------------------------------------------------------
UltimateSQL.Release             Procedure(Long pId)

    code
        
    SELF.Trace('release ' & pId)
    SELF.USQLCriticalSection.release()

        
!------------------------------------------------------------------------------
UltimateSQL.Trace               Procedure(string pStr)

szMsg                               cString(len(pStr)+10)

    code
    
    IF SELF.TraceOn
        szMsg  =  '[un] ' & Clip(pStr)
        ! SELF.DebugPrefix  =  'UNM'
        ! SELF.DebugOff     =  FALSE
        ! SELF.Debug(szMsg)
    END 
    
!region Begin Comments


!!>Version         2021.12.18.1
!!>Template        UltimateSQL.tpl
!!>Revisions       Fixed "CreateTemporaryProcedure" Method
!!>                Added "DropTemporaryProcedure" Method
!!>                Added "CheckForDebug" Method (eliminating duplicated code)
!!>Created         DEC  3,2020,  5:40:11am
!!>Modified        DEC 18,2021, 10:18:43am 
 
!!>Version         2021.12.15.1
!!>Revisions       Added SQLite In-Memory support
!!>                Optimized MSSQL batch processing
!!>                MSSQL Native Client is required
!!>                Change: ExecuteScript is now ExecuteScriptFromFile
!!>                Added VIRTUAL to all methods
!!>                New: SetPragma statement for SQLite
!!>
!!>Created         DEC  2,2020,  9:35:39am
!!>Modified        AUG 27,2021,  7:08:17am
!endregion End Comments