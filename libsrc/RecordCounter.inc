        OMIT('_EndOfInclude_',_IFDef_RecordCounter_)
_IFDef_RecordCounter_            EQUATE(1)
!

    INCLUDE('UltimateSQL.inc'),ONCE
    INCLUDE('UltimateDebug.inc'),ONCE



!--------------------------------------------------------------------------------
RecordCounter    CLASS,TYPE,MODULE('RecordCounter.CLW'),LINK('RecordCounter.CLW')  !,_ABCLinkMode_),DLL(_ABCDllMode_)

!Properties
!----------
InDebug                 BYTE
ud                      &UltimateDebug

!--------------------------------------------------------------------------------
! Add Methods here, which are just procedures or functions
!--------------------------------------------------------------------------------
!Methods
!-------
Init                    PROCEDURE()
Kill                    PROCEDURE()
Count                   PROCEDURE(UltimateSQL pSQL,STRING pTableName),LONG
RaiseError              PROCEDURE(STRING pErrorMsg)
Construct               PROCEDURE()
Destruct                PROCEDURE()
                                        END
        _EndOfInclude_
