!
UltimateSQLScripts              CLASS,TYPE,MODULE('UltimateSQLScripts.CLW'),LINK('UltimateSQLScripts.CLW')  !,_ABCLinkMode_),DLL(_ABCDllMode_)
!
!Properties
!----------
InDebug                             BYTE
!
!Methods
!-------
Construct                           PROCEDURE()
Destruct                            PROCEDURE()
Init                                PROCEDURE()
Kill                                PROCEDURE()
InsertExtendedProperty              PROCEDURE(),STRING
UpdateExtendedProperty              PROCEDURE(),STRING
RemoveExtendedProperty              PROCEDURE(),STRING 
GetExtendedProperty                 PROCEDURE(),STRING 
DropAllDependencies                 PROCEDURE(),STRING    
CreateQueryTable                    PROCEDURE(),STRING
RaiseError                          PROCEDURE(STRING pErrorMsg)
                                END



