!Created: 2019 Jan 23 by Mark Goldberg - mark@monolithcc.com

qtAny       QUEUE,TYPE 
Field         ANY ! a pointer to the orig field that was .Add()ed 
            END

ctFieldQ   CLASS,TYPE,MODULE('ctFieldQ.clw'),LINK('ctFieldQ.clw') !,_ABCLinkMode_),DLL(_ABCDllMode_)
Q                 &qtAny
 
CONSTRUCT         PROCEDURE
DESTRUCT          PROCEDURE
Free              PROCEDURE
Delete            PROCEDURE
Add               PROCEDURE(*? xOrigField)
GetRow            PROCEDURE(LONG xRow             ),LONG,PROC ! Returns ErrorCode()
SetOrigField      PROCEDURE(LONG xRow, ? xNewValue),LONG,PROC ! Returns ErrorCode() from GetRow
AddFieldsInGroup  PROCEDURE(*GROUP xGroup)
Records           PROCEDURE(),LONG
ODS               PROCEDURE(STRING xMessage) ! a clarionized version of OutputDebugString
            END

