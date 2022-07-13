                                MEMBER
                                MAP
                                    MODULE('API')
                                        OutputDebugSTRING(*cstring xMesage),PASCAL,RAW,NAME('OutputDebugStringA')
                                    END
                                END
    INCLUDE('ctFieldQ.inc'),ONCE

ctFieldQ.ODS                    PROCEDURE(String xMessage) ! clarionized OutputDebugString
szMsg                               &CSTRING

    CODE 
    
    szMsg  &=  NEW CSTRING( SIZE(xMessage) + 1)
    szMsg   =  xMessage 
    OutputDebugString( szMsg )
    DISPOSE( szMsg )
    
    
ctFieldQ.Construct              PROCEDURE

    CODE
    
    SELF.Q  &=  NEW qtAny
    CLEAR(SELF.Q)
    
    
ctFieldQ.Destruct               PROCEDURE

    CODE
    
    IF NOT (SELF.Q &= NULL)
        SELF.Free()
        DISPOSE(SELF.Q)
    END
    
    
ctFieldQ.Free                   PROCEDURE

    CODE
    
    IF (SELF.Q &= NULL) 
        RETURN 
    END
  
    LOOP
        GET(SELF.Q, 1)
        IF ERRORCODE() THEN BREAK END
    
        SELF.Delete()     
        
    END
    
    
ctFieldQ.Delete                 PROCEDURE
                                    ! Help[ANY (any simple data type)]
                                    ! "...
                                    ! In addition, you need to reference assign a NULL to all queue fields of type ANY (Queue.AnyField &= NULL),
                                    ! prior to deleting the QUEUE entry, in order to avoid memory leaks.
                                    ! ..."  
    CODE
    SELF.Q.Field  &=  NULL
    DELETE(SELF.Q)
  

ctFieldQ.Add                    PROCEDURE(*? xOrigField)

    CODE
    
    CLEAR(SELF.Q)
    SELF.Q.Field  &=  xOrigField
    ADD(SELF.Q)
    
    
ctFieldQ.AddFieldsInGroup       PROCEDURE(*GROUP xGroup)
_What                               ANY 
CurrField                           LONG,AUTO 

    CODE
    
    CurrField  =  1
    LOOP 
        _What  &=  WHAT( xGroup, CurrField ) 
        IF        _What &= NULL THEN BREAK END 
        SELF.Add( _What ) 
        CurrField  +=  1 
    END 
    
    
ctFieldQ.Records                PROCEDURE()

    CODE 
    
    RETURN RECORDS(SELF.Q)
    
    
ctFieldQ.GetRow                 PROCEDURE(LONG xRow)!,LONG,PROC             ! Returns ErrorCode()

    CODE 
    
    GET(SELF.Q, xRow)
    RETURN ErrorCode()
    
    
ctFieldQ.SetOrigField           PROCEDURE(LONG xRow, ? xNewValue)
RetError                            LONG,AUTO

    CODE 
    
    RetError  =  SELF.GetRow( xRow ) 
    IF RetError = 0                    ! 0 is NoError
        SELF.Q.Field  =  xNewValue        ! this assignment will alter the field that was passed in to .Add
    END 

    RETURN RetError 

