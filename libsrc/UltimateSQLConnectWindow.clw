Window                          WINDOW('Connect'),AT(,,391,190),CENTER,FONT('Segoe UI',10)
                                    SHEET,AT(2,2,390,188),USE(?SHEET)
                                        TAB('Tab1'),USE(?TAB1)
                                            PANEL,AT(29,34,301,99),USE(?PANEL),BEVEL(1)
                                            STRING('Connecting to Server....'),AT(137,68),USE(?STRING),FONT('Segoe UI',14,,FONT:bold+FONT:italic)
                                        END
                                        TAB('Tab2'),USE(?TAB2)
                                            OPTION,AT(85,34,105,20),USE(LocalOrNetwork),TRN 
                                                RADIO('Local'),AT(86,41),USE(?OPTIONLocal),TRN
                                                RADIO('Network'),AT(127,41,47,10),USE(?OPTIONNetwork),TRN
                                            END
                                            COMBO(@s200),AT(86,57,224,12),USE(TheServer),VSCROLL,DROP(10),FROM(SQLServers),FORMAT('1020L(2)@s255@')
                                            LIST,AT(86,74,224,11),USE(?LISTAuthentication),DROP(2),FROM('Windows Authentication|SQL Server Authentication')
                                            ENTRY(@s200),AT(86,89,224),USE(TheUserName)
                                            ENTRY(@s200),AT(86,106,224),USE(ThePassword),PASSWORD
                                            COMBO(@s200),AT(86,123,224,12),USE(TheDatabase),VSCROLL,DROP(10),FROM(SQLDatabases),FORMAT('1020L(2)|M@s255@')
                                            BUTTON('OK'),AT(251,167,65,21),USE(?OkButton),DEFAULT
                                            BUTTON('Cancel'),AT(320,167,69,21),USE(?CancelButton)
                                            STRING('Username:'),AT(31,92),USE(?STRING3),TRN
                                            STRING('Database:'),AT(31,126),USE(?STRING2),TRN
                                            STRING('Server Host:'),AT(31,60),USE(?STRING1),TRN
                                            STRING('Password:'),AT(31,109),USE(?STRING4),TRN
                                            STRING('Authentication:'),AT(31,75),USE(?STRING5),TRN
                                            BUTTON('Test'),AT(3,167,65,21),USE(?BUTTONTest)
                                            PROMPT('Enter connection values below:'),AT(45,19,277,12),USE(?PROMPTInstructions),TRN
                                        END
                                    END
                                END