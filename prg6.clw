prg6                    PROGRAM

                        MAP
                        END
                        INCLUDE( 'ODBCClass.inc' )

query                   CSTRING( 4096 )
wdw                     WINDOW( 'Table grid' ), CENTER, DOUBLE, GRAY, AT( , , 300, 300 )
                          PROMPT( '&Query:' ), AT( 5, 5 ), TRN
                          TEXT, USE( query ), AT( 5, 20, , 50 ), FONT( 'Courier New', 10 ), VSCROLL
                          LIST, USE( ?List ), AT( 5, 90, , 100 ), VSCROLL
                          BUTTON( '&Ok' ), USE( ?btnOK ), AT( , 5, 40, 12 ), DEFAULT
                        END

db                      ODBCClass
ds                      ODBCDataSet
szConnection            CSTRING( 512 )

  CODE
  szConnection = 'Driver={{PostgreSQL ANSI};server=localhost;database=yourdb;uid=youruser;pwd=yourpwd;'

  OPEN( wdw )

  IF NOT INRANGE( db.Connect( szConnection ), SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )
    IF MESSAGE( 'Database connection error<13,10>Continue anyway?', 'Connection error', ICON:Question, |
      BUTTON:Yes + BUTTON:No, BUTTON:Yes ) = BUTTON:No
      RETURN
    END
  END

  ds.Init( db, ?List )

  ACCEPT
    DISPLAY()

    CASE EVENT()
    OF EVENT:OpenWindow                             ! Cosmetics
      ?query{ PROP:Width } = 0{ PROP:Width } - 10
      ?List{ PROP:Width } = 0{ PROP:Width } - 10
      ?List{ PROP:LineHeight } = 11
      ?btnOK{ PROP:XPos } = 0{ PROP:Width } - ?btnOK{ PROP:Width } - 5

    OF EVENT:Accepted
      CASE FIELD()
      OF ?btnOK
        ds.Exec( query, ?List )                     ! Executes query and fills listbox
      END
    END
  END