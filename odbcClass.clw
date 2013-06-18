  MEMBER

  MAP
  END

  INCLUDE( 'odbcClass.inc' ), ONCE
  INCLUDE( 'equates.clw' ), ONCE
  INCLUDE( 'property.clw' ), ONCE
  INCLUDE( 'keycodes.clw' ), ONCE
  INCLUDE( 'DebugQueue.inc' ), ONCE

wProcessing             WINDOW, DOUBLE, AT( , , 350, 45 ), CENTER, GRAY
                          STRING( '0%' ), AT( 5, 5 ), TRN
                          STRING( '50%' ), AT( 5, 5, 290 ), TRN, CENTER
                          STRING( '100%' ), AT( 270, 5, 20 ), TRN, RIGHT
                          PROGRESS, USE( ?csvProgress ), AT( 5, 15, 290, 12 ), SMOOTH
                          BUTTON( 'Cancelar' ), USE( ?wProcessingCancel ), AT( 300, 5, 45, 12 )
                        END

ODBCGetDiagCalls        QUEUE                               ! Tracks calls to GetDiag
handleType                LONG
handle                    LONG
status                    LONG
                        END

ODBCDefaulObject        &ODBCClass                          ! If first object, identifies itself as default

ODBCClass.Connect       PROCEDURE( STRING connectString )
desktopHnd              LONG
retCode                 SHORT
lConnLength             LONG

  CODE
  retCode = SQLAllocHandle( SQL_HANDLE_ENV, SQL_NULL_HANDLE, SELF.henv )
  IF INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )

    retCode = SQLSetEnvAttr( SELF.henv, SQL_ATTR_ODBC_VERSION, SQL_OV_ODBC3, SQL_IS_INTEGER )
    IF INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )

      retCode = SQLAllocHandle( SQL_HANDLE_DBC, SELF.henv, SELF.hdbc )
      IF INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )

        SQLSetConnectAttr( SELF.hdbc, SQL_LOGIN_TIMEOUT, 5, 0 )
        IF INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )
          desktopHnd = GetDesktopWindow()
          SELF.connRequest = connectString
          retCode = SQLDriverConnect( SELF.hdbc, desktopHnd, SELF.connRequest, |
            LEN( SELF.connRequest ), SELF.connectionString, SIZE( SELF.connectionString ), |
            lConnLength, SQL_DRIVER_NOPROMPT )

          SQLFreeHandle( SQL_HANDLE_ENV, SELF.henv )
        END
      END
    END
  END

  RETURN retCode

ODBCClass.GetConnectionString PROCEDURE
  CODE
  RETURN SELF.ConnectionString

ODBCClass.Prepare       PROCEDURE( ? variable )
  CODE
  CLEAR( SELF.PreparedVariables )
  SELF.PreparedVariables = variable
  ADD( SELF.PreparedVariables )

ODBCClass.Exec          PROCEDURE( STRING sqlCommand )
retCode                 SHORT
strParsed               BSTRING
szSQL                   &CSTRING
lToken                  LONG                                ! Current token ordinal being replaced

  CODE
  IF NOT SELF.hdbc THEN RETURN SQL_INVALID_HANDLE END
  retCode = SQLAllocHandle( SQL_HANDLE_STMT, SELF.hdbc, SELF.hstmt )
  IF NOT INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )
    RETURN retCode
  END
  strParsed = sqlCommand
  IF RECORDS( SELF.PreparedVariables )                      ! Are there any prepared variable to substitute query tokens?
    LOOP i# = 1 TO RECORDS( SELF.PreparedVariables )
      GET( SELF.PreparedVariables, i# )
      lToken = INSTRING( '$', strParsed, 1, lToken + 1 )
      IF NOT lToken THEN BREAK END
      strParsed = SUB( strParsed, 1, lToken - 1 ) & CLIP( SELF.PreparedVariables.value ) & SUB( strParsed, lToken + 1, LEN( CLIP( strParsed ) ) )
      lToken += 1
    END
  END
  szSQL &= NEW CSTRING( LEN( CLIP( strParsed ) ) + 1 )      ! Must use CSTRING to cope with ODBC API prototype
  szSQL = strParsed
  retCode = SQLExecDirect( SELF.hstmt, szSQL, SQL_NTS )
  IF NOT INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )
    ODBCGetDiag( SQL_HANDLE_STMT, SELF.hstmt, SELF )
  END
  DISPOSE( szSQL )
  IF RECORDS( SELF.PreparedVariables )                      ! Can't leave any prepared variables for next calls
    FREE( SELF.PreparedVariables )
  END
  RETURN SQL_SUCCESS

ODBCClass.Disconnect    PROCEDURE
  CODE
  RETURN SQLDisconnect( SELF.hdbc )

ODBCClass.Construct     PROCEDURE
  CODE
  IF ODBCDefaulObject &= NULL                               ! No object instantiated so far
    ODBCDefaulObject &= SELF                                ! Makes this object the default
  END
  SELF.PreparedVariables &= NEW ODBCPreparedVars

ODBCClass.Destruct      PROCEDURE
  CODE
  FREE( SELF.PreparedVariables )
  DISPOSE( SELF.PreparedVariables )

ODBCDataSet.Init        PROCEDURE( ODBCClass Connection, LONG listBox=0, <STRING name> )
  CODE
  SELF.Connection &= Connection
  IF listBox                                                ! User passed a listbox?
    SELF.SetListbox( listbox )
  END
  IF name
    SELF.SetName( name )
  ELSE
    SELF.SetName( 'Data Set: ' & ADDRESS( SELF ) )
  END

ODBCDataSet.setName     PROCEDURE( STRING name )
  CODE
  SELF.dataSetName = name

ODBCDataSet.BindDefaultObject   PROCEDURE
  CODE
  IF SELF.Connection &= NULL
    SELF.Connection &= ODBCDefaulObject
  END

ODBCDataSet.SetQuery    PROCEDURE( STRING query )
  CODE
  SELF.query = query

ODBCDataSet.SetListbox  PROCEDURE( LONG feq )
szFormat                CSTRING( 1024 )
lTarget                 LONG

  CODE
  szFormat = '20L(2)|M'                                     ! A valid generic list format
  SELF.listBox = feq
  CASE feq{ PROP:Type }                                     ! What type of listbox is it?
  OF CREATE:List                                            ! Regular listbox, even with a drop
    lTarget = feq                                           ! Adjust the target
    IF feq{ PROP:Format } THEN szFormat = feq{ PROP:Format } END ! Saves the current format, if any
  OF CREATE:Combo
    lTarget = feq{ PROP:ListFeq }                           ! The target is the hidden control of the combo
    IF lTarget{ PROP:Format } THEN szFormat = lTarget{ PROP:Format } END ! Saves the format of the hidden control
  END
  lTarget{ PROP:Format } = szFormat                         ! Make sure the listbox has a valid format
  feq{ PROP:VLBVal } = ADDRESS( SELF )
  feq{ PROP:VLBProc } = ADDRESS( SELF.GetCell )
  SELF.listBox{ PROP:Alrt, 255 } = CtrlMouseRight           ! This will call the run time user interface
  REGISTER( EVENT:AlertKey, ADDRESS( SELF.HandleClicks ), ADDRESS( SELF ), , SELF.listBox )

ODBCDataSet.Exec        PROCEDURE( STRING sqlCommand )
retCode                 SHORT
szSQL                   &CSTRING

  CODE
  SELF.BindDefaultObject()                                        ! Make sure there's one ODBC object bound
  IF NOT SELF.connection.hdbc THEN RETURN SQL_INVALID_HANDLE END
  retCode = SQLAllocHandle( SQL_HANDLE_STMT, SELF.connection.hdbc, SELF.hstmt )
  IF NOT INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )
    RETURN retCode
  END
  szSQL &= NEW CSTRING( LEN( CLIP( sqlCommand ) ) + 1 )
  szSQL = sqlCommand
  retCode = SQLExecDirect( SELF.hstmt, szSQL, SQL_NTS )
  IF NOT INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )
    ODBCGetDiag( SQL_HANDLE_STMT, SELF.hstmt, SELF.Connection, 'Exec( ' & sqlCommand & ' )' )
  END
  DISPOSE( szSQL )
  RETURN SQL_SUCCESS

ODBCDataSet.Exec        PROCEDURE( STRING sqlCommand, LONG feq=0 )
  CODE
  IF NOT sqlCommand THEN RETURN SQL_ERROR END
  IF SELF.Exec( sqlCommand ) = SQL_SUCCESS
    SELF.LoadResult
    IF feq AND feq{ PROP:Type } = CREATE:List
      SELF.LoadList( feq )
    END
  END
  RETURN SQL_SUCCESS

ODBCDataSet.Load        PROCEDURE( LONG feq=0 )
targetFeq               LONG

  CODE
  IF NOT SELF.query THEN RETURN END
  IF NOT INLIST( SELF.Exec( SELF.query ), SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )
    RETURN
  END
  SELF.LoadResult
  IF feq
    targetFeq = feq
  ELSE
    targetFeq = SELF.listBox
  END
  SELF.LoadList( targetFeq )

ODBCDataSet.LoadResult  PROCEDURE
retCode                 SHORT
cols                    LONG                                ! Number of columns in the result set
rows                    LONG                                ! Number of rows in the result set
lValuePtr               LONG                                ! Pointer to the value read
lLength                 LONG                                ! Length of the read value
columnNameLength        SHORT
columnType              SHORT
columnSize              LONG
decimalDigits           SHORT
nullable                SHORT
szColumnName            CSTRING( 64 )

  CODE
  SELF.BindDefaultObject()                                  ! Make sure there's one ODBC object bound
  retCode = SQLNumResultCols( SELF.hstmt, cols )            ! How many columns the query produced?
  FREE( SELF.ResultSetCols )
  IF NOT cols                                               ! No columns produced
    SELF.Reset()                                            ! Reset this
    RETURN                                                  ! and don't waste time processing
  END
  LOOP i# = 1 TO cols                                       ! Process all returned columns
    retCode = SQLDescribeCol( SELF.hstmt, i#, szColumnName, SIZE( szColumnName ), columnNameLength, |
      columnType, columnSize, decimalDigits, nullable )

    IF INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )
      SELF.ResultSetCols.name = CLIP( szColumnName )
      SELF.ResultSetCols.key = UPPER( SELF.ResultSetCols.name )
      SELF.ResultSetCols.size = columnSize
      SELF.ResultSetCols.type = columnType
      SELF.ResultSetCols.decimals = decimalDigits
      SELF.ResultSetCols.nullable = nullable

      CASE columnType                                       ! Determine address of target structure
      OF SQL_GUID
        SELF.ResultSetCols.targetAddr = ADDRESS( SELF.tagSQLGUID )
        SELF.ResultSetCols.targetSize = SIZE( SELF.tagSQLGUID )

      OF SQL_INTEGER
        SELF.ResultSetCols.targetAddr = ADDRESS( SELF.tagINTEGER )
        SELF.ResultSetCols.targetSize = SIZE( SELF.tagINTEGER )

      OF SQL_SMALLINT
        SELF.ResultSetCols.targetAddr = ADDRESS( SELF.tagSMALLINT )
        SELF.ResultSetCols.targetSize = SIZE( SELF.tagSMALLINT )

      OF SQL_BIGINT OROF SQL_REAL
        SELF.ResultSetCols.type = SQL_INTEGER
        SELF.ResultSetCols.targetAddr = ADDRESS( SELF.tagINTEGER )
        SELF.ResultSetCols.targetSize = SIZE( SELF.tagINTEGER )

      OF SQL_TYPE_DATE OROF SQL_INTERVAL_DAY
        SELF.ResultSetCols.targetAddr = ADDRESS( SELF.tagDATE_STRUCT )
        SELF.ResultSetCols.targetSize = SIZE( SELF.tagDATE_STRUCT )

      OF SQL_TYPE_TIMESTAMP
        SELF.ResultSetCols.targetAddr = ADDRESS( SELF.tagTIMESTAMP_STRUCT )
        SELF.ResultSetCols.targetSize = SIZE( SELF.tagTIMESTAMP_STRUCT )

      OF SQL_TYPE_TIME
        SELF.ResultSetCols.targetAddr = ADDRESS( SELF.tagTIME_STRUCT )
        SELF.ResultSetCols.targetSize = SIZE( SELF.tagTIME_STRUCT )

      ELSE
        SELF.ResultSetCols.type = SQL_CHAR
        SELF.ResultSetCols.targetAddr = ADDRESS( SELF.tagChar )
        SELF.ResultSetCols.targetSize = SELF.ResultSetCols.size
      END

      ADD( SELF.ResultSetCols )
    END
  END

  FREE( SELF.ResultSet )                                    ! Wipes out the previous result set
  LOOP
    retCode = SQLFetch( SELF.hstmt )                        ! Read the next row in result set
    IF NOT INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )
      BREAK
    END
    rows += 1                                               ! A sure way to count how many rows
    LOOP i# = 1 TO cols                                     ! Lets process each column
      GET( SELF.ResultSetCols, i# )

      retCode = SQLGetData( SELF.hstmt, i#, SELF.ResultSetCols.type, |
        SELF.ResultSetCols.targetAddr, SELF.ResultSetCols.targetSize, ADDRESS( lLength ) )

      IF lLength < 1                                        ! Nothing to read
        CYCLE                                               ! don't even waste time
      END

      GET( SELF.ResultSet, 0 )
      CLEAR( SELF.ResultSet )
      SELF.ResultSet.row = rows
      SELF.ResultSet.column = i#
      SELF.ResultSet.length = lLength
      IF INRANGE( retCode, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO )
        SELF.ResultSet.value = SELF.ReadMem( lLength )
  OMIT( '//', _USE_ASTRINGS_ )
        SELF.ResultSet.search = SELF.ResultSet.value
  //
        ADD( SELF.ResultSet )
      ELSE
        ODBCGetDiag( SQL_HANDLE_STMT, SELF.hstmt, SELF.Connection, 'LoadResult( ' & SELF.query & ' )' )
      END
    END
  END
  SELF.ResultSetRowCount = rows                             ! Update the cell counters
  SELF.ResultSetColCount = cols

ODBCDataSet.LoadList    PROCEDURE( LONG feq )
colCaption              CSTRING( 64 )
colWidth                LONG                                ! Listbox column width
strFormat               CSTRING( 4096 )
colPicture              CSTRING( 127 )                      ! Listbox column picture token
colToken                CSTRING( 32 )

  CODE
  IF feq{ PROP:Type } <> CREATE:List THEN RETURN END        ! Avoid wasting time and jeopardizing the code if control is no a listbox
  IF NOT RECORDS( SELF.ResultSet )
    RETURN
  END

  IF SELF.listFormat                                        ! If there's already a format
    feq{ PROP:Format } = SELF.listFormat                    ! Apply the stored format to listbox
    RETURN                                                  ! and return
  END

  SELF.BindDefaultObject()                                  ! Make sure there's one ODBC object bound
  LOOP i# = 1 TO SELF.ResultSetColCount
    GET( SELF.ResultSetCols, i# )

    colCaption = SELF.ResultSetCols.name
    colWidth = SELF.ResultSetCols.size
    IF colWidth > 200 THEN colWidth = 200 END
    colToken = ''

    SELF.listColumns.colName = SELF.ResultSetCols.key       ! Prepare to locate a matching column, by KEY
    GET( SELF.listColumns, SELF.listColumns.colName )
    IF NOT ERRORCODE()                                      ! Found an entry
      IF NOT SELF.listColumns.visible THEN CYCLE END        ! If not visible, skipt format preparation
      IF SELF.listColumns.width
        colWidth = SELF.listColumns.width
      END
      IF SELF.listColumns.caption
        colCaption = SELF.listColumns.caption
      END
      IF SELF.listColumns.picture
        colToken = SELF.listColumns.picture                 ! Sets or clears the picture token (see below)
        IF colToken THEN colToken = colToken & '@' END
      END

    ELSE                                                    ! Result set column was not found
      IF RECORDS( SELF.listColumns )                        ! But is there an override columns collections?
        CYCLE                                               ! if so, then this read column can't go to the listbox
      END

    END

    IF NOT colWidth                                         ! No suitable width could be determined, lets calculate it
      colWidth = CHOOSE( INT( LOG10( colWidth * 10 ) ), 30, 60, 100, 150, 200 )
    END

    CASE SELF.ResultSetCols.type
    OF SQL_CHAR OROF SQL_VARCHAR
      IF NOT colToken THEN colToken = '@S' & colWidth & '@' END
      colPicture = 'L(2)|M~'

    OF SQL_DECIMAL OROF SQL_NUMERIC OROF SQL_SMALLINT OROF SQL_INTEGER OROF SQL_BIGINT OROF SQL_REAL
      IF NOT colToken THEN colToken = '@N' & colWidth & '.' END
      IF SELF.ResultSetCols.decimals THEN colToken = colToken & '`' & SELF.ResultSetCols.decimals END
      colToken = colToken & '@'
      colPicture = 'R(2)|M~'

    OF SQL_TYPE_DATE OROF SQL_INTERVAL_DAY
      IF NOT colToken THEN colToken = '@S12@' END
      colPicture = 'C|M~'

    OF SQL_TYPE_TIMESTAMP OROF SQL_TYPE_TIME
      IF NOT colToken THEN colToken = '@S19@' END
      colPicture = 'L(2)|M~'

    END

    colPicture = colPicture & colCaption & '~' & colToken
    strFormat = strFormat & colWidth & colPicture
  END
  Feq{ PROP:Format } = strFormat

ODBCDataSet.Reset       PROCEDURE
  CODE
  OMIT( '//', _USE_ASTRINGS_ )
  LOOP i# = 1 TO RECORDS( SELF.ResultSet )
    GET( SELF.ResultSet, i# )
    SELF.ResultSet.value &= NULL
    PUT( SELF.ResultSet )
  END
  //
  FREE( SELF.ResultSet )
  SELF.ResultSetChanges = 0
  LOOP i# = 1 TO RECORDS( SELF.Dependent )
    GET( SELF.Dependent, i# )
    SELF.Dependent.O.Sync()
  END

ODBCDataSet.Go          PROCEDURE( LONG direction=-1 )
rowTarget               LONG

  CODE
  IF NOT SELF.ResultSetRowCount THEN RETURN 0 END
  IF direction = 0 THEN RETURN 0 END
  IF direction < GO:Last THEN RETURN 0 END
  IF direction > SELF.ResultSetRowCount THEN RETURN 0 END
  CASE direction
  OF GO:First
    rowTarget = 1

  OF GO:Last
    GET( SELF.ResultSet, RECORDS( SELF.ResultSet ) )
    rowTarget = SELF.ResultSet.row

  ELSE
    rowTarget = direction

  END
  SELF.CurrentRow = rowTarget
  SELF.ResultSet.row = rowTarget
  SELF.ResultSet.column = 1
  GET( SELF.ResultSet, SELF.ResultSet.row, SELF.ResultSet.column )
  LOOP i# = 1 TO RECORDS( SELF.Vars )
    GET( SELF.Vars, i# )
    SELF.Vars.var = SELF.Get( SELF.Vars.id )
  END
  RETURN SELF.CurrentRow

ODBCDataSet.Previous    PROCEDURE
  CODE
  IF NOT SELF.ResultSetRowCount THEN RETURN GO:Last END
  IF SELF.CurrentRow = 1 THEN RETURN GO:First END
  SELF.Go( SELF.CurrentRow - 1 )
  RETURN 0

ODBCDataSet.Next        PROCEDURE
  CODE
  IF NOT SELF.Go( SELF.CurrentRow + 1 )
    RETURN GO:Last
  END
  RETURN 0

ODBCDataSet.Fetch       PROCEDURE( STRING column, STRING value )
col                     LONG

  CODE
  IF NOT SELF.ResultSetRowCount THEN RETURN 1 END           ! No rows

  SELF.ResultSetCols.key = UPPER( column )
  GET( SELF.ResultSetCols, SELF.ResultSetCols.key )
  IF ERRORCODE() THEN RETURN 2 END                          ! Column not found
  col = POINTER( SELF.ResultSetCols )                       ! Save the column number

  SELF.ResultSet.column = col
  COMPILE( '//', _USE_ASTRINGS_ )
  SELF.ResultSet.value = value
  GET( SELF.ResultSet, SELF.ResultSet.column, SELF.ResultSet.value )
  //
  OMIT( '//', _USE_ASTRINGS_ )
  SELF.ResultSet.search = value
  GET( SELF.ResultSet, SELF.ResultSet.column, SELF.ResultSet.search )
  //
  IF ERRORCODE() THEN RETURN 3 END

  IF SELF.listBox                                           ! Is there an associated listbox:
    SELF.listBox{ PROP:Selected } = SELF.ResultSet.row      ! Selects listbox row
  END
  RETURN 0                                                  ! If reached here, record found

ODBCDataSet.Get         PROCEDURE( STRING column )
  CODE
  IF NOT SELF.ResultSetRowCount THEN RETURN '' END
  IF NOT POINTER( SELF.ResultSet )
    GET( SELF.ResultSet, 1 )
  END
  SELF.ResultSetCols.key = UPPER( column )
  GET( SELF.ResultSetCols, SELF.ResultSetCols.key )
  IF ERRORCODE() THEN RETURN '' END
  RETURN SELF.GetCell( SELF.ResultSet.row, POINTER( SELF.ResultSetCols ) )

ODBCDataSet.GetCurrentRow PROCEDURE
  CODE
  IF NOT SELF.ResultSetRowCount THEN RETURN '' END
  RETURN SELF.ResultSet.row

ODBCDataSet.GetCell     PROCEDURE( LONG row, SHORT column )
lChanges                LONG
  CODE
  CASE row
  OF -1                                                     ! Internal ask how many rows
    RETURN( SELF.ResultSetRowCount )

  OF -2                                                     ! Internal ask how many columns
    RETURN( SELF.ResultSetColCount )

  OF -3                                                     ! Internal ask if any changes
    lChanges = CHANGES( SELF.ResultSet )
    IF lChanges <> SELF.ResultSetChanges
      SELF.ResultSetChanges = lChanges
      RETURN 1
    ELSE
      RETURN 0
    END

  ELSE
    IF SELF.TraceGetCell
      CLEAR( SELF.GetCellHistory )
      SELF.GetCellHistory.reference = SELF.dataSetName
      SELF.GetCellHistory.row = row
      SELF.GetCellHistory.col = column
    END

    SELF.ResultSet.row = row
    SELF.ResultSet.column = column
    GET( SELF.ResultSet, SELF.ResultSet.row, SELF.ResultSet.column )
    IF ERRORCODE()
      IF SELF.TraceGetCell
        SELF.GetCellHistory.value = 'NOT FOUND'
      END
      CLEAR( SELF.ResultSet )
    ELSE
      IF SELF.TraceGetCell
        SELF.GetCellHistory.value = SELF.ResultSet.value
      END
    END

    IF SELF.TraceGetCell
      ADD( SELF.GetCellHistory )
    END
    RETURN( CLIP( SELF.ResultSet.value ) )
  END

ODBCDataSet.Construct     PROCEDURE
  CODE
  SELF.ResultSetCols &= NEW ODBCColumnsType
  SELF.listColumns &= NEW ODBCListFormat
  SELF.ResultSet &= NEW ODBCRowsType
  SELF.Filters &= NEW ODBCRelatedFilters
  SELF.FilterVars &= NEW ODBCRelatedVars
  SELF.Vars &= NEW ODBCRelatedVars
  SELF.Dependent &= NEW ODBCDependentType
  SELF.GetCellHistory &= NEW ODBCGetCellHistory
  SELF.dataSetName = 'Data set ' & ADDRESS( SELF )

ODBCDataSet.Destruct      PROCEDURE
  CODE
  DISPOSE( SELF.ResultSetCols )
  DISPOSE( SELF.ResultSet )
  DISPOSE( SELF.Filters )
  DISPOSE( SELF.listColumns )
  DISPOSE( SELF.GetCellHistory )

  LOOP i# = 1 TO RECORDS( SELF.FilterVars )
    GET( SELF.FilterVars, i# )
    SELF.FilterVars.var &= NULL
    PUT( SELF.FilterVars )
  END
  FREE( SELF.FilterVars )
  DISPOSE( SELF.FilterVars )

  LOOP i# = 1 TO RECORDS( SELF.Vars )
    GET( SELF.Vars, i# )
    SELF.Vars.var &= NULL
    PUT( SELF.Vars )
  END
  FREE( SELF.Vars )
  DISPOSE( SELF.Vars )

  LOOP i# = 1 TO RECORDS( SELF.Dependent )
    GET( SELF.Dependent, i# )
    SELF.Dependent.O &= NULL
  END
  DISPOSE( SELF.Dependent )

  SELF.ParentDS &= NULL

ODBCDataSet.ReadMem     PROCEDURE( LONG pLength )
ReturnValue             BSTRING
PeekBuffer              &STRING

  CODE
  CASE SELF.ResultSetCols.type
  OF SQL_CHAR OROF SQL_VARCHAR OROF SQL_LONGVARCHAR         ! Atenção, por que a variável pode variar no comprimento
    PeekBuffer &= NEW STRING( pLength + 1 )
    PEEK( SELF.ResultSetCols.targetAddr, PeekBuffer )
    ReturnValue = PeekBuffer
    DISPOSE( PeekBuffer )

  OF SQL_GUID
    PEEK( SELF.ResultSetCols.targetAddr, SELF.tagSQLGUID )
    ReturnValue = SELF.tagSQLGUID

  OF SQL_INTEGER
    PEEK( SELF.ResultSetCols.targetAddr, SELF.tagINTEGER )
    ReturnValue = SELF.tagINTEGER

  OF SQL_SMALLINT
    PEEK( SELF.ResultSetCols.targetAddr, SELF.tagSMALLINT )
    ReturnValue = SELF.tagSMALLINT

  OF SQL_DECIMAL OROF SQL_NUMERIC
    PEEK( SELF.ResultSetCols.targetAddr, SELF.tagSQL_NUMERIC_STRUCT )
    IF NOT SELF.tagSQL_NUMERIC_STRUCT THEN ReturnValue = '-' END
    ReturnValue = ReturnValue & SELF.tagSQL_NUMERIC_STRUCT.value & |
      ',' & SELF.tagSQL_NUMERIC_STRUCT.precision & |
      ',' & SELF.tagSQL_NUMERIC_STRUCT.scale

  OF SQL_BIGINT OROF SQL_REAL
    PEEK( SELF.ResultSetCols.targetAddr, SELF.tagBIGINT_STRUCT )
    ReturnValue = SELF.tagBIGINT_STRUCT

  OF SQL_TYPE_DATE OROF SQL_INTERVAL_DAY
    PEEK( SELF.ResultSetCols.targetAddr, SELF.tagDATE_STRUCT )
    ReturnValue = FORMAT( SELF.tagDATE_STRUCT.day, @N02 ) & '/' & |
      FORMAT( SELF.tagDATE_STRUCT.month, @N02 ) & '/' & SELF.tagDATE_STRUCT.year

  OF SQL_TYPE_TIMESTAMP
    PEEK( SELF.ResultSetCols.targetAddr, SELF.tagTIMESTAMP_STRUCT )
    ReturnValue = FORMAT( SELF.tagTIMESTAMP_STRUCT.day, @n02 ) & '/' & |
      FORMAT( SELF.tagTIMESTAMP_STRUCT.month, @n02 ) & '/' & |
      SELF.tagTIMESTAMP_STRUCT.year & ' ' & |
      FORMAT( SELF.tagTIMESTAMP_STRUCT.hour, @N02 ) & ':' & |
      FORMAT( SELF.tagTIMESTAMP_STRUCT.minute, @n02 ) & ':' & |
      FORMAT( SELF.tagTIMESTAMP_STRUCT.second, @n02 )

  OF SQL_TYPE_TIME
    PEEK( SELF.ResultSetCols.targetAddr, SELF.tagTIME_STRUCT )
    ReturnValue = FORMAT( SELF.tagTIME_STRUCT.hour, @N02 ) & ':' & |
      FORMAT( SELF.tagTIME_STRUCT.minute, @n02 ) & ':' & |
      FORMAT( SELF.tagTIME_STRUCT.second, @n02 )
      ReturnValue = DEFORMAT( ReturnValue, @t04 )
  END
  RETURN ReturnValue

ODBCDataSet.ExportCSV   PROCEDURE
row                     LONG                                          ! Current row counter
col                     LONG                                          ! Current col counter
fileName                CSTRING( 256 )

ODBCCSVFile             FILE, DRIVER( 'ASCII' ), CREATE, THREAD
RECORD                    RECORD
l                           STRING( 65520 )
                          END
                        END

  CODE
  IF NOT SELF.ResultSetRowCount
    MESSAGE( 'Nada para ser copiado', 'Copiar', ICON:Exclamation )
    RETURN
  END
  IF NOT FILEDIALOG( 'Informe o nome do arquivo', fileName, 'Arquivos CSV|*.CSV|Todos os Arquivos|*.*', |
    FILE:Save + FILE:LongName + FILE:KeepDir + FILE:AddExtension )
    RETURN
  END
  ODBCCSVFile{ PROP:Name } = fileName
  CREATE( ODBCCSVFile )
  IF ERRORCODE()
    MESSAGE( 'Erro ao criar o arquivo: ' & ERROR(), 'Exportar CSV', ICON:Hand )
    RETURN
  END
  OPEN( ODBCCSVFile, 42H )
  IF ERRORCODE()
    MESSAGE( 'Erro ao abrir o arquivo: ' & ERROR(), 'Exportar CSV', ICON:Hand )
    RETURN
  END

  OPEN( wProcessing )
  0{ PROP:Text } = 'Exportando CSV'
  ?csvProgress{ PROP:RangeHigh } = SELF.ResultSetRowCount             ! Ajusta a largura da barra de progresso, com a quantidade de registros
  DISPLAY()
  STREAM( ODBCCSVFile )                                               ! Arma escritas em sequência no disco, sem atualização do diretório

  LOOP i# = 1 TO SELF.ResultSetColCount                               ! Monta a linha de cabeçalho
    GET( SELF.ResultSetCols, i# )
    IF i# > 1
      ODBCCSVFile.l = CLIP( ODBCCSVFile.l ) & ';'
    END
    ODBCCSVFile.l = CLIP( ODBCCSVFile.l ) & CLIP( SELF.ResultSetCols.name )
  END
  ADD( ODBCCSVFile )
  LOOP row = 1 TO SELF.ResultSetRowCount                              ! Percorre todas as linhas
    SELF.Go( row )                                                    ! Lê a próxima linha
    CLEAR( ODBCCSVFile.RECORD )                                       ! Limpa o registro do arquivo
    IF not row % 100                                                  ! A cada 100 iterações
      ?csvProgress{ PROP:Progress } = row                             ! Atualiza a barra de progresso
      DISPLAY()                                                       ! e apresenta a janela
    END
    LOOP col = 1 TO SELF.ResultSetColCount                            ! Estando em uma linha, percorre todas as colunas
      IF col > 1                                                      ! Já lemos ao menos uma coluna?
        ODBCCSVFile.l = CLIP( ODBCCSVFile.l ) & ';'                   ! Vamos colocar o separador de colunas
      END
      ODBCCSVFile.l = CLIP( ODBCCSVFile.l ) & SELF.GetCell( row, col ) ! Compõe a linha
    END
    ADD( ODBCCSVFile )                                                ! Ao terminar de processar as colunas, inclui o registro
  END
  CLOSE( ODBCCSVFile )
  CLOSE( wProcessing )
  BEEP( BEEP:SystemExclamation )
  MESSAGE( row & ' registros exportados para ' & fileName, 'Exportar CSV', ICON:Exclamation )

ODBCDataSet.CopyClipboard PROCEDURE
strClipboard            BSTRING                                       ! This will be put in the clipboard
row                     LONG                                          ! Current row counter
col                     LONG                                          ! Current col counter
rows                    LONG                                          ! How many rows are there to process
cols                    LONG                                          ! How many columns are there to process
registro                BSTRING

  CODE
  IF NOT SELF.ResultSetRowCount
    MESSAGE( 'Nada para ser copiado', 'Copiar', ICON:Exclamation )
    RETURN
  END
  LOOP WHILE KEYBOARD()
    ASK
  END
  OPEN( wProcessing )
  0{ PROP:Text } = 'Copiando para o clipboard'
  ?csvProgress{ PROP:RangeHigh } = SELF.ResultSetRowCount             ! Ajusta a largura da barra de progresso, com a quantidade de registros
  DISPLAY()
  LOOP i# = 1 TO SELF.ResultSetColCount                               ! Monta a linha de cabeçalho
    GET( SELF.ResultSetCols, i# )
    IF i# > 1
      strClipboard = CLIP( strClipboard ) & '<9>'
    END
    strClipboard =  CLIP( strClipboard ) & CLIP( SELF.ResultSetCols.name )
  END
  strClipboard = strClipboard & '<13,10>'
  LOOP row = 1 TO SELF.ResultSetRowCount                              ! Percorre todas as linhas
    SELF.Go( row )                                                    ! Lê a próxima linha
    IF not row % 100                                                  ! A cada 100 iterações
      IF KEYBOARD()                                                   ! Anything in the keyboard?
        ASK                                                           ! Read the buffer
        IF KEYCODE() = EscKey                                         ! User wants to cancel the process
          strClipboard = ''                                           ! Clears the buffer, signaling the cancel
          BREAK                                                       ! Exits the loop
        END
      END
      ?csvProgress{ PROP:Progress } = row                             ! Atualiza a barra de progresso
      DISPLAY()                                                       ! e apresenta a janela
    END
    LOOP col = 1 TO SELF.ResultSetColCount                            ! Estando em uma linha, percorre todas as colunas
      IF col > 1                                                      ! Já lemos ao menos uma coluna?
        strClipboard = CLIP( strClipboard ) & '<9>'                   ! Vamos colocar o separador de colunas
      END
      strClipboard = CLIP( strClipboard ) & SELF.GetCell( row, col )  ! Compõe a linha
    END
    strClipboard = CLIP( strClipboard ) & '<13,10>'
  END
  CLOSE( wProcessing )
  IF strClipboard                                                     ! User didn't stop the process
    SETCLIPBOARD( strClipboard )                                      ! Transfer string to clipboard
  END

ODBCDataSet.UI          PROCEDURE
wdw_odbcUI              WINDOW('Propriedades de DataSet'),AT(,,492,207),|
                        FONT('Microsoft Sans',8,,),CENTER,IMM,GRAY, RESIZE
                          PROMPT('&Consulta:'),AT(5,5),USE(?promptConsulta),TRN
                          TEXT,AT(5,16,271,152),USE(SELF.query,, ?SELFQUERY),VSCROLL,FONT('Courier New',10,,)
                          BUTTON('&Ok'),AT(392,184,45,15),USE(?btn_odbcUI_ok),DEFAULT
                          BUTTON('&Cancelar'),AT(440,184,45,15),USE(?btn_odbcUI_cancel)
     END
ExecuteQuery            LONG

  CODE
  OPEN( wdw_odbcUI )
  0{ PROP:Text } = SELF.dataSetName
  ACCEPT
    DISPLAY()
    CASE EVENT()
    OF EVENT:Sized
!      SETPOSITION( ?btn_odbcUI_ok, 0{ PROP:XPos } - 80, 0{ PROP:Ypos } - 20 )
!      SETPOSITION( ?btn_odbcUI_cancel, 0{ PROP:XPos } - 45, 0{ PROP:Ypos } - 20 )
!      SETPOSITION( ?SELFQUERY, 5, 5, 0{ PROP:Width } - 10, 0{ PROP:Height } - 35 )

    OF EVENT:Accepted
      CASE FIELD()
      OF ?btn_odbcUI_ok
        ExecuteQuery = TRUE
        BREAK

      OF ?btn_odbcUI_cancel
        BREAK

      END
    END
  END
  CLOSE( wdw_odbcUI )
  IF ExecuteQuery AND SELF.query
    SELF.Exec( SELF.query, SELF.listBox )
    SELF.Sync()
  END

ODBCDataSet.HandleClicks  PROCEDURE
  CODE
  CASE EVENT()
  OF EVENT:NewSelection
    IF SELF.GetCell( CHOICE( FIELD() ), 1 ) END
    LOOP i# = 1 TO RECORDS( SELF.Dependent )
      GET( SELF.Dependent, i# )
      SELF.Dependent.O.Sync()
    END

  OF EVENT:AlertKey
    CASE KEYCODE()
    OF CtrlMouseRight
      IF FIELD() = SELF.listBox
        SELF.UI()
      END

    COMPILE( '//', _DebugQueuePresent_ )
    OF SELF.debugKey
      phDebugQueue( SELF.ResultSetCols )
      phDebugQueue( SELF.ResultSet )
    //
    END
  END
  RETURN 0

ODBCDataSet.SyncDependents PROCEDURE
  CODE
  LOOP i# = 1 TO RECORDS( SELF.Dependent )
    GET( SELF.Dependent, i# )
    SELF.Dependent.O.Sync()
  END

ODBCDataSet.Sync        PROCEDURE
pos                     LONG
endPos                  LONG
i                       LONG
shift                   LONG                                ! Total of chars to shift
bParentDS               LONG                                ! Has a parent data set?
sql                     CSTRING( 8192 )
szValue                 CSTRING( 256 )

  CODE
  pos = INSTRING( ' WHERE ', UPPER( SELF.query ), 1, 1 )    ! Is there a WHERE clause in the query?
  IF NOT pos THEN RETURN END                                ! Ignore if no where clause at SQL query

  IF NOT SELF.ParentDS &= NULL                              ! Checks if parent data set exists
    bParentDS = TRUE                                        ! Flag the presence of a parent data set
    SELF.ParentDS.Dependent.id = ADDRESS( SELF )            ! Prepare to access parent's filter register
    GET( SELF.ParentDS.Dependent, SELF.ParentDS.Dependent.id )
    IF ERRORCODE() THEN RETURN END                          ! Ignore if not found (this is an error, should never happen)
  END

  SELF.ResultSetChanges = 0                                 ! Signals a change in the result set
  sql = SELF.query                                          ! Copies the original query so we can update it

  LOOP                                                      ! Lets process the filter clause
    pos = INSTRING( '?', sql, 1, pos )                      ! Where is the next ? param
    IF pos                                                  ! If found a ?
      i += 1                                                ! Increment the param counter
      GET( SELF.Filters, i )                                ! and get the field register
      IF ERRORCODE() THEN BREAK END
      IF bParentDS                                          ! Only access a supposed existing parent data set if exists
        szValue = SELF.ParentDS.Get( SELF.Filters.field )
        shift += LEN( CLIP( szValue ) )                     ! Adds the shift chars
        sql = sql[ 1 : pos - 1 ] & szValue & sql[ pos + 1 : LEN( sql ) ]
      END
      pos += 1
    ELSE                                                    ! No -more- ? param found
      BREAK
    END
  END

  LOOP i# = 1 TO RECORDS( SELF.FilterVars )                 ! Lets process any variable filters
    GET( SELF.FilterVars, i# )
    pos = INSTRING( CLIP( SELF.FilterVars.id ), LOWER( sql ), 1, 1 ) ! Is there a variable filter name in the SQL?
    IF sql[ pos - 1 ] = '@'                                 ! Found a beginning marker
      endPos = INSTRING( '@', sql, 1, pos + 1 )             ! Is there an "ending" marker?
      IF NOT endPos THEN CYCLE END                          ! No, didn't find another @, lets recycle (probably in vain)
      IF SUB( LOWER( sql ), pos, endPos - pos ) = CLIP( SELF.FilterVars.id ) ! Found the variable name between marks!
        sql = sql[ 1 : pos - 2 ] & SELF.FilterVars.var & sql[ endPos + 1 : LEN( sql ) ]
      END
    END
  END

  SELF.Exec( sql, SELF.listBox )                            ! Update dataset's list control
  IF SELF.GetCell( 1, 1 ) THEN pos = 1 ELSE pos = 2 END     ! Are there any results?
  LOOP i# = 1 TO RECORDS( SELF.Dependent )                  ! Loops trough all dependent objects
    GET( SELF.Dependent, i# )
    EXECUTE pos                                             ! Set just before the loop (IF SELF.GetCell...)
      SELF.Dependent.O.Sync()                               ! and asks for them to synchronize
      SELF.Dependent.O.Reset()
    END
  END

ODBCDataSet.AddParent   PROCEDURE( ODBCDataSet parentObject, STRING filter )
O                       &ODBCDataSet

  CODE
  O &= parentObject
  SELF.ParentDS &= parentObject
  SELF.ParentDS.Dependent.O &= SELF
  SELF.ParentDS.Dependent.id = ADDRESS( SELF )
  ADD( SELF.ParentDS.Dependent )
  FREE( SELF.Filters )
  SELF.AddFilter( filter )
  REGISTER( EVENT:NewSelection, ADDRESS( O.HandleClicks ), ADDRESS( O ), , O.listBox )
  O &= NULL                                                 ! Release handle

ODBCDataSet.AddFilter   PROCEDURE( STRING filter )
lStart                  LONG( 1 )
lEnd                    LONG
  CODE
  LOOP
    lEnd = INSTRING( ',', filter, 1, lStart )
    IF NOT lEnd THEN lEnd = LEN( filter ) ELSE lEnd -= 1 END
    SELF.Filters.field = LEFT( CLIP( filter[ lStart : lEnd ] ) )
    ADD( SELF.Filters )
    lStart = lEnd + 2
    IF lStart >= LEN( filter ) THEN BREAK END
  END

ODBCDataSet.AddFilter   PROCEDURE( *? var, STRING varName )
  CODE
  SELF.FilterVars.id = LOWER( varName )
  GET( SELF.FilterVars, SELF.FilterVars.id )
  IF NOT ERRORCODE() THEN RETURN END
  CLEAR( SELF.FilterVars )
  SELF.FilterVars.id = LOWER( varName )
  SELF.FilterVars.var &= var
  SELF.FilterVars.addr = ADDRESS( var )
  ADD( SELF.FilterVars )

ODBCDataSet.AddVar      PROCEDURE( *? var, STRING varName )
  CODE
  SELF.Vars.id = LOWER( varName )
  GET( SELF.Vars, SELF.Vars.id )
  IF NOT ERRORCODE() THEN RETURN END
  CLEAR( SELF.Vars )
  SELF.Vars.id = LOWER( varName )
  SELF.Vars.var &= var
  SELF.Vars.addr = ADDRESS( var )
  ADD( SELF.Vars )

ODBCDataSet.AddResetControl PROCEDURE( LONG feq )
  CODE
  IF NOT feq THEN RETURN END
  REGISTER( EVENT:Accepted, ADDRESS( SELF.Refresh ), ADDRESS( SELF ), , feq )

ODBCDataSet.Refresh     PROCEDURE
  CODE
  UPDATE( FIELD() )
  SELF.Sync()
  RETURN 0

ODBCDataSet.SetDebugKey PROCEDURE( LONG feq )
  CODE
  COMPILE( '//', _DebugQueuePresent_ )
  SELF.DebugKey = feq
  0{ PROP:Alrt, 255 } = feq                                 ! Alerts for requested debug fire up key
  REGISTER( EVENT:AlertKey, ADDRESS( SELF.HandleClicks ), ADDRESS( SELF ) )
  //

ODBCDataSet.AddListColumn PROCEDURE( STRING pColumn, STRING pCaption, STRING pPicture, LONG pWidth=0 )
  CODE
  SELF.listColumns.colNumber = RECORDS( SELF.listColumns )
  SELF.listColumns.colName = UPPER( pColumn )
  SELF.listColumns.caption = pCaption
  SELF.listColumns.width = pWidth
  SELF.listColumns.picture = pPicture
  SELF.listColumns.visible = TRUE
  ADD( SELF.listColumns )

ODBCGetDiagnostic       PROCEDURE( ODBCClass O )
  CODE
  ODBCGetDiag( SQL_HANDLE_ENV, O.hstmt, O )

ODBCSQLState            PROCEDURE( ODBCClass O )
diagSQLState            CSTRING( 6 )                        ! Diagnostic variables
diagNativeError         LONG
diagNativertext         CSTRING( 256 )
diagMessageText         CSTRING( 1024 )
diagTextLength          LONG
odbcObj                 &ODBCClass

  CODE
  odbcObj &= O
  SQLGetDiagRec( SQL_HANDLE_DBC, O.hdbc, 1, diagSQLState, |
    diagNativeError, diagMessageText, SIZE( diagMessageText ), diagTextLength )
  odbcObj &= NULL
  RETURN diagSQLState

ODBCGetDiag             PROCEDURE( LONG handleType, LONG handle, ODBCClass O, <STRING info> )
lRet                    LONG
diagSQLState            CSTRING( 6 )                        ! Diagnostic variables
diagNativeError         LONG
diagNativertext         CSTRING( 256 )
diagMessageText         CSTRING( 1024 )
diagTextLength          LONG

  CODE

  ODBCGetDiagCalls.handleType = handleType
  ODBCGetDiagCalls.handle = handle
  GET( ODBCGetDiagCalls, ODBCGetDiagCalls.handleType, ODBCGetDiagCalls.handle )
  IF ERRORCODE()
    ODBCGetDiagCalls.status = 0
    ADD( ODBCGetDiagCalls )
  END
  IF ODBCGetDiagCalls.status = 2 THEN RETURN END            ! Don't present the window

  SQLGetDiagRec( handleType, handle, 1, diagSQLState, |
    diagNativeError, diagMessageText, SIZE( diagMessageText ), diagTextLength )

  IF diagNativeError = 'HY000'
    RETURN
  END

  lRet = MESSAGE( |
    'henv:<9,9>' & O.henv & '<13,10>' & |
    'hdbc:<9,9>' & O.hdbc & '<13,10>' & |
    'Native code:<9>' & diagNativeError & '<13,10>' & |
    'Error:<9,9>' & diagMessageText & '<13,10>' & |
    'SQLState:<9>' & diagSQLState & '<13,10,13,10>' & |
    info , |
    'ODBCGetDiag( ' & ADDRESS( O ) & ' )', |
    ICON:Exclamation, '&Continuar|&Parar|&Encerrar', 1, MSGMODE:CanCopy )

  IF lRet = 3 THEN HALT( 0 ) END
  IF lRet = 2
    ODBCGetDiagCalls.status = 2
    PUT( ODBCGetDiagCalls )
  END

