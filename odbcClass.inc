  MAP

    MODULE( 'odbcClass.clw' )
      ODBCGetDiagnostic( ODBCClass O )
      ODBCGetDiag( LONG handleType, LONG handle, ODBCClass O, <STRING info> )
      ODBCSQLState( ODBCClass O ), STRING
    END

    MODULE( 'ODBC dll' )
      GetDesktopWindow, LONG, PASCAL
      SQLAllocHandle( SHORT, LONG, *LONG ), LONG, PASCAL
      SQLDriverConnect( LONG, LONG, *CSTRING, SHORT, *CSTRING, SHORT, LONG, LONG ), LONG, PASCAL, RAW, NAME( 'SQLDriverConnectA' )
      SQLSetEnvAttr( LONG, LONG, LONG, LONG ), SHORT, PASCAL
      SQLSetConnectAttr( LONG, LONG, LONG, LONG ), PASCAL
      SQLDisconnect( LONG ), SHORT, PASCAL
      SQLDriverConnect( LONG, LONG, *CSTRING, SHORT, *CSTRING, SHORT, *SHORT, USHORT ), LONG, PASCAL, RAW, NAME( 'SQLDriverConnectA' )
      SQLPrepare( LONG, *CSTRING, LONG ), SHORT, PASCAL, RAW, NAME( 'SQLPrepareA' )
      SQLExecDirect( LONG, *CSTRING, LONG ), SHORT, PASCAL, RAW, PROC, NAME( 'SQLExecDirectA' )
      SQLNumResultCols( LONG, *LONG ), SHORT, PASCAL
      SQLDescribeCol( LONG, USHORT, *CSTRING, SHORT, *SHORT, *SHORT, *LONG, *SHORT, *SHORT ), SHORT, PASCAL, RAW, NAME( 'SQLDescribeColA' )
      SQLFetch( LONG ), SHORT, PASCAL
      SQLGetData( LONG, USHORT, SHORT, LONG, LONG, LONG ), SHORT, PASCAL
      SQLFreeHandle( LONG, LONG ), LONG, PASCAL, PROC
      SQLGetDiagRec( SHORT, LONG, SHORT, *CSTRING, *LONG, *CSTRING, SHORT, SHORT ), LONG, PASCAL, RAW, PROC, NAME( 'SQLGetDiagRecA' )
    END
  END

SQL_INVALID_HANDLE      EQUATE( -2 )
SQL_ERROR               EQUATE( -1 )
SQL_SUCCESS             EQUATE( 0 )
SQL_SUCCESS_WITH_INFO   EQUATE( 1 )
SQL_STILL_EXECUTING     EQUATE( 2 )
SQL_NO_DATA             EQUATE( 100 )
SQL_PARAM_DATA_AVAILABLE EQUATE( 101 )

SQL_NO_TOTAL            EQUATE( -4 )
SQL_NULL_DATA           EQUATE( -1 )

SQL_NULL_HANDLE         EQUATE( 0 )
SQL_HANDLE_ENV          EQUATE( 1 )
SQL_HANDLE_DBC          EQUATE( 2 )
SQL_HANDLE_STMT         EQUATE( 3 )
SQL_HANDLE_DESC         EQUATE( 4 )

SQL_INTERVAL_MINUTE_TO_SECOND EQUATE( -92 )
SQL_INTERVAL_HOUR_TO_SECOND EQUATE( -91 )
SQL_INTERVAL_HOUR_TO_MINUTE EQUATE( -90 )
SQL_INTERVAL_DAY_TO_SECOND EQUATE( -89 )
SQL_INTERVAL_DAY_TO_MINUTE EQUATE( -88 )
SQL_INTERVAL_DAY_TO_HOUR EQUATE( -87 )
SQL_INTERVAL_SECOND     EQUATE( -86 )
SQL_INTERVAL_MINUTE     EQUATE( -85 )
SQL_INTERVAL_HOUR       EQUATE( -84 )
SQL_INTERVAL_DAY        EQUATE( -83 )
SQL_INTERVAL_YEAR_TO_MONTH EQUATE( -82 )
SQL_INTERVAL_MONTH      EQUATE( -81 )
SQL_INTERVAL_YEAR       EQUATE( -80 )
SQL_GUID                EQUATE( -11 )
SQL_BIT                 EQUATE( -7 )
SQL_TINYINT             EQUATE( -6 )
SQL_BIGINT              EQUATE( -5 )
SQL_LONGVARBINARY       EQUATE( -4 )
SQL_VARBINARY           EQUATE( -3 )
SQL_BINARY              EQUATE( -2 )
SQL_LONGVARCHAR         EQUATE( -1 )                        ! TEXT
SQL_UNKNOWN_TYPE        EQUATE( 0 )
SQL_CHAR                EQUATE( 1 )
SQL_NUMERIC             EQUATE( 2 )
SQL_DECIMAL             EQUATE( 3 )
SQL_INTEGER             EQUATE( 4 )
SQL_SMALLINT            EQUATE( 5 )
SQL_FLOAT               EQUATE( 6 )
SQL_REAL                EQUATE( 7 )
SQL_DOUBLE              EQUATE( 8 )
SQL_VARCHAR             EQUATE( 12 )
SQL_TYPE_DATE           EQUATE( 91 )
SQL_TYPE_TIME           EQUATE( 92 )
SQL_TYPE_TIMESTAMP      EQUATE( 93 )
SQL_WLONGVARCHAR        EQUATE( 125 )
SQL_WCHAR               EQUATE( 126 )

SQL_IS_INTEGER          EQUATE( -6 )
SQL_NTS                 EQUATE( -3 )
SQL_ATTR_ODBC_VERSION   EQUATE( 200 )
SQL_OV_ODBC2            EQUATE( 2 )
SQL_OV_ODBC3            EQUATE( 3 )
SQL_OV_ODBC3_80         EQUATE( 380 )

SQL_DRIVER_NOPROMPT     EQUATE( 0 )
SQL_DRIVER_PROMPT       EQUATE( 2 )
SQL_LOGIN_TIMEOUT       EQUATE( 103 )

GO:First                EQUATE( -1 )
GO:Last                 EQUATE( -2 )

ODBCColumnsType         QUEUE, TYPE                         ! Columns descriptions of a result set
name                      CSTRING( 64 )
key                       CSTRING( 64 )                     ! This is UPPERed, to maximize finding probability
type                      LONG
size                      LONG
decimals                  SHORT
nullable                  SHORT
targetAddr                LONG                              ! Target structure address
targetSize                LONG                              ! Target sttructure size
                        END

ODBCRowsType            QUEUE, TYPE                         ! Rows that compose a result set
row                       LONG
column                    LONG
length                    LONG                              ! Specific row, col value length
value                     ANY
                        END

ODBCRelatedFilters      QUEUE, TYPE                         ! Filter fields collection
field                     CSTRING( 64 )                     ! field name
                        END

ODBCRelatedVars         QUEUE, TYPE                         ! Variables collection
id                        STRING( 64 )                      ! Variable name - avoids duplicates
var                       ANY
addr                      LONG
                        END

ODBCDependentType       QUEUE, TYPE                         ! Colection of child datasets
id                        LONG                              ! Address of child data set
O                         &ODBCDataSet
                        END

ODBCClass               CLASS, TYPE, MODULE( 'odbcClass.clw' ), LINK( 'odbcClass.clw', _ABCLinkMode_ ), DLL( _ABCDllMode_ )
henv                      LONG, PRIVATE                     ! ENVIRONMENT handle
hdbc                      LONG                              ! CONNECTION handle
hstmt                     LONG, PRIVATE                     ! STATEMENT handle
connRequest               CSTRING( 1024 ), PRIVATE          ! What we requested to connect
connectionString          CSTRING( 1024 ), PRIVATE          ! Complete connection as responded by the driver

Connect                   PROCEDURE( STRING connectString ), LONG
Exec                      PROCEDURE( STRING sqlCommand ), LONG, PROC
GetConnectionString       PROCEDURE, STRING
Disconnect                PROCEDURE, LONG, PROC
Construct                 PROCEDURE
                        END

ODBCListFormat          QUEUE, TYPE                         ! Listbox columns options
colNumber                 LONG                              ! ordinal position in the listbox
colName                   CSTRING( 64 )                     ! map to result set column
caption                   CSTRING( 64 )                     ! Header text of the column
width                     LONG                              ! column width
picture                   CSTRING( 32 )                     ! column picture
visible                   LONG                              ! this column is to be presented in listbox?
                        END

ODBCGetCellHistory      QUEUE, TYPE                         ! GetCell() method call history
reference                 STRING( 64 )
row                       LONG
col                       LONG
value                     STRING( 256 )
                        END

ODBCDataSet             CLASS, TYPE, MODULE( 'odbcClass.clw' ), LINK( 'odbcClass.clw', _ABCLinkMode_ ), DLL( _ABCDllMode_ )
Connection                &ODBCClass                        ! Reference to the ODBC Connection Class
listColumns               &ODBCListFormat                   ! Collection of columns
hdbc                      LONG, PRIVATE                     ! CONNECTION handle
hstmt                     LONG, PRIVATE                     ! STATEMENT handle
dataSetName               CSTRING( 64 ), PROTECTED          ! Humanized name

tagCHAR                   CSTRING( 256 )
tagBIGINT_STRUCT          STRING( 19 )
tagDATE_STRUCT            GROUP                             ! Target structure of type DATE
year                        SHORT
month                       USHORT
day                         USHORT
                          END
tagTIME_STRUCT            GROUP                             ! Target structure of type TIME
hour                        USHORT
minute                      USHORT
second                      USHORT
                          END
tagTIMESTAMP_STRUCT       GROUP                             ! Target structure of type TIMESTAMP
year                        SHORT
month                       USHORT
day                         USHORT
hour                        USHORT
minute                      USHORT
second                      USHORT
fraction                    ULONG
                          END
tagINTEGER                LONG
tagSMALLINT               SHORT
tagSQL_NUMERIC_STRUCT     GROUP                             ! Target structure of type NUMERIC
precision                   STRING( 1 )
scale                       STRING( 1 )
sign                        STRING( 1 )                     ! 1: Positive; 0: Negative
value                       STRING( 30 )
                          END
tagSQLGUID                GROUP                             ! Target structure of type GUID
data1                       STRING( 4 )
data2                       STRING( 2 )
data3                       STRING( 2 )
data4                       STRING( 8 )
                          END

traceGetCell              LONG
query                     CSTRING( 8192 )                   ! Default SQL query
CurrentRow                LONG
ResultSetColCount         LONG
ResultSetRowCount         LONG
ResultSetCols             &ODBCColumnsType
ResultSet                 &ODBCRowsType
ResultSetChanges          LONG                              ! HASH returned by CHANGES()
ParentDS                  &ODBCDataSet                      ! Handle to parent's variables ;)
Filters                   &ODBCRelatedFilters               ! Filter fields in parent object
FilterVars                &ODBCRelatedVars                  ! Variable filters registered
Vars                      &ODBCRelatedVars                  ! Variables to be updated
Dependent                 &ODBCDependentType
GetCellHistory            &ODBCGetCellHistory

listBox                   LONG                              ! Listbox to show result set
listFormat                CSTRING( 1024 )
debugKey                  LONG                              ! DebugRequested ?
Init                      PROCEDURE( ODBCClass Connection, LONG listBox=0, <STRING name> )
SetName                   PROCEDURE( STRING name )
SetQuery                  PROCEDURE( STRING query )
SetListbox                PROCEDURE( LONG feq )
Exec                      PROCEDURE( STRING sqlCommand ), LONG, PROC
Exec                      PROCEDURE( STRING sqlCommand, LONG feq ), LONG, PROC
Load                      PROCEDURE( LONG feq=0 )
LoadResult                PROCEDURE
LoadList                  PROCEDURE( LONG feq )
Reset                     PROCEDURE
Go                        PROCEDURE( LONG direction=-1 ), LONG, PROC
Previous                  PROCEDURE, LONG, PROC
Next                      PROCEDURE, LONG, PROC
Get                       PROCEDURE( STRING column ), STRING
GetCurrentRow             PROCEDURE, LONG
GetCell                   PROCEDURE( LONG row, SHORT column ), STRING
BindDefaultObject         PROCEDURE
CopyClipboard             PROCEDURE
ExportCSV                 PROCEDURE
UI                        PROCEDURE
AddParent                 PROCEDURE( ODBCDataSet parentObject, STRING filter )
AddFilter                 PROCEDURE( STRING filter ), PRIVATE
AddFilter                 PROCEDURE( *? var, STRING varName )
AddVar                    PROCEDURE( *? var, STRING varName )
AddResetControl           PROCEDURE( LONG feq )                                 ! Register control do fire a reset
AddListColumn             PROCEDURE( STRING pColumn, STRING pCaption, STRING pPicture, LONG pWidth=0 )
Construct                 PROCEDURE
Destruct                  PROCEDURE
ReadMem                   PROCEDURE( LONG pLength ), STRING
HandleClicks              PROCEDURE, BYTE
Refresh                   PROCEDURE, BYTE, PROC
Sync                      PROCEDURE
SyncDependents            PROCEDURE
SetDebugKey               PROCEDURE( LONG feq )
                        END

  PRAGMA( 'link(odbc32.lib)' )
