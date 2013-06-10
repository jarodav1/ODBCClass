  OMIT( '//', _DebugQueuePresent_ )
_DebugQueuePresent_     EQUATE( 1 )
  //

  MAP
    MODULE( 'debugQueue.clw' )
      phDebugQueue( QUEUE pQ, <STRING pFile>, BYTE pFormat=1, <STRING pFieldSeparator>, <STRING pStringDelimiter>, BYTE pRecNumbers=TRUE ),PASCAL
      phExportQueue( QUEUE pQ, <STRING pFile>, BYTE pFormat=1, <STRING pFieldSeparator>, <STRING pStringDelimiter>, BYTE pRecNumbers=TRUE ),LONG,PROC,PASCAL
    END
  END

  PRAGMA( 'compile(debugQueue.clw)' )
  PRAGMA( 'link(C%V%ASC%X%%L%.LIB)' )