ODBCClass
=========

Clarion class to access databases without the need to use a DCT

This class allows a programmer to access a database via ODBC without the need to define a table.
Generally, a table needs to be defined prior to access any file, like:

MyTable       FILE, DRIVER( 'ODBC' ), NAME( 'public.myTable' ), OWNER( szConn )
RECORD          RECORD
id                LONG
name              CSTRING( 200 )
                END
              END
              
While this allows for clearer documentation and stronger system stability, it contributes to lower productivity in the sense that it forces the programmer to maintain an updated database dictionary.

I needed much more flexibility, whilst being able to provide good stability. Also, I wanted to be able to fill a grid/listbox with just one line of code. Then I remembered the old Clipper command, BROWSE.

With this class, the programmer can:

  db.Connect( connctionString )
  ds.Init( db )
  ds.Exec( 'SELECT id, name FROM myTable ORDER BY id', ?List )
  
The above sequence will:

1) Connect to a database, creating a connection object (db);
2) Attach a dataset object (ds) to a specified connection (db);
3) Execute, load, and display the result set on a list control;

Please notice that this class is not intended to completelly substitute the need of standard file definition, even though it can. So, use it with caution.

WARNING: 
I DIDN'T PROGRAM THE CLASS TO REJECT DATABASE MAINTENANCE COMMANDS, LIKE CREATE TABLE, DROP TABLE, DROP DATABASE ETC. IF YOU ALLOW THE END USERS TO EMMIT WHATEVER COMMANDS THEY LIKE, YOU'RE RISKING RUINING COMPLETELLY YOUR DATABASE. IT IS UP TO YOU TO BLOCK USERS FROM INSERTING COMMANDS.

