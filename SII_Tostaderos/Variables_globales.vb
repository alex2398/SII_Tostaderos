Module Variables_globales
    'Public odbc_cadena As String = "Driver={Microsoft ODBC for Oracle};CONNECTSTRING=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.0.221)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=V153)));Uid=system;Pwd=manager;"
    'Public odbc_cadena As String = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.0.221)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=V153)));Uid=system;Pwd=manager;"
    'Public odbc_cadena As String = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(CID=GTU_APP)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.0.221)(PORT=1521)))(CONNECT_DATA=(SID=V153)(SERVER=DEDICATED)));User Id=system;Password=manager;"
    'Public odbc_cadena As String = "Provider=MSDAORA;Data Source=TOSTADEROS;Persist Security Info=True;Password=manager;User ID=system"
    Public oledb_cadena As String = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(CID=GTU_APP)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.0.221)(PORT=1521)))(CONNECT_DATA=(SID=V153)(SERVER=DEDICATED)));User Id=system;Password=manager;"
    Public version As String = "v1.6 (02/08/17)"
    'Public oledb_cadena As String = "Provider=MSDAORA;Data Source=TOSTADEROS;Persist Security Info=True;Password=manager;User ID=system"


End Module
