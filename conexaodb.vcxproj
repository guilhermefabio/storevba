Function CONEXAODB()

'CODING FOR ACCESS TO DATABASE

server= <your server databse>
port = <port server database>
db = <database>                                             'DECLARE THE DATABASES VARIABLES
login = <login for acess database>
passw = <passworf for acess database>

Set CON = New ADODB.Connection
 CON.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};    'important you install the corresponding mysql driver
						SERVER= '" & server & "';
						PORT='" & port & "';
						DATABASE='" & db & "';
						UID='" & login & "';
						PASSWORD='" & passw & "';
						OPTION='" & op & "'"
  CON.CursorLocation = adUseClient
    CON.Open

End Function
