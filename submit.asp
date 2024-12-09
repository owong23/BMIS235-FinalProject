<%
' Retrieve form data from the query string
Dim customerName, customerEmail, product, quantity, shipping
customerName = Request.QueryString("customerName")
customerEmail = Request.QueryString("customerEmail")
product = Request.QueryString("product")
quantity = Request.QueryString("quantity")
shipping = Request.QueryString("shipping")

' Check for required inputs
If customerName = "" Or customerEmail = "" Or product = "" Or quantity = "" Or shipping = "" Then
    Response.Write("<h3>Please complete the form.</h3>")
Else
    ' Define database connection and query
    Dim conn, connString, sql
    connString = "Driver={SQL Server};Server=BMIS235;Database=BMIS235OTW;UID=owong;PWD=otw@zaga27785jd4S;"
    sql = "INSERT INTO Orders (CustomerName, CustomerEmail, Product, Quantity, ShippingMethod) VALUES (?, ?, ?, ?, ?)"
    
    ' Create and open the database connection
    Set conn = Server.CreateObject("ADODB.Connection")
    On Error Resume Next
    conn.Open connString
    If Err.Number <> 0 Then
        Response.Write("<h3>Error connecting to the database: " & Err.Description & "</h3>")
    Else
        ' Use a parameterized query to insert the data
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = conn
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("CustomerName", 200, 1, 50, customerName) ' 200 = adVarChar
        cmd.Parameters.Append cmd.CreateParameter("CustomerEmail", 200, 1, 50, customerEmail)
        cmd.Parameters.Append cmd.CreateParameter("Product", 200, 1, 50, product)
        cmd.Parameters.Append cmd.CreateParameter("Quantity", 3, 1, , CInt(quantity)) ' 3 = adInteger
        cmd.Parameters.Append cmd.CreateParameter("ShippingMethod", 200, 1, 50, shipping)

        On Error Resume Next
        cmd.Execute
        If Err.Number <> 0 Then
            Response.Write("<h3>Error executing query: " & Err.Description & "</h3>")
        Else
            ' Redirect on successful insertion
            Response.Redirect "Thanksforshoppingpage.asp"
        End If
        Set cmd = Nothing
    End If
    conn.Close
    Set conn = Nothing
End If
%>





%>

