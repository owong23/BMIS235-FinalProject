<%
Dim customerName, customerEmail, product, quantity, shipping
customerName = Request.QueryString("customerName")
customerEmail = Request.QueryString("customerEmail")
product = Request.QueryString("product")
quantity = Request.QueryString("quantity")
shipping = Request.QueryString("shipping")

sql = "INSERT INTO Orders (CustomerName, CustomerEmail, Product, Quantity, ShippingMethod) " & "VALUES ('" & customerName & "', '" & customerEmail & "', '" & product & "', " & quantity & ", '" & shipping & "')"

' Database connection
Dim conn, connStr, sql
connString = "Driver={SQL Server};Server=BMIS235; Database=BMIS235OTW; UID=owong; PWD=otw@zaga27785jd4S;"
Set conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open connString
conn.Execute SQL
If item = "" Or quantity = "" Then
    Response.Write("<h3>Please complete the form.</h3>")
Else
    Response.Redirect "Thanksforshoppingpage.html"
End If
%>
conn.Close




%>

