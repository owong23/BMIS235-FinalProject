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
If Err.Number = 0 Then  <a href="productpage.html"</a>
    Response.Write("<h3>Thank you, " & customerName & "! Your order has been placed successfully.</h3>")
Else
    Response.Write("<h3>Error: Unable to process your order. Please try again later.</h3>")
    Response.Write("Error details: " & Err.Description)
End If
conn.Close




%>

