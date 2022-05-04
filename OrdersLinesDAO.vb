Imports System.Data.SqlClient
Public Class OrdersLinesDAO

#Region "VARIABLES"

    Public con As New SqlConnection()
    Private orderNumber As String
    Private orderLineNumber As String

#End Region

#Region "FUNCTIONS"
    'Test connection
    Public Function testConnection() As Boolean
        Try
            con.Open()
            con.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    'Insert articles missing
    Sub insertMissingArticles(monoSKU As String, orderNumber As String, orderNumberLine As String)
        Try
            Dim sel As String = "declare @monoSKU varchar(300) = '" & monoSKU & "'
declare @orderNumber varchar(300) = '" & orderNumber & "'
declare @orderNumberLines varchar(300) ='" & orderNumberLine & "'
if not exists(select * from articles_missing where monoSKU = @monoSKU and orderNumber = @orderNumber and orderNumberLines = @orderNumberLines )
begin
insert into articles_missing values (@monoSKU , @orderNumber ,@orderNumberLines )
end"
            con.Open()
            Dim cmd As New SqlCommand(sel, con)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        con.Close()
    End Sub

    'Get configuration
    Public Function getConfig(ByRef importMinInterval As Integer) As Boolean
        getConfig = False
        Try
            Dim sel As String = "Select importsMinutsInterval from master where numYear = 2022"
            Dim da As New SqlDataAdapter(sel, con)
            Dim dt As New DataTable
            da.Fill(dt)
            For Each row In dt.Rows
                importMinInterval = row(0)
            Next
            Return True
        Catch ex As Exception
        End Try
    End Function
    'Insert new register in import Services
    Public Function createImportServices(ByVal path As String) As Integer
        createImportServices = 0
        con.Open()
        Try
            Dim insertQuery As String = "INSERT INTO imports_services (path , start_services , imports_services ) values ('" & path & "',GETDATE (),GETDATE()); select SCOPE_IDENTITY();"
            Dim cmd As New SqlCommand(insertQuery, con)
            createImportServices = cmd.ExecuteScalar()
        Catch ex As Exception
        End Try
        con.Close()
    End Function
    'Update log import services
    Public Sub updateImportServices(id As Integer, text As String)
        con.Open()
        Try
            Dim insertQuery As String = "update imports_services set logText = '" & text & "' where id = " & id
            Dim cmd As New SqlCommand(insertQuery, con)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        con.Close()
    End Sub

    'Test if exists the client.
    Public Function testClient(ByVal idClient As Integer) As Boolean
        testClient = False
        Try
            Dim sel As String = "Select count(*) from clients where ax_reference = '" & idClient & "'"
            Dim cmd As New SqlCommand(sel, con)
            con.Open()
            If cmd.ExecuteScalar > 0 Then
                testClient = True
            End If
        Catch ex As Exception
        End Try
        con.Close()
    End Function
    'Test if exists the article.
    Public Function testArticles(ByVal referenceAX As String, ByVal id_client As Integer) As Integer
        testArticles = 0
        Try

                Dim sel As String = "
declare @id_articles int;
declare @id_clients	int = '" & id_client & "';
declare @axReference varchar(300) = '" & referenceAX & "';
if ((select  COUNT(*) from articles where name = @axReference) = 1)
begin
set @id_articles = (select  id  from articles  where name =@axReference)
end;
if @id_articles is null
begin
set @id_articles = (select id_articles from articles_articles_clients where id_articles in (select id from articles where name = @axReference) and id_clients = @id_clients)
end;
if @id_articles is null
begin
set @id_articles = (select top(1) id  from articles  where name = @axReference and name = monoSKU)
end;
select @id_articles 
"
                Dim cmd As New SqlCommand(sel, con)
                con.Open()
                testArticles = cmd.ExecuteScalar()
            Catch ex As Exception
            End Try
        con.Close()
    End Function
    'Insert new client
    Public Function CreateClient(ByVal idClient As String, ByVal name As String) As Boolean
        CreateClient = False
        Try
            Dim sel As String = "Insert into clients (name, ax_reference, blocked) values ('" & idClient & "', '" & name & "', 0)"
            Dim cmd As New SqlCommand(sel, con)
            con.Open()
            cmd.ExecuteNonQuery()
            CreateClient = True
        Catch ex As Exception
        End Try
        con.Close()
    End Function
    'Insert new client
    Public Function UpdateClient(ByVal idClient As String, ByVal block As Boolean, ByVal name As String) As Boolean
        UpdateClient = False
        Try
            Dim sel As String = "UPDATE [dbo].[clients]
   SET [name] = '" & name & "'
      ,[modify_date] = getdate()
      ,[blocked] = '" & block & "'
 WHERE  ax_reference = '" & idClient & "'"
            Dim cmd As New SqlCommand(sel, con)
            con.Open()
            cmd.ExecuteNonQuery()
            UpdateClient = True
        Catch ex As Exception
        End Try
        con.Close()
    End Function
    'Compare register
    Public Function compareLines(ByVal ol As orderLine) As String
        compareLines = ""
        Dim olA As New orderLine
        olA = existsLine(ol.SOrderAX, ol.SOrderLineAx)
        Dim diferents As String = ""
        If Not olA.IId = 0 Then
            Dim result As Boolean = True
            If Not olA.sIdClient = ol.sIdClient Then
                diferents = " set id_clients = (select id from clients where ax_reference = '" & ol.sIdClient & "') "
                result = False
            End If
            If Not olA.SOrderAX = ol.SOrderAX Then
                diferents = formatDiferents(diferents, "order_number", ol.SOrderAX)
                result = False
            End If
            If Not olA.OrderAXDate = ol.OrderAXDate Then
                diferents = formatDiferents(diferents, "order_date", ol.OrderAXDate)
                result = False
            End If
            If Not olA.SOrderLineAx = ol.SOrderLineAx Then
                diferents = formatDiferents(diferents, "order_line_number", ol.SOrderLineAx)
                result = False
            End If
            If Not olA.OrderLineAXDate = ol.OrderLineAXDate Then
                diferents = formatDiferents(diferents, "order_line_date", ol.OrderLineAXDate)
                result = False
            End If
            If Not olA.DesiredDate = ol.DesiredDate Then
                diferents = formatDiferents(diferents, "desired_date", ol.DesiredDate)
                result = False
            End If
            If Not olA.RequestedDate = ol.RequestedDate Then
                diferents = formatDiferents(diferents, "requested_date", ol.RequestedDate)
                result = False
            End If
            If Not Trim(olA.SNotes) = Trim(ol.SNotes) Then
                diferents = formatDiferents(diferents, "notes", ol.SNotes)
                result = False
            End If
            If Not olA.BClientBlock = ol.BClientBlock Then
                If olA.IId_states <> 2 And olA.BClientBlock <> 3 Then
                    updateLine(" set client_block = '" & ol.BClientBlock & "'", olA.IId)
                End If
                diferents = formatDiferents(diferents, "client_block", ol.BClientBlock)
                result = False
            End If
            If Not olA.IIdArticles = ol.IIdArticles Then
                diferents = formatDiferents(diferents, "id_articles", ol.IIdArticles)
                result = False
            End If
            If Not olA.IQuantity = ol.IQuantity Then
                diferents = formatDiferents(diferents, "total_quantity", ol.IQuantity)
                result = False
            End If
            If Not olA.BIsDeleted = ol.BIsDeleted And result = True Then
                updateLine(" set is_deleted = 0 ", olA.IId)
                Return "Se ha restaurado la línea " & olA.SOrderLineAx & " del pedido " & olA.SOrderAX & ". "
            End If
            If result = False Then
                ol.IId = olA.IId
                If olA.IId_states = 2 Then
                    If updateLine(diferents, ol.IId) Then
                        Return "Se ha actualizado la línea " & olA.SOrderLineAx & " del pedido " & olA.SOrderAX & "."
                    Else
                        Return "*** Error al actualizar la línea " & olA.SOrderLineAx & " del pedido " & olA.SOrderAX & "."
                    End If
                Else
                    Select Case createTemporalOrderLine(ol)
                        Case 0
                            Return "*** Error al crear el registo temporal de la línea " & olA.SOrderLineAx & " del pedido " & olA.SOrderAX & ", pendiente de confirmación."
                        Case 1
                            Return "Se ha creado un registo temporal de la línea " & olA.SOrderLineAx & " del pedido " & olA.SOrderAX & ", pendiente de confirmación."
                        Case 2
                            Return "Se ha actualizado el registo temporal de la línea " & olA.SOrderLineAx & " del pedido " & olA.SOrderAX & ", pendiente de confirmación."
                    End Select
                End If
            Else
                Return "NR"
            End If
        Else
            Return ""
        End If
    End Function
    'Insert temporal line.
    Public Function createTemporalOrderLine(ol As orderLine) As Integer
        createTemporalOrderLine = 0
        Try
            con.Open()
            Dim sel As String = "
Declare @iid int;
set @iid = (select id from orders_lines_temp where [order_number] = '" & ol.SOrderAX & "' and [order_line_number]= '" & ol.SOrderLineAx & "');
if (@iid is null)
begin
SET IDENTITY_INSERT orders_lines_temp ON;
INSERT INTO [dbo].[orders_lines_temp]
           ([id]
		   ,[id_clients]
           ,[id_articles]
           ,[order_number]
           ,[order_date]
           ,[order_line_number]
           ,[order_line_date]
           ,[desired_date]
           ,[requested_date]
           ,[total_quantity]
           ,[delivered_quantity]
           ,[quantity]
           ,[id_orders_lines_states]
           ,[client_block]
           ,[notes])
     VALUES
           (" & ol.IId & ", 
			(select id from clients where ax_reference ='" & ol.sIdClient & "'),
            " & ol.IIdArticles & ",
            '" & ol.SOrderAX & "',
            '" & ol.OrderAXDate & "',
            '" & ol.SOrderLineAx & "',
            '" & ol.OrderLineAXDate & "',
            '" & ol.OrderLineAXDate & "',
            '" & ol.OrderLineAXDate & "',  
            '" & ol.IQuantity & "',
            0,
            0,
            (select id_orders_lines_states from orders_lines where id = " & ol.IId & "),
            '" & If(ol.BClientBlock = 1, 2, 1) & "',  
            '" & ol.SNotes & "')
SET IDENTITY_INSERT orders_lines_temp OFF;
Update orders_lines set id_orders_lines_states = 3 where id = " & ol.IId & ";
select 1;
END
ELSE
Begin
UPDATE orders_lines_temp
set [id_clients] = (select id from clients where ax_reference ='" & ol.sIdClient & "')
           ,[id_articles] = " & ol.IIdArticles & "
           ,[order_number] ='" & ol.SOrderAX & "'
           ,[order_date] = '" & ol.OrderAXDate & "'
           ,[order_line_number] = '" & ol.SOrderLineAx & "'
           ,[order_line_date] = '" & ol.OrderLineAXDate & "'
           ,[desired_date] = '" & ol.OrderLineAXDate & "'
           ,[requested_date] = '" & ol.OrderLineAXDate & "'
           ,[total_quantity] = '" & ol.IQuantity & "'
           ,[delivered_quantity] = 0
           ,[quantity] = 0
           ,[client_block] = '" & ol.BClientBlock & "'
           ,[notes] = '" & ol.SNotes & "'
		   where id =  @iid;
Update orders_lines set id_orders_lines_states = 3 where id = " & ol.IId & ";
select 2;
END"
            Dim cmd As New SqlCommand(sel, con)

            createTemporalOrderLine = cmd.ExecuteScalar()

        Catch ex As Exception
        End Try
        con.Close()
    End Function
    'Insert new order line
    Public Function CreateOrderLine(ByVal ol As orderLine) As String
        CreateOrderLine = ""
        Try
            con.Open()
            Dim insertOrder As String = "
INSERT INTO [dbo].[orders_lines]
           ([id_clients]
           ,[id_articles]
           ,[order_number]
           ,[order_date]
           ,[order_line_number]
           ,[order_line_date]
           ,[desired_date]
           ,[requested_date]
           ,[total_quantity]
           ,[delivered_quantity]
           ,[quantity]
           ,[id_orders_lines_states]
           ,[client_block]
           ,[notes])
     VALUES
           ((select id from clients where ax_reference ='" & ol.sIdClient & "') ,
            " & ol.IIdArticles & ", 
            '" & ol.SOrderAX & "',
            '" & ol.OrderAXDate & "',
            '" & ol.SOrderLineAx & "',
            '" & ol.OrderLineAXDate & "',
            '" & ol.DesiredDate & "',
            '" & ol.RequestedDate & "',  
            '" & ol.IQuantity & "',
            0,
            0,
            2,
            '" & ol.BClientBlock & "',  
            '" & ol.SNotes & "')"
            Dim cmd As New SqlCommand(insertOrder, con)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            CreateOrderLine = ex.Message
        End Try
        con.Close()
    End Function
    'Format update
    Function formatDiferents(sWhere As String, nameColumn As String, sValue As String) As String
        If sWhere = "" Then
            formatDiferents = " Set " & nameColumn & " = '" & sValue & "'"
        Else
            formatDiferents = sWhere & ", " & nameColumn & " = '" & sValue & "'"
        End If
    End Function
    'Test if exists the order line
    Public Function existsLine(ByVal order As String, ByVal line As String) As orderLine
        Dim sel As String = "Select 
(select ax_reference from clients where id = id_clients), 
(select name from clients where id = ol.id_clients),
(select name from articles where id = ol.id_articles),
ol.order_number,
ol.order_date,
ol.order_line_number,
ol.order_line_date,
ol.desired_date,
requested_date,
total_quantity,
client_block ,
notes, 
id_orders_lines_states,
id,
id_articles,
is_deleted
from orders_lines ol where order_number ='" & order & "' and order_line_number ='" & line & "'"
        Dim ola As New orderLine
        Try
            Dim dt As New DataTable
            Dim da As New SqlDataAdapter(sel, con)
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                For Each row In dt.Rows
                    With ola
                        .sIdClient = row(0)
                        .SClient = row(1)
                        .SReferenciaAX = row(2)
                        .SOrderAX = row(3)
                        .OrderAXDate = row(4)
                        .SOrderLineAx = row(5)
                        .OrderLineAXDate = row(6)
                        .DesiredDate = row(7)
                        .RequestedDate = row(8)
                        .IQuantity = row(9)
                        .BClientBlock = row(10)
                        .SNotes = row(11)
                        .IId_states = row(12)
                        .IId = row(13)
                        .IIdArticles = row(14)
                        .BIsDeleted = row(15)
                    End With
                Next
            End If
        Catch ex As Exception
        End Try
        Return ola
    End Function
    'Update order line
    Private Function updateLine(ByVal setString As String, ByVal id As Integer) As Boolean
        updateLine = False
        Dim updateQuery As String = "Update orders_lines " & setString & " where id = " & id
        Try
            Dim cmd As New SqlCommand(updateQuery, con)
            con.Open()
            cmd.ExecuteNonQuery()
            updateLine = True
        Catch ex As Exception
        End Try
        con.Close()
    End Function
    'Find deleted lines
    Public Function findDeleteLines(ByVal ol As List(Of orderLine), ByRef withErrors As Boolean) As String
        findDeleteLines = ""
        Try
            con.Open()
            Dim sel As String = "select * from orders_lines where id_orders_lines_states <> 7 and (internal is null or internal = 0)"
            Dim da As New SqlDataAdapter(sel, con)
            Dim dt As New DataTable
            da.Fill(dt)
            For Each row In dt.Rows
                orderNumber = row("order_number")
                orderLineNumber = row("order_line_number")
                Dim listResult As List(Of orderLine) = ol.FindAll(Function(p) p.SOrderAX = orderNumber And p.SOrderLineAx = orderLineNumber)
                Dim delQuery As String = ""
                If Not listResult.Count > 0 Then
                    If row("id_orders_lines_states") = 2 Then
                        delQuery = "Delete from orders_lines where id = " & row("id")
                        findDeleteLines = findDeleteLines & "Línea " & orderLineNumber & " del pedido " & orderNumber & " eliminada definitivamente."
                    Else
                        delQuery = "update orders_lines set is_deleted = 1 where id = " & row("id")
                        findDeleteLines = findDeleteLines & "Línea " & orderLineNumber & " del pedido " & orderNumber & " pendiente de confirmación."
                    End If
                    Dim cmd As New SqlCommand(delQuery, con)
                    cmd.ExecuteNonQuery()
                End If
            Next
        Catch ex As Exception
            withErrors = True
            findDeleteLines = findDeleteLines & "*** Error al intenter eliminar líneas." & vbCrLf & ex.Message
        End Try
        con.Close()
    End Function
#End Region
End Class
