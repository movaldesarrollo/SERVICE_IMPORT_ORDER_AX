Imports System.IO

Public Class main

#Region "VARIABLES"
    Dim funcOl As New OrdersLinesDAO
    Dim funcEncript As New Simple3Des
    Dim msnLog As String
    Dim importPath As String
    Dim importMinInterval As Integer
    Dim imported As Boolean
    Dim withErrors As Boolean
    Dim iIdImport As Integer
    Public stringConnection As String
    Dim t As New Timer
#End Region

#Region "EVENTS"
    'Button
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        If funcOl.testConnection Then
            If funcOl.getConfig(importPath, importMinInterval) Then
                ImportOrders()
                t.Interval = importMinInterval * 60000
                AddHandler t.Tick, AddressOf t_tick
                t.Start()
            End If
        End If
    End Sub
    'Timer
    Private Sub t_tick()
        ImportOrders()
    End Sub
#End Region

#Region "FUNCTIONS"
    Private Sub main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            funcOl.con.ConnectionString = funcEncript.decryptConnectionString()
        Catch ex As Exception
            Close()
        End Try
    End Sub
    Sub ImportOrders()
        For Each fileImport As String In My.Computer.FileSystem.GetFiles(importPath, FileIO.SearchOption.SearchTopLevelOnly, "*.txt")
            Try
                'Test if exists file.
                If File.Exists(fileImport) Then
                    Dim ordersLinesList As New List(Of orderLine)
                    'Create stream reader
                    Dim objReader As New System.IO.StreamReader(fileImport)
                    'Create variable for register log
                    logText("Inicio de la importación." & vbCrLf & "Archivo a importar: " & fileImport & vbCrLf)
                    'Contains number line at file import
                    Dim numberLine As Integer = 1
                    'We begin the loop, we read lines.
                    Do While objReader.Peek() <> -1
                        Dim textArray() As String = Split(objReader.ReadLine(), ";")
                        Try
                            Dim ol As New orderLine
                            With ol
                                .sIdClient = textArray(0)
                                .SClient = textArray(1)
                                .SMonoSKU = textArray(2)
                                .SOrderAX = textArray(3)
                                .OrderAXDate = textArray(4)
                                .SOrderLineAx = textArray(5)
                                .OrderLineAXDate = textArray(6)
                                .DesiredDate = textArray(7)
                                .RequestedDate = textArray(8)
                                .IQuantity = textArray(9)
                                If textArray(10) = 0 Then
                                    .BClientBlock = 1
                                Else
                                    .BClientBlock = 2
                                End If
                                .SNotes = textArray(11)
                            End With
                            ordersLinesList.Add(ol)
                        Catch ex As Exception
                            logText("*** Error Núm. línea: " & numberLine & vbNewLine)
                        End Try
                        numberLine += 1
                    Loop
                    objReader.Close()
                    If ordersLinesList.Count > 0 Then
                        logText("Se han encontrado un total de " & ordersLinesList.Count & " líneas a importar.")
                        'Create imports_services in BBDD
                        iIdImport = funcOl.createImportServices(fileImport)
                        Dim sRegisterClients As String = registerClients(ordersLinesList)
                        'Register clients.
                        If sRegisterClients = "" Then
                            'Register lines
                            If Not registerLines(ordersLinesList) Then
                                withErrors = True
                                logText("*** Error al registrar las líneas.")
                            End If
                            'Delete waiting lines missing.
                            Dim logDeleted As String = funcOl.findDeleteLines(ordersLinesList, withErrors)
                            If logDeleted <> "" Then
                                logText("Se han eliminado las siguientes líneas:" & vbCrLf & logDeleted)
                            End If
                        Else
                            withErrors = True
                            logText(" *** Error al crear los clientes." & vbCrLf & sRegisterClients)
                        End If
                    Else
                        logText("No se han encontrado líneas a importar.")
                    End If
                    imported = True
                Else
                    logText("No hay archivo en la ruta.")
                End If
                If imported = True Then
                    Dim copyToPath As String
                    If withErrors = True Then
                        copyToPath = importPath & "\Importados_con_errores"
                    Else
                        copyToPath = importPath & "\Importados"
                    End If
                    If Not Directory.Exists(copyToPath) Then
                        Directory.CreateDirectory(copyToPath)
                    End If
                    Dim origin As String = fileImport
                    Dim destiny As String = copyToPath & "\" & My.Computer.FileSystem.GetFileInfo(fileImport).Name
                    If File.Exists(destiny) Then
                        File.Delete(destiny)
                    End If
                    File.Move(origin, destiny)
                End If
            Catch ex As Exception
                logText(ex.Message)
            End Try
            logText("Fin de la importación")
            If iIdImport > 0 Then
                funcOl.updateImportServices(iIdImport, msnLog)
            End If
            imported = False
            withErrors = False
            TextBox1.Text = msnLog & vbCrLf & "-----------------------------------------------------" & vbCrLf & TextBox1.Text
            msnLog = ""
            Application.DoEvents()
        Next
    End Sub

    'Write in var log
    Sub logText(ByVal logTextNow As String)
        If msnLog = "" Then
            msnLog = Now() & " - " & logTextNow
        Else
            msnLog = msnLog & vbCrLf & Now() & " - " & logTextNow
        End If
    End Sub
    'Register new clients.
    Function registerClients(ByVal ordersLinesList As List(Of orderLine)) As String
        Try
            For Each ol In ordersLinesList
                With ol
                    If Not funcOl.testClient(.sIdClient) Then
                        If funcOl.CreateClient(.SClient, .sIdClient) Then
                            logText("Nuevo cliente " & .SClient & " registrado, identificador de AX " & .sIdClient & " .")
                        End If
                    End If
                End With
            Next
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    'Register lines
    Function registerLines(ByVal ordersLinesList As List(Of orderLine)) As Boolean
        registerLines = True
        For Each ol In ordersLinesList
            Try
                ol.IIdArticles = funcOl.testArticles(ol.SMonoSKU, ol.sIdClient)
                If ol.IIdArticles = 0 Then
                    logText("No se ha encontrado el articulo " & ol.SMonoSKU & " en la base de datos. No se importará la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX)
                Else
                    'Compare lines
                    Dim sCompareLines As String = funcOl.compareLines(ol)
                    If sCompareLines = "" Then
                        If funcOl.CreateOrderLine(ol) Then
                            logText("Se ha creado correctamente la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX & ".")
                        Else
                            withErrors = True
                            logText("*** Error al crear la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX & ".")
                        End If
                    Else
                        If Not sCompareLines = "NR" Then
                            logText(sCompareLines)
                        End If
                    End If
                End If
            Catch ex As Exception
                withErrors = True
                registerLines = False
                logText("*** Error al registrar la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX & vbCrLf & ex.Message)
            End Try
            Application.DoEvents()
        Next
    End Function
#End Region
End Class
