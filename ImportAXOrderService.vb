Imports System.IO
Imports System.Threading

Public Class ImportAXOrderService

#Region "VARIABLES"
    Dim funcOl As New OrdersLinesDAO
    Dim funcEncript As New Simple3Des
    Dim msnLog As String
    Dim importPath As String = System.AppDomain.CurrentDomain.BaseDirectory & "\Pedidos"
    Dim importMinInterval As Integer
    Dim imported As Boolean
    Dim withErrors As Boolean
    Dim iIdImport As Integer
    Public stringConnection As String
    Dim Worker As Thread
#End Region
    Protected Overrides Sub OnStart(ByVal args() As String)
        'Start the worker thread
        funcOl.con.ConnectionString = funcEncript.decryptConnectionString()
        logService(funcOl.con.ConnectionString & vbCrLf)
        If funcOl.testConnection Then
            logService("Conexión OK" & vbCrLf)
            If funcOl.getConfig(importMinInterval) Then
                logService("Configuración OK" & vbCrLf)
                ImportOrders()
                logService("Inicio de ciclo" & vbCrLf)
                Worker = New Thread(AddressOf DoWork)
                Worker.Start()
            Else
                logService("Configuración fallida" & vbCrLf)
            End If
        Else
            logService("Conexión fallida" & vbCrLf)
        End If
    End Sub
    'Log in folder
    Sub logService(txt As String)
        Dim ruta As String = System.AppDomain.CurrentDomain.BaseDirectory & "\servicelog.txt"
        If Not File.Exists(ruta) Then
            Dim fs As New FileStream(ruta, FileMode.CreateNew)
            fs.Close()
        End If
        Dim escritor As StreamWriter
        escritor = File.AppendText(ruta)
        escritor.Write(Now & " - " & txt)
        escritor.Flush()
        escritor.Close()
    End Sub
    Protected Overrides Sub OnStop()
        Worker.Interrupt()
        Worker.Join()
    End Sub
    Protected Sub DoWork(ByVal args() As String)
        logService("Espera " & Now & vbCrLf)
        'Worker thread loop
        While True
            Try
                Thread.Sleep(importMinInterval * 60000)
                ImportOrders()
            Catch ex As Exception
                logService("***Error en el loop." & Now & vbCrLf)
                Return
            End Try
        End While
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
                    logService("Inicio de la importación." & vbCrLf & "Archivo a importar: " & fileImport & vbCrLf)
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
                                .SReferenciaAX = textArray(2)
                                .SOrderAX = textArray(3)
                                .OrderAXDate = textArray(4)
                                .SOrderLineAx = textArray(5)
                                .OrderLineAXDate = textArray(6)
                                .DesiredDate = textArray(7)
                                .RequestedDate = textArray(8)
                                .IQuantity = textArray(9)
                                If textArray(10) = 0 Or textArray(10) = "" Then
                                    .BClientBlock = 1
                                Else
                                    .BClientBlock = 2
                                End If
                                .SNotes = textArray(11)
                            End With
                            ordersLinesList.Add(ol)
                        Catch ex As Exception
                            logService("*** Error Núm. línea: " & numberLine & vbNewLine & vbCrLf)
                            logText("*** Error Núm. línea: " & numberLine & vbCrLf)
                        End Try
                        numberLine += 1
                    Loop
                    objReader.Close()
                    If ordersLinesList.Count > 0 Then
                        logService("Se han encontrado un total de " & ordersLinesList.Count & " líneas a importar." & vbCrLf)
                        logText("Se han encontrado un total de " & ordersLinesList.Count & " líneas a importar.")
                        'Create imports_services in BBDD
                        iIdImport = funcOl.createImportServices(fileImport)
                        Dim sRegisterClients As String = registerClients(ordersLinesList)
                        'Register clients.
                        If sRegisterClients = "" Then
                            'Register lines
                            If Not registerLines(ordersLinesList) Then
                                withErrors = True
                                logService("*** Error al registrar las líneas." & vbCrLf)
                                logText("*** Error al registrar las líneas.")
                            End If
                            'Delete waiting lines missing.
                            Dim logDeleted As String = funcOl.findDeleteLines(ordersLinesList, withErrors)
                            If logDeleted <> "" Then
                                logService("Se han eliminado las siguientes líneas:" & vbCrLf & logDeleted & vbCrLf)
                                logText("Se han eliminado las siguientes líneas:" & vbCrLf & logDeleted)
                            End If
                        Else
                            withErrors = True
                            logService("*** Error al crear los clientes." & vbCrLf & sRegisterClients & vbCrLf)
                            logText("*** Error al crear los clientes." & vbCrLf & sRegisterClients)
                        End If
                    Else
                        logService("No se han encontrado líneas a importar." & vbCrLf)
                        logText("No se han encontrado líneas a importar.")
                    End If
                    imported = True
                Else
                    logService("No hay archivo en la ruta." & vbCrLf)
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
            logService("Fin de la importación" & vbCrLf)
            logText("Fin de la importación")
            If iIdImport > 0 Then
                funcOl.updateImportServices(iIdImport, msnLog)
            End If
            imported = False
            withErrors = False
            msnLog = ""
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
        registerClients = ""
        Try
            For Each ol In ordersLinesList
                With ol
                    If Not funcOl.testClient(.sIdClient) Then
                        If funcOl.CreateClient(.SClient, .sIdClient) Then
                            logText("Nuevo cliente " & .SClient & " registrado, identificador de AX " & .sIdClient & " .")
                            logService("Nuevo cliente " & .SClient & " registrado, identificador de AX " & .sIdClient & " ." & vbCrLf)
                        Else
                            logText("Error al crear  el cliente " & .SClient & " registrado, identificador de AX " & .sIdClient & " .")
                            logService("Error al crear  el cliente " & .SClient & " registrado, identificador de AX " & .sIdClient & " ." & vbCrLf)
                        End If
                    Else
                        If Not funcOl.UpdateClient(.sIdClient, .BClientBlock, .SClient) Then
                            logText("Error al  actualizar el cliente " & .SClient & " registrado, identificador de AX " & .sIdClient & " .")
                            logService("Error al  actualizar el cliente " & .SClient & " registrado, identificador de AX " & .sIdClient & " ." & vbCrLf)
                        End If
                    End If
                End With
            Next
        Catch ex As Exception
            registerClients = ex.Message
            logText("Error al registrar los clientes.")
            logService("Error al registrar los clientes." & vbCrLf)
        End Try
    End Function
    'Register lines
    Function registerLines(ByVal ordersLinesList As List(Of orderLine)) As Boolean
        registerLines = True
        For Each ol In ordersLinesList
            Try
                ol.IIdArticles = funcOl.testArticles(ol.SReferenciaAX, ol.sIdClient)
                If ol.IIdArticles = 0 Then
                    withErrors = True
                    funcOl.insertMissingArticles(ol.SReferenciaAX, ol.SOrderAX, ol.SOrderLineAx)
                    logService("*** Error, no se ha encontrado el articulo " & ol.SReferenciaAX & " en la base de datos. No se importará la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX & vbCrLf)
                    logText("*** Error, no se ha encontrado el articulo " & ol.SReferenciaAX & " en la base de datos. No se importará la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX)
                Else
                    'Compare lines
                    Dim sCompareLines As String = funcOl.compareLines(ol)
                    If sCompareLines = "" Then
                        Dim resultCreate As String = funcOl.CreateOrderLine(ol)
                        If resultCreate = "" Then
                            logService("Se ha creado correctamente la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX & "." & vbCrLf)
                            logText("Se ha creado correctamente la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX & ".")
                        Else
                            withErrors = True
                            logService("*** Error al crear la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX & "." & vbCrLf & resultCreate & vbCrLf)
                            logText("*** Error al crear la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX & ".")
                        End If
                    Else
                        If Not sCompareLines = "NR" Then
                            logText(sCompareLines)
                            logService(sCompareLines & vbCrLf)
                        End If
                    End If
                End If
            Catch ex As Exception
                withErrors = True
                registerLines = False
                logService("*** Error al registrar la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX & vbCrLf & ex.Message & vbCrLf)
                logText("*** Error al registrar la línea " & ol.SOrderLineAx & " del pedido " & ol.SOrderAX & vbCrLf & ex.Message)
            End Try
        Next
    End Function

End Class
