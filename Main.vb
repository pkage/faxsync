Imports FAXCOMEXLib
Module Main
    Private Sub ShellSync(command As String)
        ' Much, much faster to do this async
        Dim wsh As Object = CreateObject("WScript.Shell")
        Dim waitOnReturn As Boolean : waitOnReturn = True
        Dim windowStyle As Integer : windowStyle = 0

        wsh.Run(command, windowStyle, waitOnReturn)
    End Sub
    Private Function GetFaxInboxFolder() As String
        Dim destination As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        Return System.IO.Path.Combine(destination, "Fax Inbox")
    End Function

    Private Function RemoveInvalidFileNameChars(Unsafe As String) As String
        For Each invalidChar In IO.Path.GetInvalidFileNameChars
            Unsafe = Unsafe.Replace(invalidChar, "")
        Next
        Return Unsafe
    End Function

    Private Function GetFaxFilename(fax As FAXCOMEXLib.FaxIncomingMessage) As String
        Dim CallerId As String = fax.CallerId
        Dim FaxId As String = fax.Id

        If CallerId.Trim() = "" Then
            CallerId = "Private Number"
        End If

        Dim filename = "Fax from " + CallerId.Trim() + " (" + FaxId + ").tif"
        Return RemoveInvalidFileNameChars(filename)
    End Function
    Private Sub EnsureFaxInboxFolderExists()
        My.Computer.FileSystem.CreateDirectory(GetFaxInboxFolder())
    End Sub

    Private Sub ChangeFileTimes(filename As String, creation As Date, lastwrite As Date)
        ' HERE BE DRAGONS: set the creation time of the file with powershell
        ' There's not a way to do this with VB
        ' Format : "mm/dd/yyyy hh:mm am/pm")
        Dim lastwrite_str As String = lastwrite.ToString("MM/dd/yyyy hh:mm tt")
        Dim creation_str As String = creation.ToString("MM/dd/yyyy hh:mm tt")

        Dim lastwrite_cmd As String = "powershell.exe -Command ""$(Get-Item \""" + filename + "\"").lastwritetime=$(Get-Date \""" + lastwrite_str + "\"")"""
        Dim creation_cmd As String = "powershell.exe -Command ""$(Get-Item \""" + filename + "\"").creationtime=$(Get-Date \""" + creation_str + "\"")"""

        Shell(lastwrite_cmd)
        Shell(creation_cmd)
    End Sub

    Private Sub CopyFax(fax As FAXCOMEXLib.FaxIncomingMessage)
        On Error GoTo CopyFaxError_Handler
        Dim destination As String = GetFaxInboxFolder()
        Dim filename As String = GetFaxFilename(fax)
        destination = System.IO.Path.Combine(destination, filename)
        'Console.WriteLine(filename)
        'Exit Sub
        fax.CopyTiff(destination)
        ChangeFileTimes(destination, fax.TransmissionEnd, fax.TransmissionEnd)
        Exit Sub
CopyFaxError_Handler:
        Console.WriteLine("Error: " & Hex(Err.Number) & ", " & Err.Description)
    End Sub

    Private Sub CopyLastNDaysFaxes(ServerString As String, nDays As Integer)
        Dim objFaxServer As New FAXCOMEXLib.FaxServer
        Dim objFaxIncomingMessageIterator As FAXCOMEXLib.FaxIncomingMessageIterator
        Dim objFaxIncomingMessage As FAXCOMEXLib.FaxIncomingMessage
        Dim Prefetch As String
        Dim Answer As String
        Dim FileName As String

        Dim i As Integer

        Dim A As Object

        Dim cutoffDate As Date = Date.Today.AddDays(-1 * nDays)

        'Error handling
        On Error GoTo Error_Handler

        'Connect to the fax server
        Console.WriteLine("Connecting to the fax server...")
        objFaxServer.Connect(ServerString)
        Console.WriteLine("Connected to " + ServerString)

        'Get the iterator and Set the prefetch buffer size
        Prefetch = 50

        'Refresh the archive
        Console.Write("Refreshing incoming fax archive...")
        'objFaxServer.Folders.IncomingArchive.Refresh()
        Console.WriteLine(" skipped.")

        Console.Write("Getting an iterator for the incoming archive (prefetch " + Prefetch + ")...")
        objFaxIncomingMessageIterator = objFaxServer.Folders.IncomingArchive.GetMessages(Prefetch)
        Console.WriteLine("done.")

        'Set the iterator cursor to the first message in the buffer
        Console.Write("Getting the first message...")
        objFaxIncomingMessageIterator.MoveFirst()
        Console.WriteLine("done.")
        While True

            'Get the message.
            objFaxIncomingMessage = objFaxIncomingMessageIterator.Message

            'Check for end of file.
            If objFaxIncomingMessageIterator.AtEOF = True Then
                Console.WriteLine("End of iterator Reached")
                Exit Sub
            End If

            'FileName = InputBox("Provide path and name of file for TIFF copy, e.g. c:\MyFax.tiff")
            'objFaxIncomingMessage.CopyTiff(FileName)
            Console.Write(GetFaxFilename(objFaxIncomingMessage))
            If cutoffDate < objFaxIncomingMessage.TransmissionEnd Then
                Console.WriteLine(" ...processed.")
                CopyFax(objFaxIncomingMessage)
            Else
                Console.WriteLine(" ...skipped.")
                'CopyFax(objFaxIncomingMessage)
            End If

            'Set the iterator cursor to the next message
            objFaxIncomingMessageIterator.MoveNext()

        End While
        Exit Sub

Error_Handler:
        'Implement error handling at the end of your subroutine. This 
        ' implementation is for demonstration purposes
        Console.WriteLine("Error number: " & Hex(Err.Number) & ", " & Err.Description)

    End Sub

    Private Sub ShowFaxInfo(ServerString As String)
        Dim objFaxServer As New FAXCOMEXLib.FaxServer

        'Error handling
        On Error GoTo Error_Handler

        'Connect to the fax server
        Console.WriteLine("Connecting to " + ServerString + "...")
        objFaxServer.Connect(ServerString)
        Console.WriteLine("Connected!")

        'Display server properties

        MsgBox("Server information" & vbCrLf &
        vbCrLf & "API Version: " & objFaxServer.APIVersion &
        vbCrLf & "Debug: " & objFaxServer.Debug &
        vbCrLf & "Build and version: " & objFaxServer.MajorBuild & "." &
        objFaxServer.MinorBuild & "." & objFaxServer.MajorVersion & "." &
        objFaxServer.MinorVersion & "." &
        vbCrLf & "Server name: " & objFaxServer.ServerName)
        Exit Sub

Error_Handler:
        'Implement error handling at the end of your subroutine. This 
        ' implementation is for demonstration purposes
        MsgBox("Error number: " & Hex(Err.Number) & ", " & Err.Description)

    End Sub

    Sub Main()
        Dim clArgs() As String = Environment.GetCommandLineArgs()
        Dim daysBack As Integer = 0
        Dim ServerString As String = ""
        Dim showInfo As Boolean = False

        ' Process command line arguments
        If clArgs.Count() = 3 Then
            If clArgs(2) = "info" Then
                showInfo = True
            Else
                daysBack = Integer.Parse(clArgs(2))
            End If
        ElseIf clArgs.Count() = 2 Then
            daysBack = 7
        Else
            Console.WriteLine("Usage: FaxSync \\faxserver (days back OR ""info"")")
            Exit Sub
        End If
        ServerString = clArgs(1)

        If showInfo Then
            ShowFaxInfo(ServerString)
        Else
            EnsureFaxInboxFolderExists()
            Console.WriteLine("Copying last " + daysBack.ToString() + " days.")
            CopyLastNDaysFaxes(ServerString, daysBack)
        End If

    End Sub

End Module
