Imports System.Threading
Imports System.Net
Imports System.Net.Sockets
Imports System.IO

'20140319 - wave 3 test version on sysfeed5

'20150413 - VS2013 version

Module Module1

    'client.Connect("127.0.0.1", 2202)
    'client.Connect("192.168.1.67", 2202) ' hp desktop
    'client.Connect("10.250.186.147", 33401) ' McKesson Server
    'client.Connect("10.48.240.67", 10250) ' dave's connection


    Dim NWStream As NetworkStream
    Dim client As New TcpClient
    Dim file As System.IO.StreamWriter
    'Public objIniFile As New INIFile("d:\W3Production\HL7Transmitter.ini") 'Prod 20140818
    Public objIniFile As New INIFile("C:\W3Feeds\HL7Transmitter.ini") 'Test 20140818
    Dim IPAddress As String = ""
    Dim port As String = ""
    Dim DFTDirectory As String = ""





    Sub main()

        Try
            file = My.Computer.FileSystem.OpenTextFileWriter("d:\dftLog\ULH\transmitlog.txt", True)
            IPAddress = objIniFile.GetString("Transmitter", "IPAddressULH", "(none)")
            port = objIniFile.GetString("Transmitter", "PortULH", "(none)")
            DFTDirectory = objIniFile.GetString("Transmitter", "DFTDirectoryULH", "(none)")
            'start connection
            client.Connect(IPAddress, port)

            Application.DoEvents()

            If client.Connected Then
                file.WriteLine("----------------------------------------------------")
                file.WriteLine("Server: " & IPAddress & ":" & port & " connected.")
                file.WriteLine("----------------------------------------------------" & vbCrLf)
                Application.DoEvents()
                SendFiles()
            End If




        Catch ex As Exception
            'handle errors
            ' file.WriteLine(ex.Message & vbCrLf)
            '20171127 - Create an error to receive email warning
            CreateErrorFile(ex.ToString())


        Finally
            'finalize
            file.Close()

        End Try



    End Sub

    Public Sub SendFiles()
        '20121108 - this is the file sending routine broken out to itterate through a directory
        'and process all the files in it.
        Try
            Dim filename As String = ""
            Dim dir As String = ""
            NWStream = client.GetStream()
            Dim bytesToSend(client.SendBufferSize) As Byte

            Dim dirs As String() = Directory.GetFiles(DFTDirectory, "HL7.*")


            For Each filename In dirs

                Dim theFile As New FileInfo(filename)
                Dim fileStr As New FileStream(filename, FileMode.Open, FileAccess.Read)
                Dim fileReader As New BinaryReader(fileStr)
                Dim numBytesRead As Integer = 0
                Dim i As Integer = 0
                Do Until i >= theFile.Length - 1
                    numBytesRead = fileStr.Read(bytesToSend, 0, bytesToSend.Length)
                    NWStream.Write(bytesToSend, 0, numBytesRead)
                    i = i + numBytesRead
                    NWStream.Flush()

                Loop
                file.WriteLine(filename & " Sent.")

                ' Receive the TcpServer.response.
                ' Buffer to store the response bytes.
                Dim Data = New [Byte](256) {}

                ' String to store the response ASCII representation.
                Dim responseData As [String] = [String].Empty

                ' Read the first batch of the TcpServer response bytes.
                Dim bytes As Int32 = NWStream.Read(Data, 0, Data.Length)
                responseData = System.Text.Encoding.ASCII.GetString(Data, 0, bytes)

                'check to see if we got a positive response. Looking for MSA|AA|
                If InStr(responseData, "MSA|AA|") <> 0 Then
                    'file.WriteLine("======================================== " & vbCrLf)
                    file.WriteLine("Received: " & responseData & vbCrLf)

                    '20121108 - send complete. delete the file but close the filestream first because it is using f1.
                    Application.DoEvents()
                    fileStr.Close()
                    theFile.Delete()

                Else
                    'did not receive a positice acknowledgement so save it in the problems directory
                    file.WriteLine(filename & " not acknowledged for. Copied to problems directory." & vbCrLf)

                    Application.DoEvents()

                    'first, close the file stream to release the file so we can work with it.
                    fileStr.Close()
                    Dim fi2 As FileInfo = New FileInfo(DFTDirectory & "problems\" & theFile.Name)
                    fi2.Delete()
                    theFile.CopyTo(DFTDirectory & "problems\" & theFile.Name)
                    theFile.Delete()
                End If

            Next

        Catch ex As Exception
            ' file.WriteLine(ex.Message & vbCrLf)
            '20171127 - Create an error to receive email warning
            CreateErrorFile(ex.ToString())
        End Try
    End Sub

    Public Sub CreateErrorFile(ByVal errorstring As String)
        Dim errorfilepath As String = objIniFile.GetString("Transmitter", "ErrorDirULH", "(none)")
        Dim errorfilename As String = String.Format("_DFTError-'" & "'_{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
        Dim errorfile = New StreamWriter(errorfilepath & errorfilename, True)
        errorfile.Write(errorstring)
        errorfile.Close()
    End Sub

End Module
