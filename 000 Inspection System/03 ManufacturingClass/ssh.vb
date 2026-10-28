Imports Renci.SshNet
Imports System.IO
Imports System.Net.Mail
Imports System.Threading.Tasks

Module ssh

    Public Function UploadFileOverSftp(
        host As String,
        username As String,
        password As String,
        localFilePath As String,
        remoteFilePath As String
    ) As Boolean
        Try
            Using sftp As New SftpClient(host, username, password)
                sftp.Connect()
                If Not sftp.IsConnected Then Throw New Exception("SFTP 연결 실패")

                Using fileStream As New FileStream(localFilePath, FileMode.Open)
                    sftp.UploadFile(fileStream, remoteFilePath)
                End Using

                sftp.Disconnect()
            End Using

            Return True

        Catch ex As Exception
            Throw New Exception("SFTP 파일 전송 중 오류 발생: " & ex.Message, ex)
        End Try
    End Function

    Public Function ExecuteRemoteCommand(
        host As String,
        username As String,
        password As String,
        commandToExecute As String
    ) As String
        Try
            Using client As New SshClient(host, username, password)
                client.Connect()
                If Not client.IsConnected Then Throw New Exception("SSH 연결 실패")

                Dim command = client.CreateCommand(commandToExecute)
                Dim result As String = command.Execute()

                client.Disconnect()
                Return result
            End Using

        Catch ex As Exception
            Throw New Exception("SSH 명령 실행 중 오류 발생: " & ex.Message, ex)
        End Try
    End Function

    Public Function ExecuteRemoteCommandAsync(
    host As String,
    username As String,
    password As String,
    commandToExecute As String
) As Task(Of String)
        Return Task.Run(Function()
                            Using client As New SshClient(host, username, password)
                                client.Connect()
                                If Not client.IsConnected Then Throw New Exception("SSH 연결 실패")

                                Dim cmd = client.CreateCommand(commandToExecute)
                                Dim result = cmd.Execute()

                                client.Disconnect()
                                Return result
                            End Using
                        End Function)
    End Function
End Module