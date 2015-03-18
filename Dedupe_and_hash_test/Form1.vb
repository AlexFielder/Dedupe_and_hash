Imports System.Management.Automation.Runspaces
Imports System.Management.Automation
Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Text

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim filedialog As OpenFileDialog = New OpenFileDialog()
        filedialog.Filter = "All files (*.*)|*.*"
        filedialog.Multiselect = False
        Dim file1 As String = String.Empty
        Dim file2 As String = String.Empty
        If filedialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            file1 = filedialog.FileName
            If filedialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                file2 = filedialog.FileName
            End If
        End If
        Dim result As String = RunScript(file1, file2)
        MessageBox.Show("The file hash result is: " & result)
    End Sub
    Private Function RunScript(ByVal file1 As String, file2 As String) As String

        'Throw New notimplementedexception
        ' create Powershell runspace 
        Dim MyRunSpace As Runspace = RunspaceFactory.CreateRunspace()
        ' open it 
        MyRunSpace.Open()
        ' create a pipeline and feed it the script text 
        Dim scripttext As String = "get-filehash '" & file1 & "' -Algorithm MD5"
        Dim MyPipeline As Pipeline = MyRunSpace.CreatePipeline()
        MyPipeline.Commands.AddScript(scripttext)
        Dim Results As Collection(Of PSObject) = MyPipeline.Invoke()
        Dim record As PSObject = Results.Item(0)
        MessageBox.Show(Results.Item(0).ToString())
        Dim hash1 As PSMemberInfo = Results(0).Properties("Hash")
        Results.Clear()
        MyPipeline = MyRunSpace.CreatePipeline()
        'file two
        scripttext = "get-filehash '" & file2 & "' -Algorithm MD5"
        MyPipeline.Commands.AddScript(scripttext)
        Results = MyPipeline.Invoke()
        Dim hash2 As PSMemberInfo = Results(0).Properties("Hash")
        If Not hash1.Value = hash2.Value Then
            Return "No file hash match!"
        Else
            Return "Files hashes match!"
        End If
        MyRunSpace.Close()

    End Function
    
End Class
