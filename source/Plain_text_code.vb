'created by Tim241 on github
'You can edit my Source code, but you can't publish it by yourself. You need to ask me if I want to edit the code in the Source.
Imports System.IO
Public Class Form1
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

        TextBox1.Text = ("")


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox2.Text = "'This script is generated in the Duckyscript2VBScript © converter." 'Do not remove this line!!
        TextBox2.Text = TextBox2.Text & Environment.NewLine & "set wshshell = wscript.createobject(""wscript.shell"")"
        TextBox2.Text = TextBox2.Text & Environment.NewLine & "set objshell = createobject(""shell.application"")"
        If File.Exists(IO.Path.Combine(IO.Path.GetTempPath, "data.txt")) = True Then
            File.Delete(IO.Path.Combine(IO.Path.GetTempPath, "data.txt"))
        End If
        TextBox2.Text = TextBox2.Text & Environment.NewLine & TextBox1.Text
        TextBox2.Text = TextBox2.Text.Replace("^", "{^}")
        TextBox2.Text = TextBox2.Text.Replace("(", "{(}")
        TextBox2.Text = TextBox2.Text.Replace(")", "{)}")
        TextBox2.Text = TextBox2.Text.Replace("'", "{'}")
        TextBox2.Text = TextBox2.Text.Replace("|", "{|}")
        TextBox2.Text = TextBox2.Text.Replace("%", "{%}")
        TextBox2.Text = TextBox2.Text.Replace("""", """""")
        TextBox2.Text = TextBox2.Text.Replace("{'}This script is generated in the Duckyscript2VBScript © converter.", "'This script is generated in the Duckyscript2VBScript © converter.") 'Do not remove this line!!
        TextBox2.Text = TextBox2.Text.Replace("set wshshell = wscript.createobject{(}""""wscript.shell""""{)}", "set wshshell = wscript.createobject(""wscript.shell"")")
        TextBox2.Text = TextBox2.Text.Replace("set objshell = createobject{(}""""shell.application""""{)}", "set objshell = createobject(""shell.application"")")
        TextBox2.Text = TextBox2.Text.Replace("STRING ", "WshShell.SendKeys(""") 'The space "string[space]" is needed.
        TextBox2.Text = TextBox2.Text.Replace("ENTER", "WshShell.SendKeys(""{ENTER}"")")
        TextBox2.Text = TextBox2.Text.Replace("ALT F4", "WshShell.SendKeys(""%{f4}"")")
        TextBox2.Text = TextBox2.Text.Replace("WINDOWS M", "objShell.MinimizeAll")
        TextBox2.Text = TextBox2.Text.Replace("WINDOWS L", "objShell.Run ""%windir%\System32\rundll32.exe user32.dll,LockWorkStation""")
        TextBox2.Text = TextBox2.Text.Replace("GUI L", "objShell.Run ""%windir%\System32\rundll32.exe user32.dll,LockWorkStation""")
        TextBox2.Text = TextBox2.Text.Replace("WINDOWS D", "objShell.MinimizeAll")
        TextBox2.Text = TextBox2.Text.Replace("GUI D", "objShell.MinimizeAll")
        TextBox2.Text = TextBox2.Text.Replace("GUI M", "objShell.MinimizeAll")
        TextBox2.Text = TextBox2.Text.Replace("ALT", "WshShell.SendKeys(""%"")")
        TextBox2.Text = TextBox2.Text.Replace("WINDOWS r", "objShell.FileRun")
        TextBox2.Text = TextBox2.Text.Replace("GUI r", "objShell.FileRun")
        TextBox2.Text = TextBox2.Text.Replace("GUI", "WshShell.SendKeys(""^{Esc}"")")
        TextBox2.Text = TextBox2.Text.Replace("WINDOWS", "WshShell.SendKeys(""^{Esc}"")")
        TextBox2.Text = TextBox2.Text.Replace("MENU", "WshShell.SendKeys(""{HOME}"")")
        TextBox2.Text = TextBox2.Text.Replace("DELAY", "WScript.Sleep")
        TextBox2.Text = TextBox2.Text.Replace("UPARROW", "WshShell.SendKeys(""{UP}"")")
        TextBox2.Text = TextBox2.Text.Replace("UP", "WshShell.SendKeys(""{UP}"")")
        TextBox2.Text = TextBox2.Text.Replace("DOWNARROW", "WshShell.SendKeys(""{DOWN}"")")
        TextBox2.Text = TextBox2.Text.Replace("DOWN", "WshShell.SendKeys(""{DOWN}"")")
        TextBox2.Text = TextBox2.Text.Replace("LEFTARROW", "WshShell.SendKeys(""LEFT}"")")
        TextBox2.Text = TextBox2.Text.Replace("LEFT", "WshShell.SendKeys(""LEFT}"")")
        TextBox2.Text = TextBox2.Text.Replace("RIGHTARROW", "WshShell.SendKeys(""{RIGHT}"")")
        TextBox2.Text = TextBox2.Text.Replace("RIGHT", "WshShell.SendKeys(""{RIGHT}"")")
        TextBox2.Text = TextBox2.Text.Replace("F1", "WshShell.SendKeys(""{F1}}"")")
        TextBox2.Text = TextBox2.Text.Replace("F2", "WshShell.SendKeys(""{F2}"")")
        TextBox2.Text = TextBox2.Text.Replace("F3", "WshShell.SendKeys(""{F3}"")")
        TextBox2.Text = TextBox2.Text.Replace("F4", "WshShell.SendKeys(""{F4}"")")
        TextBox2.Text = TextBox2.Text.Replace("F5", "WshShell.SendKeys(""{F5}"")")
        TextBox2.Text = TextBox2.Text.Replace("F6", "WshShell.SendKeys(""{F6}"")")
        TextBox2.Text = TextBox2.Text.Replace("F7", "WshShell.SendKeys(""{F7}"")")
        TextBox2.Text = TextBox2.Text.Replace("F8", "WshShell.SendKeys(""{F8}"")")
        TextBox2.Text = TextBox2.Text.Replace("F9", "WshShell.SendKeys(""{F9}"")")
        TextBox2.Text = TextBox2.Text.Replace("F10", "WshShell.SendKeys(""{F10}"")")
        TextBox2.Text = TextBox2.Text.Replace("F11", "WshShell.SendKeys(""{F11}"")")
        TextBox2.Text = TextBox2.Text.Replace("F12", "WshShell.SendKeys(""{F12}"")")
        TextBox2.Text = TextBox2.Text.Replace("F13", "WshShell.SendKeys(""{F13}"")")
        TextBox2.Text = TextBox2.Text.Replace("F14", "WshShell.SendKeys(""{F14}"")")
        TextBox2.Text = TextBox2.Text.Replace("F15", "WshShell.SendKeys(""{F15}"")")
        TextBox2.Text = TextBox2.Text.Replace("F16", "WshShell.SendKeys(""{F16}"")")
        TextBox2.Text = TextBox2.Text.Replace("ESCAPE", "WshShell.SendKeys(""{ESC}"")")
        TextBox2.Text = TextBox2.Text.Replace("ESC", "WshShell.SendKeys(""{ESC}"")")
        TextBox2.Text = TextBox2.Text.Replace("REM", "'")
        TextBox2.Text = TextBox2.Text.Replace("CTRL", "WshShell.SendKeys(""^")
        TextBox2.Text = TextBox2.Text.Replace("TAB", "WshShell.SendKeys(""{TAB}"")")
        TextBox2.Text = TextBox2.Text.Replace("WshShell.SendKeys(""%{f4}"")", "WshShell.SendKeys(""%{F4}"")")
        My.Computer.FileSystem.WriteAllText(IO.Path.Combine(IO.Path.GetTempPath, "data.txt"), TextBox2.Text, False)
        For Each Line As String In File.ReadLines(IO.Path.Combine(IO.Path.GetTempPath, "data.txt"))
            If Line.StartsWith("STRING", StringComparison.InvariantCultureIgnoreCase) Then
                Line = "WshShell.SendKeys(""" & Line.Substring(6)
            End If
        Next
        TextBox2.Text = ""
        For Each Line As String In File.ReadLines(IO.Path.Combine(IO.Path.GetTempPath, "data.txt"))
            If Line.Contains("(""") Then

                TextBox2.AppendText(Line & """)" & Environment.NewLine)
            Else

                TextBox2.AppendText(Line & Environment.NewLine)
            End If
        Next
        TextBox2.Text = TextBox2.Text.Replace(""")"")", """)")
        TextBox2.Text = TextBox2.Text.Replace(""") "")", """)")

        ' Handle the REPEAT command. Created by Andrio
        Dim reader As StringReader = New StringReader(TextBox2.Text), outputLines As New List(Of String)
        Do
            Dim line = reader.ReadLine()
            If line Is Nothing Then Exit Do  ' ReadLine returns null when the end of the text is reached.
            Dim parts = line.TrimStart().Split(New Char() {" "c}, 2, StringSplitOptions.RemoveEmptyEntries)  ' TODO: Is this correct?
            Dim parameter = If(parts.Length = 1, "", parts(1).Trim())
            If parts(0) = "REPEAT" Then
                If outputLines.Count <= 2 Then Throw New FormatException("REPEAT is not valid on the first line.")
                Dim count As Integer
                If Not Integer.TryParse(parts(1), count) Then
                    ' TODO: Error handling
                    Throw New FormatException("Parameter to the REPEAT command must be an integer.")
                Else
                    ' Wrap the previous instruction in a For loop.
                    Dim counterVariable As Integer = AscW("i"c)
                    Dim targetLine As Integer = outputLines.Count - 1
                    Dim instructionLine As Integer = targetLine
                    Do While outputLines(instructionLine).EndsWith("Next")
                        instructionLine -= 1
                        targetLine -= 2
                        counterVariable += 1
                    Loop
                    ' Insert the loop.
                    For i As Integer = targetLine To outputLines.Count - 1
                        outputLines(i) = vbTab & outputLines(i)
                    Next
                    outputLines.Insert(targetLine, "For " & ChrW(counterVariable) & " = 1 To " & count)
                    outputLines.Add("Next")
                End If
            Else
                outputLines.Add(line)
            End If
        Loop
        TextBox2.Text = String.Join(Environment.NewLine, outputLines) & Environment.NewLine
        My.Computer.FileSystem.WriteAllText(IO.Path.Combine(IO.Path.GetTempPath, "data.txt"), TextBox2.Text, False)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = "Put your DuckyScript code here."
        TextBox2.Text = "Here will be your VBScript output."

        Timer1.Start()
        If File.Exists(IO.Path.Combine(IO.Path.GetTempPath, "data.txt")) = True Then
            File.Delete(IO.Path.Combine(IO.Path.GetTempPath, "data.txt"))
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox1.Paste()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.ResetText()
        TextBox2.Text = "Set WshShell = WScript.CreateObject(""WScript.Shell"")"
        TextBox2.Text = TextBox2.Text & Environment.NewLine & "set objShell = CreateObject(""Shell.Application"")"
    End Sub

    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If TextBox1.Text = ("") = True Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
        If TextBox1.Text = ("Put your DuckyScript code here.") = True Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        SaveFileDialog1.InitialDirectory = "C:\"
        SaveFileDialog1.Title = "Save vbs Files"
        SaveFileDialog1.CheckFileExists = False
        SaveFileDialog1.CheckPathExists = True
        SaveFileDialog1.DefaultExt = "VBS"
        SaveFileDialog1.Filter = "VBS files (*.vbs)|*.vbs"
        SaveFileDialog1.FilterIndex = 2
        SaveFileDialog1.RestoreDirectory = True
        SaveFileDialog1.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        MsgBox("People who worked hard on this program: Tim Wanders and Andrio")
    End Sub

    Private Sub SaveFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        Dim ANSIString() As Byte, MyEncoder As New System.Text.ASCIIEncoding()
        Dim fFile As IO.File, fStream As IO.FileStream
        Dim sSourceString As String = TextBox2.Text
        Dim sFilename As String = SaveFileDialog1.FileName
        ANSIString = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString))
        Try
            If fFile.Exists(sFilename) Then
                fFile.SetAttributes(sFilename, IO.FileAttributes.Normal)
                fFile.Delete(sFilename)
            End If
            fStream = fFile.OpenWrite(sFilename)
            fStream.Write(ANSIString, 0, ANSIString.Length)
            fStream.Close()
        Catch myException As Exception
            MessageBox.Show(myException.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub FormStudents_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        File.Delete(IO.Path.Combine(IO.Path.GetTempPath, "data.txt"))


    End Sub
    Private Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        If e.Control And e.KeyCode = Keys.A Then
            TextBox2.SelectAll()
        End If
    End Sub
End Class
