﻿Imports System.IO
Public Class Form1
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

        TextBox1.Text = ("")


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox2.Text = "set wshshell = wscript.createobject(""wscript.shell"")"
        TextBox2.Text = TextBox2.Text & Environment.NewLine & "set objshell = createobject(""shell.application"")"
        If File.Exists("data.txt") = True Then
            File.Delete("data.txt")
        End If

        ' Get the current lines.
        Dim lines() As String = TextBox2.Lines
        ' At line 0 Insert the "Hello World" string at position 2

        ' Reassign the lines property.
        TextBox2.Lines = lines



        TextBox2.Text = TextBox2.Text & Environment.NewLine & TextBox1.Text
        TextBox2.CharacterCasing = CharacterCasing.Upper
        TextBox2.Text = TextBox2.Text.Replace("^", "{^}")
        TextBox2.Text = TextBox2.Text.Replace("(", "{(}")
        TextBox2.Text = TextBox2.Text.Replace(")", "{)}")
        TextBox2.Text = TextBox2.Text.Replace("'", "{'}")
        TextBox2.Text = TextBox2.Text.Replace("|", "{|}")
        TextBox2.Text = TextBox2.Text.Replace("%", "{%}")
        TextBox2.Text = TextBox2.Text.Replace("""", """""")
        TextBox2.Text = TextBox2.Text.Replace("SET WSHSHELL = WSCRIPT.CREATEOBJECT{(}""""WSCRIPT.SHELL""""{)}", "set wshshell = wscript.createobject(""wscript.shell"")")
        TextBox2.Text = TextBox2.Text.Replace("SET OBJSHELL = CREATEOBJECT{(}""""SHELL.APPLICATION""""{)}", "set objshell = createobject(""shell.application"")")
        'TextBox2.Text = TextBox2.Text.Replace("STRING", "WshShell.SendKeys(""")
        If TextBox2.Text.StartsWith("STRING", StringComparison.CurrentCultureIgnoreCase) Then
            TextBox2.Text = "WshShell.SendKeys(""" & TextBox2.Text.Substring(6)
        End If

        TextBox2.Text = TextBox2.Text.Replace("ENTER", "WshShell.SendKeys(""{ENTER}"")")
        TextBox2.Text = TextBox2.Text.Replace("ALT", "WshShell.SendKeys(""%"")")
        TextBox2.Text = TextBox2.Text.Replace("GUI R", "objShell.FileRun")
        TextBox2.Text = TextBox2.Text.Replace("GUI", "WshShell.SendKeys(""^{Esc}"")")
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
        TextBox2.Text = TextBox2.Text.Replace("TAB", "WshShell.SendKeys(""{TAB}"")")
        My.Computer.FileSystem.WriteAllText("data.txt", TextBox2.Text, False)
        For Each Line As String In File.ReadLines("data.txt")


            If Line.StartsWith("STRING", StringComparison.InvariantCultureIgnoreCase) Then
                Line = "WshShell.SendKeys(""" & Line.Substring(6)
            End If

        Next







        'Dim Line As String = File.ReadLines("data.txt")
        'If TextBox2.Lines.Contains("(""") Then
        'TextBox2.AppendText(Line & """)" & Environment.NewLine)


        'End If


        'Dim index As Integer
        'Dim customerName As String = "KEYS.SENDKEYS("" "

        'For i As Integer = 1 To TextBox2.Lines.Length - 1
        'If TextBox2.Lines(i - 1).Contains("Customer") Then
        'customerName = TextBox2.Lines(i)
        'Exit For
        'End If
        'Next


















        TextBox2.CharacterCasing = CharacterCasing.Lower

        TextBox2.Text = ""

        For Each Line As String In File.ReadLines("data.txt")


            If Line.Contains("(""") Then

                TextBox2.AppendText(Line & """)" & Environment.NewLine)
            Else

                TextBox2.AppendText(Line & Environment.NewLine)
            End If

        Next





        TextBox2.Text = TextBox2.Text.Replace(""")"")", """)")
        TextBox2.Text = TextBox2.Text.Replace(""") "")", """)")
        My.Computer.FileSystem.WriteAllText("data.txt", TextBox2.Text, False)



    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = "Put your DuckyScript code here."
        TextBox2.Text = "set wshshell = wscript.createobject(""wscript.shell"")"
        TextBox2.Text = TextBox2.Text & Environment.NewLine & "set objshell = createobject(""shell.application"")"
        Timer1.Start()
        If File.Exists("data.txt") = True Then
            File.Delete("data.txt")
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
        If TextBox2.Text = ("") = True Then
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
        'My.Computer.FileSystem.WriteAllText("script.vbs", TextBox2.Text, False)
        Dim ANSIString() As Byte, MyEncoder As New System.Text.ASCIIEncoding()
        Dim fFile As IO.File, fStream As IO.FileStream

        ' Set the source string and filename
        Dim sSourceString As String = TextBox2.Text
        Dim sFilename As String = "script.vbs"

        ' Store the ANSI encoded string in ANSIString
        ANSIString = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString))

        Try

            ' Delete the file if it already exists
            If fFile.Exists(sFilename) Then
                fFile.SetAttributes(sFilename, IO.FileAttributes.Normal)
                fFile.Delete(sFilename)
            End If

            ' Output the bytes
            fStream = fFile.OpenWrite(sFilename)
            fStream.Write(ANSIString, 0, ANSIString.Length)
            fStream.Close()

        Catch myException As Exception
            ' Show a message box if an error occurs
            MessageBox.Show(myException.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        MsgBox("People who worked hard on this program: Tim Wanders and Andrio")
    End Sub
End Class
