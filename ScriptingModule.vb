Imports System.IO
Imports System.Reflection 'compilers
Imports System.CodeDom
Imports System.CodeDom.Compiler
Imports Microsoft.VisualBasic
Imports Microsoft.CSharp 'Imports System.Security.Permissions
Imports System.Threading


'<Assembly: SecurityPermissionAttribute(SecurityAction.RequestMinimum, _
'  ControlThread:=True)> 


Public Module scripting
    Public CurrentLanguage As lang = lang.cs

    'Python, C#, VB, Matlab
    '-------------- generic -------------------
    'Dim pythonOutput = ""
    Public pythOut$
    Public newPythOutp$
    Public PythIn As New List(Of Integer)
    Dim pythonProcess As Process
    Dim psinfo As ProcessStartInfo
    Dim pythonTh As Thread

    <MTAThread()>
    Sub runPython()



        ' On Error Resume Next 'Make sure it works

        pythOut = ""
        newPythOutp = ""

        psinfo.RedirectStandardInput = True
        psinfo.RedirectStandardOutput = True
        psinfo.RedirectStandardError = True
        psinfo.CreateNoWindow = True
        psinfo.UseShellExecute = False





        'Threading.Thread.CurrentThread.Priority = ThreadPriority.Highest
        If pythonProcess Is Nothing OrElse pythonProcess.HasExited Then
            pythonProcess = New Process()
            pythonProcess.StartInfo = psinfo
        Else
            pythonProcess.Kill()
            pythonProcess = New Process()
            pythonProcess.StartInfo = psinfo
            'pythonProcess.StandardInput.WriteLine(ChrW(3)) 'Cancel
            'pythonProcess.StandardInput.WriteLine(" execfile({0})", Replace(psinfo.Arguments, "-i ", ""))

        End If
        pythOut &= "Started Python ! ..." & vbNewLine
        '  pythonProcess.PriorityClass = ProcessPriorityClass.High


        AddHandler pythonProcess.OutputDataReceived, AddressOf dataPythRec

        AddHandler pythonProcess.ErrorDataReceived, AddressOf dataPythRec


        pythonProcess.Start()
        pythonProcess.BeginErrorReadLine()
        pythonProcess.BeginOutputReadLine()


        ' While Not pythonProcess.HasExited



        '  If pythonProcess.StandardError.Peek() > 0 Then
        'newPythOutp &= vbNewLine & "-- Error:" & vbNewLine & pythonProcess.StandardError.ReadLine
        '  End If

        'pythOut &= newPythOutp
        '   End While



        'Threading.Thread.CurrentThread.Priority = ThreadPriority.Normal

    End Sub

    Sub dataPythRec(ByVal sender As Object, ByVal e As DataReceivedEventArgs)

        '   If pythonProcess.StandardOutput.Peek Then
        ' newPythOutp &= pythonProcess.StandardOutput.ReadLine & vbNewLine
        '  End If
        newPythOutp &= e.Data & vbNewLine
        pythOut &= newPythOutp


        'If done, launch the script
        If InStr(newPythOutp, "@robotDone") <> 0 Then
            Dim foo = Split(pythOut, "@robot")(UBound(Split(pythOut, "@robot")) - 1)
            ' scripting.parseBestRouteAndHistory(lang.python, foo)

        End If

    End Sub



    '--------------------- MATLAB ----------------------
    Dim matlabIsLaunchedOnce = False
    Sub LocateMatlab()
        matlabIsLaunchedOnce = GetSetting(Application.ExecutablePath.Split("\").Last, "MAIN", "matlab", False)

        If matlabIsLaunchedOnce = False Then
            MsgBox("We are going to register MATLAB's built-in COM" & vbNewLine & "Please make sure to accept running as admin for the next dialog", MsgBoxStyle.Information)
            MsgBox("If it doesn't work, we (shamelessly) ask you to restart this program. Thank you :) ", MsgBoxStyle.Information)
            Dim p As New Process()
            Dim psinfo As New ProcessStartInfo
            psinfo.FileName = "matlab"
            psinfo.Verb = "runas"
            psinfo.WindowStyle = ProcessWindowStyle.Normal
            psinfo.UseShellExecute = True
            psinfo.Arguments = "/regserver"

            ' psinfo.CreateNoWindow = True
            p = Process.Start(psinfo)
            p.WaitForExit()

            SaveSetting(Application.ExecutablePath.Split("\").Last, "MAIN", "matlab", True)

        End If

    End Sub
    Dim MatLab As Object = Nothing
    Dim matlabAxTriesToCreateObject = 0
    Function ExecuteMatlab(ByRef cont$)
        LocateMatlab()

        'Dim MReal(,), MImag(,) As Double

create_obj:
        Try
            matlabAxTriesToCreateObject += 1
            If MatLab Is Nothing Then MatLab = CreateObject("Matlab.Application")

        Catch ex As Exception

            If matlabAxTriesToCreateObject < 4 Then
                GoTo create_obj
            Else
                Throw New Exception("Couldn't create ActiveX object")
            End If


        End Try

        'Calling MATLAB function from VB
        'Assuming solve_bvp exists at specified location
        ' Result = MatLab.Execute("cd d:\matlab\work\bvp")
        ' Result = MatLab.Execute("solve_bvp")

        'Executing other MATLAB commands
        ExecuteMatlab = MatLab.Execute(cont)

        Dim rawBestRoute, rawHistory As New Object

        Try
            rawBestRoute = MatLab.GetVariable("bestRoute", "base")
        Catch ex As Exception

        End Try

        Try
            rawHistory = MatLab.GetVariable("history", "base")
        Catch ex As Exception

        End Try
        ' scripting.parseBestRouteAndHistory(lang.matlab, rawBestRoute, rawHistory)



        'Result = MatLab.Execute("a = [1 2 3 4; 5 6 7 8]")
        'Result = MatLab.Execute("b = a + a ")
        'Bring matrix b into VB program
        'MatLab.GetFullMatrix("b", "base", MReal, MImag)
        'MsgBox(MatLab.GetFullMatrix("bestRoute", "base", MReal, MImag))
        'Return Result
    End Function

    '------------------- VB :) -----------------------
    Public Function ExecuteVB(ByRef cont$)


        ' Create "code" literal to pass to the compiler.  
        '  
        ' Notice the <% = input % > where the code read from the text file (Code.txt)   
        ' is inserted into the code fragment. 


        Dim code = <code>  
                     Imports  System   
                     Imports  System.Drawing
                     Imports System.Collections.Generic
                     Imports  Plan_the_road_for_the_bot.dotNETemu.dotNETscriptingPlateform
                      Namespace dotNETemu
                       Public Class ExecClass
                           Inherits  Plan_the_road_for_the_bot.dotNETemu.dotNETscriptingPlateform
                          Public Sub executeUnsafe()  
                               <%= cont %>  
                           End Sub  
                       End Class  
                      End Namespace
                   </code>

        ' Create the VB.NET compiler.  
        Dim x = New CSharpCodeProvider
        Dim vbProv = New VBCodeProvider()
        ' Create parameters to pass to the compiler.  
        Dim vbParams = New CompilerParameters()
        ' Add referenced assemblies.  
        vbParams.ReferencedAssemblies.Add("mscorlib.dll")
        vbParams.ReferencedAssemblies.Add("System.dll")
        vbParams.ReferencedAssemblies.Add(Application.ExecutablePath)
        vbParams.ReferencedAssemblies.Add("System.Drawing.dll")
        vbParams.GenerateExecutable = False
        ' Ensure we generate an assembly in memory and not as a physical file.  
        vbParams.GenerateInMemory = True

        ' Compile the code and get the compiler results (contains errors, etc.)  
        Dim compResults = vbProv.CompileAssemblyFromSource(vbParams, code.Value)

        ' Check for compile errors  
        If compResults.Errors.Count > 0 Then

            ' Show each error.  
            Dim str$ = ""
            Dim foo$, bar$, foobar$
            For Each er In compResults.Errors
                foo = Split(er.ToString(), ".vb(")(1)
                bar = Split(foo, ",")(0) - 7

                foobar = Split(foo, ")")(0)


                foo = Mid(foo, InStr(foo, ")") + 2)
                str &= "Ln: " & bar & ", Col: " & foobar & foo & vbNewLine
            Next
            MsgBox("Error occurred in your program:" & vbNewLine & str, MsgBoxStyle.Critical)


            Return False

        Else

            ' Create instance of the temporary compiled class.  
            Dim obj As Object = compResults.CompiledAssembly.CreateInstance("dotNETemu.ExecClass")
            ' obj.Grid = mainGrid.Grid
            ' obj.StartGrid = mainGrid.StartGrid
            ' obj.FinishGrid = mainGrid.FinishGrid
            ' An array of object that represent the arguments to be passed to our method (UpdateText).
            ' Execute the method by passing the method name and arguments.  
            Dim t As Type = obj.GetType().InvokeMember("executeUnsafe", BindingFlags.InvokeMethod, Nothing, obj, Nothing)
            ' bestRoute.Clear()
            ' bestRoute.AddRange(obj.bestRoute.toArray())
            ' HistoryOfExploration.Clear()
            ' HistoryOfExploration.AddRange(obj.history.toArray())

        End If


        Return True
    End Function
    '---------------- C# --------------------------------
    Function ExecuteCS(ByRef cont$) As Boolean

        ' Create "code" literal to pass to the compiler.  
        '  
        ' Notice the <% = input % > where the code read from the text file (Code.txt)   
        ' is inserted into the code fragment. 


        Dim code = <code>  
                     using  System;
                     using  System.Drawing;
                     using System.Collections.Generic;
                      namespace dotNETemu {
                       public class ExecClass :  Plan_the_road_for_the_bot.dotNETemu.dotNETscriptingPlateform {
                          public void executeUnsafe() {
                               <%= cont %>  
                           }
                       } 
                     }
                   </code>

        ' Create the VB.NET compiler.  
        Dim csProv = New CSharpCodeProvider()
        ' Create parameters to pass to the compiler.  
        Dim csParams = New CompilerParameters()
        ' Add referenced assemblies.  
        csParams.ReferencedAssemblies.Add("mscorlib.dll")
        csParams.ReferencedAssemblies.Add("System.dll")
        csParams.ReferencedAssemblies.Add(Application.ExecutablePath)
        csParams.ReferencedAssemblies.Add("System.Drawing.dll")
        csParams.GenerateExecutable = False
        ' Ensure we generate an assembly in memory and not as a physical file.  
        csParams.GenerateInMemory = True

        ' Compile the code and get the compiler results (contains errors, etc.)  
        Dim compResults = csProv.CompileAssemblyFromSource(csParams, code.Value)

        ' Check for compile errors  
        If compResults.Errors.Count > 0 Then

            Dim str$ = ""
            Dim foo$, bar$, foobar$
            For Each er In compResults.Errors
                foo = Split(er.ToString(), ".cs(")(1)
                bar = Split(foo, ",")(0) - 5

                foobar = Split(foo, ")")(0)


                foo = Mid(foo, InStr(foo, ")") + 2)
                str &= "Ln: " & bar & ", Col: " & foobar & foo & vbNewLine
            Next
            MsgBox("Error occurred in your program:" & vbNewLine & str, MsgBoxStyle.Critical)


            Return False


        Else

            ' Create instance of the temporary compiled class.  
            Dim obj As Object = compResults.CompiledAssembly.CreateInstance("dotNETemu.ExecClass")
            '' obj.Grid = mainGrid.Grid
            ' obj.StartGrid = mainGrid.StartGrid
            ' obj.FinishGrid = mainGrid.FinishGrid
            ' An array of object that represent the arguments to be passed to our method (UpdateText).
            ' Execute the method by passing the method name and arguments.  
            Dim t As Type = obj.GetType().InvokeMember("executeUnsafe", BindingFlags.InvokeMethod, Nothing, obj, Nothing)
            ' bestRoute.Clear()
            ' bestRoute.AddRange(obj.bestRoute.toArray())
            ' ' HistoryOfExploration.Clear()
            ' HistoryOfExploration.AddRange(obj.history.toArray())

        End If
    End Function
    '------------------- C -----------------------
    Sub LocateC()
        If Not GetSetting(Application.ExecutablePath.Split("\").Last, "MAIN", "C", False) Then


        End If

        If matlabIsLaunchedOnce = False Then
            MsgBox("We are going to register MATLAB's built-in COM" & vbNewLine & "Please make sure to accept running as admin for the next dialog", MsgBoxStyle.Information)
            MsgBox("If it doesn't work, we (shamelessly) ask you to restart this program. Thank you :) ", MsgBoxStyle.Information)
            Dim p As New Process()
            Dim psinfo As New ProcessStartInfo
            psinfo.FileName = "matlab"
            psinfo.Verb = "runas"
            psinfo.WindowStyle = ProcessWindowStyle.Normal
            psinfo.UseShellExecute = True
            psinfo.Arguments = "/regserver"

            ' psinfo.CreateNoWindow = True
            p = Process.Start(psinfo)
            p.WaitForExit()

            SaveSetting(Application.ExecutablePath.Split("\").Last, "MAIN", "matlab", True)

        End If

    End Sub




    Enum lang
        python = 0
        cs = 1
        vb = 2
        matlab = 3
        java = 4
        c = 5
        cpp = 6
    End Enum

End Module
