Dim strBasePath : strBasePath = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))

'Kill the processes
ConsoleOutputBlankLine(1)
Call KillProcess("UFT.exe")
Call KillProcess("QtpAutomationAgent.exe")
Call KillProcess("iexplore.exe")

'Create QTP object
ConsoleOutputBlankLine(1)
Set QTP = CreateObject("QuickTest.Application")
ConsoleOutput("Launching QTP Application")
QTP.Launch
QTP.Visible = TRUE

'Open QTP Test
ConsoleOutput("Opening Test....")
QTP.Open strBasePath & "\tagTest", TRUE 'Set the QTP test path
  
'Set Result location
Set qtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")
qtpResultsOpt.ResultsLocation = strBasePath & "\output" 'Set the results location
  
'Create Environment Variables to pass the results from QTP to runner.vbs
QTP.Test.Environment.Value("JenkinsFlag") = "N"
QTP.Test.Environment.Value("JenkinsTestCaseDescription") = ""
QTP.Test.Environment.Value("JenkinsTestCaseResult") = ""
  
'Set the Test Parameters
Set pDefColl = QTP.Test.ParameterDefinitions
Set qtpParams = pDefColl.GetParameters()
 
'Set the value for test environment through command line
On Error Resume Next
qtpParams.Item("env").Value = LCase(WScript.Arguments.Item(0))
On Error GoTo 0
 
'Attach the vbs files to the test
QTP.Test.Settings.Resources.Libraries.RemoveAll   'Remove everything
QTP.Test.Settings.Resources.Libraries.Add("..\functions\TestInitialize.vbs")
QTP.Test.Settings.Resources.Libraries.Add("..\functions\Jenkins-Output.vbs")
  
'Run QTP test
ConsoleOutput("Starting to run....")
QTP.Test.Run qtpResultsOpt, FALSE, qtpParams
 
'Write the result in the console
ConsoleOutputBlankLine(2)
While QTP.Test.isRunning
    If QTP.Test.Environment.Value("JenkinsFlag") = "Y" Then
        QTP.Test.Environment.Value("JenkinsFlag") = "N"
 
        ' Show TC ID and Description
        WScript.StdOut.Write Time() & " :: " &  QTP.Test.Environment.Value("JenkinsTestCaseNumber") & " - " & QTP.Test.Environment.Value("JenkinsTestCaseDescription") & " - "
 
        'Wait till the test is executed & result is updated
        While (QTP.Test.Environment.Value("JenkinsTestCaseResult") = "" AND QTP.Test.isRunning)
                WScript.Sleep 1000
        Wend
 
        'Show the Result
        WScript.StdOut.WriteLine QTP.Test.Environment.Value("JenkinsTestCaseResult")
    End If
    WScript.Sleep 1000
Wend
ConsoleOutputBlankLine(2)
'Close QTP
ConsoleOutput("Execution Completed Successfully!!!!!!!!!!")
QTP.Quit


Sub ConsoleOutput(ByVal MessageToBeDisplayed)
	WScript.StdOut.WriteLine Time() & " :: " & MessageToBeDisplayed
End Sub

Sub ConsoleOutputBlankLine(ByVal intNo)
	WScript.StdOut.WriteBlankLines(intNo)
End Sub

Sub KillProcess(ByVal ProcessName)
	
	On Error Resume Next
	
	Dim objWMIService : Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Dim colProcesses : Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" &  ProcessName & "'")

	ConsoleOutput("Terminating Process : " & ProcessName)
	
	For Each objProcess in colProcesses
		intTermProc = objProcess.Terminate
	Next
	
	On Error GoTo 0
	
End Sub