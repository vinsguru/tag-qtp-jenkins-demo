

'-------------------------------- Test1 ---------------------------------------------------

showMessageInJenkinsConsole "TC001", "Check If Google site is up and running"
SystemUtil.CloseProcessByName("iexplore.exe")
SystemUtil.Run "iexplore.exe", "https://www.google.com"

'Check if the search box is present
status = Browser("CreationTime:=0").Page("micclass:=Page").WebEdit("name:=q").Exist(5)
showTestCaseStatusInJenkinsConsole(status)

'-------------------------------- Test2 ---------------------------------------------------
showMessageInJenkinsConsole "TC002", "Check If Yahoo site is up and running"
SystemUtil.CloseProcessByName("iexplore.exe")
SystemUtil.Run  "iexplore.exe","https://www.yahoo.com"

'Check if the search box is present
status = Browser("CreationTime:=0").Page("micclass:=Page").WebEdit("name:=p").Exist(5)
showTestCaseStatusInJenkinsConsole(status)
