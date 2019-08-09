#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
;#AutoIt3Wrapper_usex64=n
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Toast.au3>
#include <JanisonTestAutomationReports.au3>

Local $app_name = "Janison Test Automation Reports"

; Authentication

Local $ini_filename = @ScriptDir & "\" & $app_name & ".ini"

Local $testrail_username = IniRead($ini_filename, "main", "username", "")
Local $testrail_encrypted_password = IniRead($ini_filename, "main", "password", "")
_TestRailAuthenticationWithToast($app_name, "https://janison.testrail.com", $testrail_username, $testrail_encrypted_password, $ini_filename)

Local $confluence_username = IniRead($ini_filename, "main", "confluenceusername", "")
Local $confluence_encrypted_password = IniRead($ini_filename, "main", "confluencepassword", "")
_ConfluenceAuthenticationWithToast($app_name, "https://janisoncls.atlassian.net", $confluence_username, $confluence_encrypted_password, $ini_filename)

; Insights

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
FileDelete(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Exec(-1, "PRAGMA synchronous = OFF;")		; this should speed up DB transactions
_SQLite_Exec(-1, "CREATE TABLE SavedRuntime (CreatedOn,Test,SavedRuntime integer);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE SavedCases (Test,SavedCases integer);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE Bug (Key,Created,Priority);") ; CREATE a Table

$storage_format = '<a href=\"https://github.com/seanhaydongriffin/Insights-Test-Automation-Reports/releases/download/v0.1/Insights.Test.Automation.Reports.portable.exe\">Click to update Insights reports</a><br /><br />'

ManualRunTimeSavedReport($app_name, 49, "Insights ", " SIT")
ManualTestCasesSavedReport($app_name, 49, "Insights - ")
DefectsRaisedReport($app_name, "project = INS AND issuetype = Bug AND labels = Automation AND labels = Insights", 1)
DefectsListReport("project = INS AND issuetype = Bug AND labels = Automation AND labels = Insights")

_Toast_Show(0, $app_name, "Uploading reports to confluence", -300, False, True)
Update_Confluence_Page("https://janisoncls.atlassian.net", $jira_username, $jira_password, "JAST", "469664600", "469664606", "Insights Test Automation Reports", $storage_format)
_SQLite_Close()
_SQLite_Shutdown()

; Learning

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
FileDelete(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Exec(-1, "PRAGMA synchronous = OFF;")		; this should speed up DB transactions
_SQLite_Exec(-1, "CREATE TABLE SavedRuntime (CreatedOn,Test,SavedRuntime integer);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE SavedCases (Test,SavedCases integer);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE Bug (Key,Created,Priority);") ; CREATE a Table

$storage_format = '<a href=\"https://github.com/seanhaydongriffin/Learning-Test-Automation-Reports/releases/download/v0.1/Learning.Test.Automation.Reports.portable.exe\">Click to update Learning reports</a><br /><br />'

ManualRunTimeSavedReport($app_name, 50, "", " ST")
ManualTestCasesSavedReport($app_name, 50, "")
DefectsRaisedReport($app_name, "project = CLS AND issuetype = Bug AND labels = Automation", 1)
DefectsListReport("project = CLS AND issuetype = Bug AND labels = Automation")

_Toast_Show(0, $app_name, "Uploading reports to confluence", -300, False, True)
Update_Confluence_Page("https://janisoncls.atlassian.net", $jira_username, $jira_password, "JAST", "474087425", "473989127", "Learning Test Automation Reports", $storage_format)
_SQLite_Close()
_SQLite_Shutdown()

; RMS

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
FileDelete(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Exec(-1, "PRAGMA synchronous = OFF;")		; this should speed up DB transactions
_SQLite_Exec(-1, "CREATE TABLE SavedRuntime (CreatedOn,Test,SavedRuntime integer);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE SavedCases (Test,SavedCases integer);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE Bug (Key,Created,Priority);") ; CREATE a Table

$storage_format = '<a href=\"https://github.com/seanhaydongriffin/RMS-Test-Automation-Reports/releases/download/v0.1/RMS.Test.Automation.Reports.portable.exe\">Click to update RMS reports</a><br /><br />'

ManualRunTimeSavedReport($app_name, 47, "", " SIT")
ManualTestCasesSavedReport($app_name, 47, "")
DefectsRaisedReport($app_name, "project = RMS AND issuetype = Bug AND labels = Automation", 5)
DefectsListReport("project = RMS AND issuetype = Bug AND labels = Automation")

_Toast_Show(0, $app_name, "Uploading reports to confluence", -300, False, True)
Update_Confluence_Page("https://janisoncls.atlassian.net", $jira_username, $jira_password, "JAST", "386826341", "468615262", "RMS Test Automation Reports", $storage_format)
_SQLite_Close()
_SQLite_Shutdown()

; UP

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
FileDelete(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Exec(-1, "PRAGMA synchronous = OFF;")		; this should speed up DB transactions
_SQLite_Exec(-1, "CREATE TABLE SavedRuntime (CreatedOn,Test,SavedRuntime integer);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE SavedCases (Test,SavedCases integer);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE Bug (Key,Created,Priority);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE SubTask (Key,Summary,Description,StoryKey,FixVersion,Status,EstimatedTime,TimeSpent,ProgressPercent,Environment);") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE SubTaskStateHistory (Key,Date,Status);") ; CREATE a Table

$storage_format = '<a href=\"https://github.com/seanhaydongriffin/UP-Test-Automation-Reports/releases/download/v0.1/UP.Test.Automation.Reports.portable.exe\">Click to update UP reports</a><br /><br />'

ManualRunTimeSavedReport($app_name, 49, "UP ", " SIT")
ManualTestCasesSavedReport($app_name, 49, "UP - ")
DefectsRaisedReport($app_name, "project = INS AND issuetype = Bug AND labels = Automation AND labels = UP", 1)
DefectsListReport("project = INS AND issuetype = Bug AND labels = Automation AND labels = UP")
CoreCoverageReport($app_name)

_Toast_Show(0, $app_name, "Uploading reports to confluence", -300, False, True)
Update_Confluence_Page("https://janisoncls.atlassian.net", $jira_username, $jira_password, "JAST", "469172330", "469172334", "UP Test Automation Reports", $storage_format)
_SQLite_Close()
_SQLite_Shutdown()







; Shutdown

_JiraShutdown()
_Toast_Show(0, $app_name, "Done. Refresh the page in Confluence.", -3, False, True)
Sleep(3000)
