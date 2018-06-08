# excel.save

**Syntax:**

```G1ANT
excel.save 

```

**Description:**

Command `excel.save` allows to save currently active Excel workbook.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`path`| "string":{TOPIC-LINK+string}| no |  | saving file path; if not specified, G1ANT.Robot will try to save the file under the path it was loaded from. If current Excel application was not opened with specified path, it will go into exception handling.|
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no | | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

In this example the currently active Excel workbook is saved as doc1.xlsx

```G1ANT
excel.save path ‴C:\Documents\doc1.xlsx‴

```

**Example 2:**

```G1ANT
excel.open result ♥excelHandle1
excel.setvalue value ‴random text‴ row 1 colname ‴A‴
excel.setvalue value ‴54642 5854 22‴ row 1 colname ‴B‴
excel.save path ‴C:\Tests\excelTest.xlsx‴
excel.close

```