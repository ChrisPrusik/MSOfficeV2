# excel.duplicaterow

**Syntax:**

```G1ANT
excel.duplicaterow  source ‴‴ destination ‴‴

```

**Description:**

Command `excel.duplicaterow` allows to copy specified row into a specified place.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`source`| "integer":{TOPIC-LINK+integer}| yes |  | source row's number |
|`destination`| "integer":{TOPIC-LINK+integer}| yes |  | destination row's number |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no |  | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no | | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

```G1ANT
excel.duplicaterow source 2 destination 5

```

**Example 2:**

```G1ANT
excel.open path ‴C:\Tests\excelTest.xlsx‴ result ♥excelHandle1
excel.duplicaterow source 3 destination 4
excel.save path ‴C:\Tests\excelTest.xlsx‴
excel.close

```

!{IMAGE-LINK+2017-11-21-excel-duplicaterow}! 

In this case `excel.duplicaterow` will overwrite the 4th row with the value from 3rd row.