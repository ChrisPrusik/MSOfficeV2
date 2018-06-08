# excel.getformula

**Syntax:**

```G1ANT
excel.getformula  row ‴‴ colindex ‴‴

```

or 

```G1ANT
excel.getformula  row ‴‴ colname ‴‴

```

**Description:**

Command `excel.getformula` allows to get formula from specified cell. 

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`row`| "integer":{TOPIC-LINK+integer}| yes |  | cell's row number |
|`colindex` or `colname`|  "integer":{TOPIC-LINK+integer} or "string":{TOPIC-LINK+string}| yes |  | `colindex` - cell's column number, `colname` - cell's column name |
|`result`| "variable":{TOPIC-LINK+variable}| no | "♥result":{TOPIC-LINK+common-arguments}  | name of variable where command's result will be stored |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no |  | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

```G1ANT
excel.getformula row 11 colindex 3
dialog  ♥result

```

**Example 2:**

```G1ANT
excel.getformula row 11 colname ‴3‴ result ♥var1
dialog  ♥var1

```

In this example formula used in the specified cell is shown in a dialog.

!{IMAGE-LINK+excel-getformula}! 

**Example 3:**

```G1ANT
excel.open path ‴C:\Tests\test.xlsx‴
excel.activatesheet name ‴Sheet3‴
excel.getformula row 1 colname ‴A‴ result ♥formula
dialog ♥formula

```