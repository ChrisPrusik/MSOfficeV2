# excel.getvalue

**Syntax:**

```G1ANT
excel.getvalue  row ‴‴ colindex ‴‴

```

or 

```G1ANT
excel.getvalue  row ‴‴ coname ‴‴

```

**Description:**

Command `excel.getvalue` allows to get value from specified cell. Please note, that you have to use either a colname or colindex argument.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`row`| "integer":{TOPIC-LINK+integer}| yes |  | cell's row number |
|`colindex` or `colname`|  "integer":{TOPIC-LINK+integer} or "string":{TOPIC-LINK+string}| yes |  | `colindex` - cell's column number, `colname` - cell's column name |
|`result`| "variable":{TOPIC-LINK+variable}| no | "♥result":{TOPIC-LINK+common-arguments} | name of variable where cell's value will be stored |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump`| "label":{TOPIC-LINK+label}| no |  | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

```G1ANT
excel.getvalue row 11 colindex 3 result ♥var1
dialog ♥var1

```

In this example the value from specified cell is shown in a dialog box.

!{IMAGE-LINK+excel-getvalue}! 

**Example 2:**

```G1ANT
excel.open
excel.setvalue value ‴Remember, remember!‴ row 1 colname A
excel.setvalue value ‴the fifth of November‴ row 2 colname A
excel.getvalue row 1 colname A result ♥guy
dialog ♥guy

```

**Example 3:**

```G1ANT
excel.open
keyboard ‴Remember, remember!‴
keyboard ⋘DOWN⋙
keyboard ‴The fifth of November⋘DOWN⋙The Gunpowder treason and plot‴
excel.getvalue row 1 colname A result ♥guy
dialog ♥guy

```