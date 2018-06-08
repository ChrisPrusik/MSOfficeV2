# excel.runmacro

**Syntax:**

```G1ANT
excel.runmacro  name ‴‴ 

```

**Description:**

Command `excel.runmacro` allows to run macro in currently active excel program. 

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`name`| "string":{TOPIC-LINK+string}| yes |  | name of macro that is defined in a workbook |
|`args`| "string":{TOPIC-LINK+string}| no| | comma separated arguments that will be passed to macro |
|`result`| "variable":{TOPIC-LINK+variable}| no |  "♥result":{TOPIC-LINK+common-arguments} | name of variable where macro's return value will be stored |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no | | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

```G1ANT
excel.runmacro name ‴macro1‴ args ‴1,2‴

```


In this example a previously defined "macro1" macro is run. The first column is multiplied by argument 1 (1) and copied to column D, and then it is multiplied by argument 2 (2) and copied to column E.

!{IMAGE-LINK+excel-macro}! 

**Example 2:**

```G1ANT
excel.open path ‴C:\Tests\macros.xlsm‴
excel.runmacro name ‴multiply‴

```