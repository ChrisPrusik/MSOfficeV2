# excel.open

**Syntax:**

```G1ANT
excel.open

```

**Description:**

Command `excel.open` allows to open a new Excel instance.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`path`| "string":{TOPIC-LINK+string}| no |  | path of a file that has to be opened, if not specified, excel will be opened anyway |
|`inbackground`| "bool":{TOPIC-LINK+boolean}| no | false | defines whether Excel opens in the background  |
|`sheet`| "string":{TOPIC-LINK+string}| no |  | sheet name to be activated |
|`result`| "variable":{TOPIC-LINK+variable}| no | "♥result":{TOPIC-LINK+common-arguments} | name of variable where number of currently opened Excel processes is stored, it can be used later on with command `excel.switch` |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "integer":{TOPIC-LINK+integer}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump`| "label":{TOPIC-LINK+label}| no |  | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

Here document.xlsx is being open with activated sheet_number_223 and the result is assigned to ♥excelId1 variable.

```G1ANT
excel.open path ‴C:\programs\document.xlsx‴ sheet ‴sheet_number_223‴ result ♥excelId1 

```

!{IMAGE-LINK+excel-open}! 

**Example 2:**

Below you can see how easy it is to open a new Excel file:

```G1ANT
excel.open

```

**Example 3:**

If you choose that `excel.open` command works in background, you will not notice any action, but G1ANT.Robot will perform according to the script.

```G1ANT
excel.open inbackground true
excel.setvalue value ‴Random Text‴ row 1 colname A
excel.save path ‴C:\Tests\test.xlsx‴
excel.close 

```

**Example 4:**

In this example you can create Sheet2 while opening Excel.

```G1ANT
excel.open path ‴C:\Tests\test.xlsx‴ sheet Sheet2

```