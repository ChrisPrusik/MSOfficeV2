# excel.setvalue

**Syntax:**

```G1ANT
excel.setvalue  value ‴‴  row ‴‴  colindex ‴‴ 

```

or 

```G1ANT
 excel.setvalue  value ‴‴  row ‴‴  colname ‴‴ 

```

**Description:**

Command `excel.setvalue` allows to set value into specified cell.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`value`| "string":{TOPIC-LINK+string}| yes|  | any value to be set inside of a specified cell |
|`row`| "integer":{TOPIC-LINK+integer}| yes |  | cell's row number |
|`colindex` or `colname`|  "integer":{TOPIC-LINK+integer} or "string":{TOPIC-LINK+string}| yes |  | `colindex` - cell's column number, `colname` - cell's column name |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no | | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

```G1ANT
excel.setvalue value ‴someText‴ row 2 colname ‴2‴

```

```G1ANT
excel.setvalue value ‴=A1+A5‴ row 6 colname ‴1‴

```

In this example a value and a formula as defined are inserted into specified cells.

!{IMAGE-LINK+excel-setval}! 

**Example 2:**

In order to use `excel.setvalue` command, we first need to have an Excel file opened. We can use `excel.open` command. Then we have to activate certain Sheet, as usually the Excel files we might be working on consist of many Sheets. `excel.activatesheet` command needs **name** argument with name of a sheet given as value.
The next step would be to use `excel.getvalue` command. It enables to read certain content of a cell and remember it in a **result** argument as a variable. 

```G1ANT
excel.open path ‴C:\Tests\excelTest.xlsx‴ result ♥excelHandle1
excel.activatesheet name ‴Sheet1‴
excel.getvalue row 2 colname ‴B‴ result ♥valueColB
excel.setvalue value ‴the best pony is shetland pony‴ row 3 colname ‴A‴ if ♥valueColB
excel.save path ‴C:\Tests\excelTest.xlsx‴
excel.close

```

In our script, while using `row 2 colname ‴B‴ result ♥valueColB` line, we can read the content of 2B cell- it is 'TRUE'.

!{IMAGE-LINK+2017-12-21-excel-setvalue}! 

Then, we can check the condition and inject certain new value only if the condition is fulfilled. In our case the condition we are checking is: `if ♥valueColB`. By default it means **if ♥valueColB = 'TRUE'**. It does equal 'TRUE', so G1ANT.Robot is injecting ‴the best pony is shetland pony‴ in 3A cell.

!{IMAGE-LINK+2017-12-21-excel-setvalue1}! 

**Example 3:**

In case we want to check the condition and the content of a cell is other than 'TRUE' or 'FALSE'- in this case 'window', we need to use the following formula: `if ⊂♥res == "window"⊃`

```G1ANT
excel.open 
excel.setvalue value ‴window‴ row 1 colindex 1
excel.getvalue row 1 colindex 1 result ♥res
excel.setvalue value ‴something‴ row 1 colindex 2 if ⊂♥res == "window"⊃

```

**Example 4:**

```G1ANT
excel.open
excel.setvalue value ‴random input‴ row 1 colname ‴A‴

```