# excel.addsheet

**Syntax:**

```G1ANT
excel.addsheet

```

**Description:**

Command `excel.addsheet` allows to add a new sheet to the currently active Excel instance.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`name`| "string":{TOPIC-LINK+string}| no |  | sheet name to be added |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeout":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no |  | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

In this example a new sheet in Excel is created and renamed to "new_books".

```G1ANT
excel.addsheet name ‴new_books‴

```

!{IMAGE-LINK+excel-newsheet}! 

**Example 2:**

The `name` argument for the `excel.addsheet` command can be assigned to a variable.

```G1ANT
♥name=newName
excel.open result ♥excelHandle
excel.addsheet name ♥name

```

**Example 3:**

But it can also be a string.

```G1ANT
excel.open result ♥excelHandle
excel.addsheet name ‴NewSheet‴

```