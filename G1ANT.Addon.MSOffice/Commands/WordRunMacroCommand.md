# word.runmacro

**Syntax:**

```G1ANT
word.runmacro  name ‴‴ 

```

**Description:**

Command `word.runmacro` allows to run macro in currently active Word instance. 

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`name`| "string":{TOPIC-LINK+string}| yes |  |name of the macro|
|`args`| "string":{TOPIC-LINK+string} | no |  | arguments for specified macro|
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no | | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1**:

In the following example a previously created macro is initialised. This macro requires no arguments, therefore the `args` are not specified.

```G1ANT
word.open
word.runmacro name ‴Align_Right‴

```

!{IMAGE-LINK+word-makro}! 
In this example a new Word document is opened, and a previously defined macro is initialised. This macro aligns the text (and the cursor) on the right side of the page.