# word.close

**Syntax:**

```G1ANT
word.close

```

**Description:**

Command `word.close` closes Word program that is currently active.
| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no | | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

This example closes the Word program that is currently active using `word.close` command.

```G1ANT
word.close

```

**Example 2:**

```G1ANT
word.open path ‴C:\tests\wordtest.docx‴ result ♥wordHandleTest
word.close

```