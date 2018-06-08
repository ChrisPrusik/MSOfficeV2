# word.switch

**Syntax:**

```G1ANT
word.switch  id ‴‴ 

```

**Description:**

Command `word.switch` allows to switch between Word windows. 

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`id`| "variable":{TOPIC-LINK+variable} | yes | | ID of Word window that was specified while using `word.open` command |
|`result`| "variable":{TOPIC-LINK+variable} | no |  "♥result":{TOPIC-LINK+common-arguments} |name of variable where true/false from completion the switch will be stored|
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no | | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1**:

This example switches between opened Word windows with a `word.switch` command.

```G1ANT
word.open result ♥1
word.open result ♥2
word.switch id ♥1

```

!{IMAGE-LINK+word-switch}! 

In this example the Word window no 1 has been activated.

**Example 2:**

```G1ANT
word.open result wordHandle
word.open path ‴C:\tests\wordtest.docx‴ result ♥wordHandleTest
word.switch id ♥wordHandle

```