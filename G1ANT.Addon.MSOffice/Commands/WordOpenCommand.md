# word.open

**Syntax:**

```G1ANT
word.open  path ‴‴

```

**Description:**

Command `word.open` opens Word program.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`path`| "string":{TOPIC-LINK+string} | yes |   | path of file that has to be opened, if not specified, Word will be opened anyway |
|`result`| "variable":{TOPIC-LINK+variable} | no |  "♥result":{TOPIC-LINK+common-arguments} | name of variable where command's result will be stored |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no | | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

This example opens Word document from path ‴C:\programs\document.docx‴ using `word.open` command. We are storing id of currently opened Word process in a variable **♥wld1** for later use.

```G1ANT
word.open path ‴C:\programs\document.docx‴ result ♥wId1

```

**Example 2:**

This example opens Word document using `word.open` command. 

```G1ANT
word.open

```

**Example 3:**

```G1ANT
word.open result ♥wordHandle
word.open path ‴C:\tests\wordtest.docx‴ result ♥wordHandleTest

```