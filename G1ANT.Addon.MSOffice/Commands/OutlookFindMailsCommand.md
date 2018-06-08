# outlook.findmails

**Syntax:**

```G1ANT
outlook.findmails  search ‴‴

```

**Description:**

Command `outlook.findmails` allows to search mails in Inbox and returns all mails that contain required keyword in the email Subject.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`search`| "string":{TOPIC-LINK+string}| yes | | word to be searched in Subject|
|`showmail`| "bool":{TOPIC-LINK+boolean}| no | false | if set to true, G1ANT.Robot will show all e-mails meeting the criteria |
|`result`| "variable":{TOPIC-LINK+variable}| no | "♥result":{TOPIC-LINK+common-arguments} | name of variable where command's result will be stored |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no | | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

```G1ANT
outlook.findmails search ‴G1ANT‴

```

**Example 2:**

```G1ANT
outlook.open
outlook.findmails search ‴G1ANT‴ showmail false result ♥result
dialog ♥result 

```

G1ANT.Robot shows 'true' in result if finds an email containing certain word. it will show 'false' if it does not find any.

!{IMAGE-LINK+2017-11-29-outlook-open}!