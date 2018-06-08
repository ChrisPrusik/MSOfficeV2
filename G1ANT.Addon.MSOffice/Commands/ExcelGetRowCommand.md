# excel.getrow

**Syntax:**

```G1ANT
excel.getrow  row ‴‴

```

**Description:**

Command `excel.getrow` gets all used cells of the specified row.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`row`| "integer":{TOPIC-LINK+integer}| yes |  | cell's row number |
|`result`| "variable":{TOPIC-LINK+variable}| no |  "♥result":{TOPIC-LINK+common-arguments} | name of variable where command's result will be stored |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no | | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example:**

```G1ANT
excel.open
window title ‴✱Excel‴
keyboard ‴THE wild bee reels from bough to bough ⋘DOWN⋙With his furry coat and his gauzy wing. ⋘DOWN⋙Now in a lily-cup, and now THE wild bee reels from bough to bough ⋘DOWN⋙Setting a jacinth bell a-swing,  ⋘DOWN⋙In his wandering; ⋘DOWN⋙Sit closer love: it was here I trow ⋘DOWN⋙I made that vow⋘DOWN⋙With his furry coat and his gauzy wing. ,‴ 
keyboard ⋘ENTER⋙
excel.getrow row ‴3‴ result ♥rowInput
dialog ♥rowInput

```