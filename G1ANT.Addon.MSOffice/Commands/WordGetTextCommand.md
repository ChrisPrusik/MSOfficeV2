# word.gettext

**Syntax:**

```G1ANT
word.gettext

```

**Description:**

Command `word.gettext` allows to copy text from a Word document.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`result`| "variable":{TOPIC-LINK+string}| no | "♥result":{TOPIC-LINK+common-arguments} | name of variable where command's result will be stored |
|`if`| "bool":{TOPIC-LINK+boolean}| no | true | runs the command only if condition is true |
|`timeout`| "variable":{TOPIC-LINK+variable}| no | "♥timeoutcommand":{TOPIC-LINK+special-variables} | specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
|`errorjump` | "label":{TOPIC-LINK+label}| no | | name of the label to jump to if given `timeout` expires |
|`errormessage`| "string":{TOPIC-LINK+string}| no |  | message that will be shown in case error occurs and no `errorjump` argument is specified |

For more information about `if`, `timeout`, `errorjump` and `errormessage` arguments, please visit "Common Arguments":{TOPIC-LINK+common-arguments} manual page.

This command is contained in **G1ANT.Addon.MSOffice.dll**.
See: "https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice":https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice

**Example 1:**

Opens a file from a path, copies all text from a file and stores it in `♥mytext` variable.

```G1ANT
word.open path ‴C:\Users\user1\Documents\test.docx‴
 word.gettext result ♥mytext
 dialog ♥mytext

```

**Example 2:**

In this example we are opening word using `word.open` command, then we are inserting text into the opened Word instance using `word.inserttext` command. After that we are saving the file at a chosen direction and closing Word.
When we open the file again using `word.open`  command, you can see when the `word.gettext` command becomes useful: G1ANT.Robot can read the thext from the file, save it in a variable- in our case its name is: '♥copiedText' and perform further actions on the text. We are only dialoging it in the dialog window in this example. 

```G1ANT
word.open
word.inserttext text ‴Pastry sweet roll cupcake jelly-o carrot cake cake jelly beans chocolate bar. Pudding biscuit marzipan powder. Biscuit candy canes cheesecake chupa chups. Marzipan chocolate bar liquorice pastry marzipan ice cream cheesecake. Sweet jelly-o powder donut. Marzipan jelly gingerbread candy canes tart candy canes. Jelly beans dragée apple pie.‴ replacealltext true
word.save path ‴C:\test\test.docx‴
word.close
word.open path ‴C:\test\test.docx‴ timeout 10000
word.gettext result ♥copiedText
dialog ♥copiedText

```

!{IMAGE-LINK+2018-01-16-word-gettext}!