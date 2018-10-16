# G1ANT.Addon.MSOffice
Added powerpoint support

You can test it by running this simple set of commands:

addon core version 4.61.18249.1835
addon language version 4.60.18249.1521
addon msofficev2 version 2.205.17304.1001

powerpoint.open 
powerpoint.insertslide titletext ‴I ❤ ROBOT‴ slideposition ‴1‴ slidetext ‴Robot is awesome‴ slidelayout ‴Title‴

for counter ♥idx from 2 to 3 step 1
powerpoint.insertslide titletext ‴I am the ♥idx slide!‴ slideposition ♥idx slidelayout ‴ObjectAndText‴ slidetext ‴I am a text‴
powerpoint.inserttext texttobeinserted ‴I can insert text into any textbox!‴ textposition ‴3‴
powerpoint.insertphoto pathtophoto ‴C:\Users\PathToPhoto‴  left ‴200‴ top ‴200‴ height ‴200‴  width ‴200‴
end for
powerpoint.save presname ‴test2‴  path ‴C:\Users\PathToSave‴
powerpoint.close
