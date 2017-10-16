# 4d-plugin-office-document-converter
Convert document to PDF with Microsoft Office application suite

### Platform

| carbon | cocoa | win32 | win64 |
|:------:|:-----:|:---------:|:---------:|
|<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|||

### Version

<img src="https://cloud.githubusercontent.com/assets/1725068/18940649/21945000-8645-11e6-86ed-4a0f800e5a73.png" width="32" height="32" /> <img src="https://cloud.githubusercontent.com/assets/1725068/18940648/2192ddba-8645-11e6-864d-6d5692d55717.png" width="32" height="32" />

### Remarks

Creates and removes temporary file inside ``application_darwin_temp_dir`` of ``Container.plist`` if the running Office app is version 2016 or above (sandboxed). Otherwise, uses the temporary folder allocated to 4D.

---

## Syntax

```
OFFICE DOCUMENT TO PDF (src;dst)
```

Parameter|Type|Description
------------|------------|----
src|TEXT|path (HFS)
dst|TEXT|path (HFS)

### Discussion

There are mainly 2 ways to call AppleScript from Objective C; ``NSAppleScript`` and ``SBApplication``.

To use ScriptingBridge, one must first create a header file from the scripting definition inside the scriptable target app with ``sdef`` and ``sdp``. Normally the calls can be piped, but the definition for Word and PowerPoint seems to contain errors, so one must first export the file to disk and manually edit in order to generate the full header file.

```
sdef Word.app > word.sdef
sdp --basename Word -fh Word.sdef Word.h
```

Excel seems to work with either method. The code expressed in AppleScript looks like this:

```applescript
on convert_to_pdf(src_path, dst_path)
set src_path to POSIX file src_path as string
set dst_path to POSIX file dst_path as string
tell application id "com.microsoft.Excel"
set w to open workbook workbook file name src_path ¬
with read only and ignore read only recommended without notify and add to mru
try
save workbook as w filename dst_path file format PDF file format ¬
conflict resolution other session changes with overwrite without read only recommended, create backup and add to most recently used list
end try
close w saving no
end tell
end convert_to_pdf
``` 

Notice the 2 lines outside of the call to Excel, used to convert POSIX path to HFS. The same sequence of calls work over the ScriptingBridge too, but Excel 2011 only accepts HFS paths, while Excel 2016 accepts either HFS or POSIX paths.

The ``open``, ``save as`` and ``close`` commands for Word and PowerPoint do not seem to work over the ScriptingBridge. ``NSAppleScript`` can be used as a workaround.

For Word, there is a specialised ``open`` command that returns a document.

```applescript
on convert_to_pdf(src_path, dst_path) 
set src_path to POSIX file src_path as string 
set dst_path to POSIX file dst_path as string 
tell application id "com.microsoft.Word" 
set d to open file name src_path ¬
with read only and Revert without add to recent files and confirm conversions 
save as d file name dst_path file format format PDF ¬
with embed truetype fonts, save native picture format and save forms data without add to recent files 
close d saving no saving in src_path 
end tell  
end convert_to_pd
```

For PowerPoint, there is no specialised ``open`` command that returns a presentation object, so we need to iterate the list of presentations to get what was just opened.

```applescript
on convert_to_pdf(src_path, dst_path) 
set src_path to POSIX file src_path as string 
set dst_path to POSIX file dst_path as string 
tell application id "com.microsoft.PowerPoint" 
open src_path 
set pp to presentations whose full name is src_path 
if (count pp) is 1 then 
set p to item 1 of pp 
save p in dst_path as save as PDF 
close p saving no saving in src_path 
end if 
end tell 
end convert_to_pdf
```
