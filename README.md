# 4d-plugin-office-document-converter
Convert document to PDF with Microsoft Office application suite

### Platform

| carbon | cocoa | win32 | win64 |
|:------:|:-----:|:---------:|:---------:|
|<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|||

### Version

<img src="https://cloud.githubusercontent.com/assets/1725068/18940649/21945000-8645-11e6-86ed-4a0f800e5a73.png" width="32" height="32" /> <img src="https://cloud.githubusercontent.com/assets/1725068/18940648/2192ddba-8645-11e6-864d-6d5692d55717.png" width="32" height="32" />

### Remarks

Creates and removes temporary file inside ``application_darwin_temp_dir`` of ``Container.plist`` if ``MicrosoftBuildNumber`` is ``15`` or above.

Only tested with Microsoft Excel 2016.

---

## Syntax

```
OFFICE DOCUMENT TO PDF (src;dst)
```

Parameter|Type|Description
------------|------------|----
src|TEXT|path (HFS)
dst|TEXT|path (HFS)
