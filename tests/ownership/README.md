# Ownership regressions

`engine_harness.mm` builds the real `4DPlugin.cpp`, the plugin classes and the unchanged bundled `4D Plugin API` with ASan and UBSan. It replaces the 4D engine (`Call4D`) with a synthetic one. That engine counts live engine strings, handles and reference-counted pictures, and aborts on anything it does not model.

`OFFICE DOCUMENT TO PDF` runs through `PluginMain` with a missing source and with a synthetic plain-text source. The UTI check rejects both, so no Office application is launched or scripted.

```sh
tests/ownership/run.sh . fixed                # every test must pass

B=$(mktemp -d); git archive <base-sha> | tar -x -C "$B"
tests/ownership/run.sh "$B" baseline          # named defects must fail, controls pass
```

| Test | Checks |
|---|---|
| `defect_picture_fromfile_ownership` | `C_PICTURE::fromFileAtPath` frees the path once and releases the loaded picture |
| `defect_picture_setbytes_leak` | `C_PICTURE::setBytes` releases the picture it creates |
| `defect_picture_createcopy_ownership` | `createCopy*` reads only the converted type and leaves no empty owners |
| `defect_picture_grayscale_owner_leak`, `defect_picture_getsize_owner_leak`, `defect_picture_thumbnail_owner_leak` | no leaked empty picture owners |
| `control_*` | `OFFICE DOCUMENT TO PDF` missing-source and unsupported-type paths, `ARRAY_TEXT` transfer, picture parameter/return |

`C_PICTURE` is compiled into the plugin, but `OFFICE DOCUMENT TO PDF` never calls it.

Not covered:
- Windows and a real 4D runtime.
- The Word, Excel and PowerPoint conversion paths, which script real Office applications.
- How the real engine treats the old value of a picture output parameter. The harness assumes the most forgiving model: the engine releases it.
