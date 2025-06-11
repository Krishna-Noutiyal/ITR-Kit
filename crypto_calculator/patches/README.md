# Openpyxl patch

Their is a known bug in openpyxl that doesn't preserver the External Links of the excel workbook resulting in a referencing error. As of developing this software `openpyxl` version `3.1.5` is being used and their is not official fix for this in the main release. Theirfore, you will have to apply the patch manually when building this application.

[!NOTE]
You don't have to run the script manually if you are building from `build.ps1` script.

## Fixing the patch

The fix is given in the [this](https://foss.heptapod.net/openpyxl/openpyxl/-/commit/5be1cea4e5488527b6624b7f34fe44731f0d5f50) commit. The apply_patch.ps1 script applies patch to the `openpyxl`.

## Running the patch

To run the patch you should be in the `.venv\Lib\site-packages\openpyxl` directory of the virtual environment. Run the following command in powershell:

[!NOTE]
Build the app using the `build.ps1` script first then runt he command below

```powershell
    ./patches/apply_patch.ps1
```
