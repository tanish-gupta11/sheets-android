# Sheets for Android 0.2.0

This release adds Excel 97–2003 `.xls` import and export to the existing `.xlsx` and CSV workflows. It also improves phone/tablet layout around Android system bars and updates cell references when rows or columns are inserted or deleted.

The APK was installed on an Android 16 device. Cell editing and formula calculation worked; XLS and XLSX files were imported and exported. Independent readers opened the files exported on the device. JVM tests cover workbook round trips and core formulas.

Known limits: this lightweight editor does not preserve charts, pivot tables, merged cells, macros, conditional formatting, complex number/date formats, or all Excel formulas. Export creates a new file, leaving the original workbook untouched. Use a copy for complex or valuable workbooks.

The release APK is about 2.9 MB, requires Android 8.0 or newer, needs no network permission, and is signed with a dedicated release key. The owner must retain the private signing key to publish updates under the same Android package ID.
