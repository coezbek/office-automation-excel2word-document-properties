# Excel2Word Poor Man's Office Automation using Document Properties

Generic way to put excel content into word documents as document properties (VBA)

## Objective

The goal of this template is to provide a generic method by which you can populate a Microsoft Word document with data from Excel worksheet cells.

Features:
- Only needs tinkering with VBA once to set up.
- Afterwards adding additional cells to be transferred to Word can be done without entering the VBA macro editor.
- The word template is embedded into the xlsx file, so need to keep the Word template somewhere.

## Caution

Running Office macros from untrusted source is dangerous. So make sure you know what you are doing. Please also see the [License](https://github.com/coezbek/office-automation-excel2word-document-properties/blob/main/LICENSE) for a disclaimer.

## How to set it up

1.) Open your excel file from which you want to export the information and save it as an `xlsm` (macro enabled xlsx file).

2.) Add the Word file to the Excel sheet that will be the template:

- `Insert` -> `Object` (under Text) -> `Create From File` -> Select the Word document file and check `Display as Icon`

3.) Rename the Object to "TemplateShape" so that we can reference it from VBA via the [Name Box](https://exceljet.net/glossary/name-box).

4.) Add a button on your excel sheet:

- `Developer` page -> `Insert` -> choose `Button` under `Form Controls` -> Draw the button somewhere -> Click `New`

This will open the VBA editor.

5.) Paste the entire code from [Excel2Word.bas](https://github.com/coezbek/office-automation-excel2word-document-properties/blob/main/Excel2Word.bas) (you can replace the `Sub` that was created when you created the button).

6.) [Assign names using the Name Box](https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64) to all cells which you want to export to Word.

Hint: Only individual cells have been tested.

7.) Run the `SetupWordTemplate`

`Developer` -> `Macros` -> Select `...SetupWordTemplate` -> `Run`

This will open the word file you added in step 2 and will put all named fields at the bottom of the document using [DocProperty fields]().

Rearrange these fields in your template to suit your needs, then close/save the template file (it automatically saves back into the Excel file).

8.) Assign the `ExportExcel2Word` macro to the button you created under step 4.

Right click on Button -> `Assign Macro...` -> Select `...ExportExcel2Word` -> `OK`

You might also want to rename the Button to something more meaningful such as `Export to Word` (right click on button and `Edit Text`).

9.) Press the button you created under step 4 and observe:

- Excel will shortly open the template word document and make a copy of it.
- Excel will populate all fields in the new copy and leave it open for you to edit and save.
- A suggested file name should be populated for you.

10.) Should you want to update the word template in the future, you can either run the `SetupWordTemplate` again (see step 7) or double click (i.e. edit) the word object that you added in step 2.

## Open Todos

- Add a way to automatically save/save as pdf the created report.
- Add a way to use a word template from disk rather than embedded into the report.
- Add handling of cell ranges to be exported nicely.
