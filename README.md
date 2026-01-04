# grouping-macro
VBA Macro for creating groups based on multiple conditions
The macro is intentionally **generic**: you plug in your own classification rules for what constitutes a CAN row vs a LID row, based on your plant codes, descriptions, or any other columns.

## What it does

- Reads data from the first worksheet in the workbook
- Classifies rows as CAN and/or LID using your rules
- Groups by **Material + Plant**
- Creates CAN-centric groupings and attaches LIDs
- Writes results into a newly created sheet: **CAN LID MAP**
- De-duplicates groups via a signature (based on an identifier column)

## Files

- `src/CanLidGrouping.bas` - the VBA module (import this into your workbook)

## Assumptions (default)

The code assumes:
- **Material** is in column **B**
- **Plant** is in column **E**
- **Description** is in column **J**
- **Identifier** used for grouping/signatures is in column **H**

You can adjust these references in the code if your layout differs.

## Setup / Installation

1. Download `src/CanLidGrouping.bas`
2. Open Excel -> `Alt + F11` (VBA Editor)
3. `File -> Import File…` and select `CanLidGrouping.bas`
4. Implement your rules in:
   - `IsCanRow(...)`
   - `IsLidRow(...)`
   - (optional) `ShouldAttachLidToCan(...)`
5. Run `Create_CANLID_GROUPS`

## Customizing CAN/LID detection

Open the module and edit:

```vb
Private Function IsCanRow(...)
    ' Insert your rules here
End Function

Private Function IsLidRow(...)
    ' Insert your rules here
End Function
