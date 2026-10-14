# craft-vba

VBA macros for the craft apps (GridCraft first), as an add-in that runs in its own process. The
app draws the Developer tab (Visual Basic, Macros, Record Macro); craft-vba does the language:

- reads and writes the workbook's VBA project (`vbaProject.bin`, [MS-OVBA]);
- runs its macros with an interpreter for VBA ([MS-VBAL]);
- reaches Excel's objects (`Range`, `Cells`, `Worksheets`, …) through the app's `host/sync`,
  the same object paths as the Office add-ins' bridge;
- records the app's commands as a VBA Sub.

It speaks the craft plugin protocol (JSON-RPC 2.0 over standard input and output, one message per
line), version 1.1. Everything comes from the published specifications and public documentation
(clean room).

## Installing

In GridCraft, open **File › Options › Add-ins**, type `sirmmo/craft-vba` in **Install from GitHub**
and confirm. GridCraft downloads the release zip for your computer (macOS Apple silicon or Intel,
Windows x64, Linux x86_64) and loads it; the **Developer** tab appears. Macros in a workbook
stay off until you click **Enable Content** (Developer › Macro Security changes that).

This needs a GridCraft with protocol 1.1 and the Developer tab (branch `feat/officejs` for now).

## Building from source

```sh
cargo build --release
mkdir -p ~/.config/gridcraft/plugins/vba
cp target/release/craft-vba ~/.config/gridcraft/plugins/vba/
target/release/craft-vba --manifest > ~/.config/gridcraft/plugins/vba/plugin.json
```

(`~/Library/Application Support/GridCraft/plugins/vba` on macOS.)

## Protocol 1.1

The app asks (besides `initialize`):

| Request | Answer |
|---|---|
| `macros/list {}` | `{"macros": [{"name": "Module1.Hello", "module", "procedure"}]}`: public Subs without required parameters, of standard and document modules |
| `macros/run {"name": "Module1.Hello" \| "Hello"}` | `{}`, or an error `{"code": 1, "message": "Run-time error '13': Type mismatch", "data": {"module", "line", "number", "compile"}}` |
| `macros/modules {}` | `{"modules": [{"name", "kind": "standard"\|"class"\|"document", "source"}]}` |
| `macros/saveModule {"name", "source"}` | `{}` (or `{"compileError": {"line", "message"}}`: saved anyway); creates the project when there's none |
| `macros/deleteModule {"name"}` | `{}` |
| `macros/record {"steps": [{"command", "params"}], "name"?}` | `{"source": "Sub Macro1()…End Sub"}` |

While the app waits for an answer, craft-vba asks it:

| Request | Answer |
|---|---|
| `host/sync {"actions": [...]}` | `{"results": [...]}`, or an error whose `data` is `{"code": "ItemNotFound", "action": i}` |
| `host/vbaProject/get {}` | `{"base64": "…" \| null}` |
| `host/vbaProject/set {"base64": "…"}` | `{}` |
| `host/message {"text", "title"?}` | `{"result": 1}` (MsgBox) |
| `host/log {"text"}` | `{}` (Debug.Print) |

Line numbers count the module's stored source, its `Attribute VB_Name` line included.

## What runs

- **Modules:** `Option Explicit`, `Option Base`, `Option Compare Text`; module variables and
  constants (`Public`/`Private`), `Enum`; `Sub`, `Function`, `Property Get` with `ByVal`/`ByRef`,
  `Optional` with defaults, `ParamArray`. Calls across modules (`Module1.Proc` too) and
  `Application.Run`.
- **Statements:** `Dim`/`Static`/`ReDim [Preserve]`/`Const`/`Erase`; assignment, `Set`, `Let`,
  `Call`; `If` (block and single-line), `Select Case` (`Is`, `To`, lists), `For … Step`,
  `For Each` (arrays, ranges, worksheets, collections), `Do While/Until … Loop While/Until`,
  `While … Wend`, `Exit`, `With`, labels and `GoTo`, `On Error Resume Next/GoTo label/GoTo 0`,
  `Resume Next`, `Resume label`, `End`, `Debug.Print`.
- **Values:** Variant, Integer (overflow is error 6), Long, Double, String, Boolean, Date, Object,
  arrays (multi-dimensional, bounds checked), `Collection`, `Err` (`Number`, `Description`,
  `Raise`, `Clear`).
- **Functions:** strings (`Len`, `Left`, `Right`, `Mid`, `InStr`, `InStrRev`, `Replace`, `Trim`,
  `UCase`, `LCase`, `Split`, `Join`, `Format`, `Space`, `String`, `StrReverse`, `StrComp`, `Chr`,
  `Asc`), conversions (`CStr`, `CInt`, `CLng`, `CDbl`, `CBool`, `CDate`, `Val`, `Str`, `Hex`,
  `Oct`), tests (`IsNumeric`, `IsEmpty`, `IsNull`, `IsArray`, `IsObject`, `IsMissing`, `IsDate`),
  math (`Abs`, `Int`, `Fix`, `Round` (banker's), `Sqr`, `Sgn`, `Exp`, `Log`, `Sin`, `Cos`, `Tan`,
  `Atn`, `Rnd`), dates (`Now`, `Date`, `Time`, `Timer`, `Year`, `Month`, `Day`, `Hour`, `Minute`,
  `Second`, `Weekday`, `DateSerial`, `TimeSerial`), `UBound`, `LBound`, `Array`, `IIf`, `Choose`,
  `TypeName`, `VarType`, `RGB`, `MsgBox`, and the `vb…`/`xl…` constants.
- **Excel:** `Application` (`ActiveSheet`, `ActiveCell`, `Selection`, `ActiveWorkbook`,
  `ThisWorkbook`, `Worksheets`, `Sheets`, `Range`, `Cells`, `Rows`, `Columns`,
  `WorksheetFunction.Sum/Max/Min/Average/Count/CountA`, `ScreenUpdating` and friends accepted),
  `Worksheets` (`Count`, `Item` by name or number, `Add`, `For Each`), `Worksheet` (`Name`,
  `Index`, `Range`, `Cells`, `Rows`, `Columns`, `UsedRange`, `Activate`, `Select`, `Delete`),
  `Range` (`Value`/`Value2` (one value, or a 1-based 2-D array), `Formula`, `NumberFormat`,
  `Text`, `Address`, `Row`, `Column`, `Count`, `Rows`, `Columns`, `Cells`, `Item`, `Offset`,
  `Resize`, `EntireRow`, `EntireColumn`, `End(xlUp/xlDown/xlToLeft/xlToRight)`, `Font` (`Bold`,
  `Italic`, `Color`, `Size`, `Name`, `Underline`, `Strikethrough`), `Interior.Color`,
  `ClearContents`, `ClearFormats`, `Clear`, `Select`, `Activate`). Colours are VBA's numbers
  (`RGB(255, 0, 0)` is red).

A macro is one undo step in the app. Reads are cached (a sheet's used range is read once), writes
are sent in batches.

## Limits

- Not yet: class modules (`New Class1`), `Type`, `Property Let/Set`, `GoSub`, `Resume` without a
  label, `InputBox`, UserForms, events (`Workbook_Open`, `Worksheet_Change`), `Declare` (Windows
  API calls), `CreateObject` (COM), `Shell`, file I/O, date literals (`#1/2/2026#`), `Like` is
  supported but `Option Compare Database` isn't.
- Module variables start fresh for every macro run.
- A macro stops after 20,000,000 statements and 200 nested calls; Debug.Print joins its items
  with spaces.
- New projects are written with no performance cache (`_VBA_PROJECT` version 0xFFFF, as
  [MS-OVBA] asks of writers), references to VBA, Excel and OLE Automation, code page 1252.
  Editing an existing project keeps its references, other storages and `PROJECT` lines.

## Licence

MIT OR Apache-2.0.
