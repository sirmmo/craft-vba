# craft-vba

VBA macros for the craft apps (GridCraft and DeckCraft), as an add-in that runs in its own
process. The app draws the Developer tab (Visual Basic, Macros, Record Macro); craft-vba does the
language:

- reads and writes the document's VBA project (`vbaProject.bin`, [MS-OVBA]);
- runs its macros with an interpreter for VBA ([MS-VBAL]);
- reaches the app's objects through its `host/sync`, the same object paths as the Office
  add-ins' bridge: Excel's (`Range`, `Cells`, `Worksheets`, …) in GridCraft, PowerPoint's
  (`ActivePresentation`, `Slides`, `Shapes`, …) in DeckCraft, chosen by `initialize`'s `host.id`
  (`deckcraft` gets PowerPoint's, every other app Excel's);
- records GridCraft's commands as a VBA Sub.

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
| `macros/run {"name": "Module1.Hello" \| "Hello", "control"?: {"id", "tag"}}` | `{}`, or an error `{"code": 1, "message": "Run-time error '13': Type mismatch", "data": {"module", "line", "number", "compile"}}` |
| `macros/modules {}` | `{"modules": [{"name", "kind": "standard"\|"class"\|"document", "source"}]}` |
| `macros/saveModule {"name", "source"}` | `{}` (or `{"compileError": {"line", "message"}}`: saved anyway); creates the project when there's none |
| `macros/deleteModule {"name"}` | `{}` |
| `macros/record {"steps": [{"command", "params"}], "name"?}` | `{"source": "Sub Macro1()…End Sub"}` (in DeckCraft an error: PowerPoint has no macro recorder) |

While the app waits for an answer, craft-vba asks it:

| Request | Answer |
|---|---|
| `host/sync {"actions": [...]}` | `{"results": [...]}`, or an error whose `data` is `{"code": "ItemNotFound", "action": i}` |
| `host/vbaProject/get {}` | `{"base64": "…" \| null}` |
| `host/vbaProject/set {"base64": "…"}` | `{}` |
| `host/message {"text", "title"?}` | `{"result": 1}` (MsgBox) |
| `host/log {"text"}` | `{}` (Debug.Print) |

Line numbers count the module's stored source, its `Attribute VB_Name` line included.

`control` is the ribbon button that ran the macro: a Sub with a parameter gets it as Excel passes
it (`Sub Go(control As IRibbonControl)`, with `control.Id` and `control.Tag`); a Sub without
parameters runs as it is. Such callbacks aren't in `macros/list`, as in Excel's Macros box.

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
  `ClearContents`, `ClearFormats`, `Clear`, `Select`, `Activate`), `IRibbonControl` (`Id`, `Tag`,
  `Context`). Colours are VBA's numbers
  (`RGB(255, 0, 0)` is red).

### PowerPoint (DeckCraft)

- `Application` (`ActivePresentation`, `ActiveWindow`, `Presentations`, `Name`, `Version`,
  `Run`), `Presentation` (`Slides`, `Name`).
- `Slides` (`Count`, `Item`/`(i)`, `Add(Index, ppLayout…)`, `FindBySlideID`, `For Each`),
  `Slide` (`SlideIndex`, `SlideNumber`, `SlideID`, `Shapes`, `Layout`, `CustomLayout.Name`,
  `Delete`, `Select`).
- `Shapes` (`Count`, `Item`/`(i)` by number or name, `AddTextbox`, `AddShape`, `AddLine`,
  `For Each`), `Shape` (`Name`, `Left`, `Top`, `Width`, `Height`, `Id`, `Type`, `HasTextFrame`,
  `TextFrame`, `Fill`, `Line`, `Select [Replace]`, `Delete`).
- `TextFrame` (`TextRange`, `HasText`, `WordWrap`, `AutoSize`), `TextRange` (`Text` (paragraphs
  end with `vbCr`, as in PowerPoint), `Length`, `Start`, `InsertAfter`, `InsertBefore`, `Font`),
  `Font` (`Bold`, `Italic`, `Underline`, `Size`, `Name`, `Color`), `ColorFormat` (`RGB`, `Type`),
  `FillFormat` (`ForeColor`, `Visible`, `Transparency`, `Solid`), `LineFormat` (`ForeColor`,
  `Weight`, `Visible`).
- `ActiveWindow.Selection` (`Type`, `ShapeRange`, `SlideRange`; a range of one shape or slide
  answers for it, `Left`/`Top`/`Width`/`Height` and `Delete` apply to all its shapes),
  `ActiveWindow.View` (`GotoSlide`, `Slide`), `IRibbonControl`.
- `AddShape` draws `msoShapeRectangle`, `msoShapeRoundedRectangle`, `msoShapeOval`,
  `msoShapeIsoscelesTriangle`, `msoShapeRightTriangle`, `msoShapeDiamond`,
  `msoShapeParallelogram`, `msoShapeTrapezoid`, `msoShapeOctagon`, `msoShapeHexagon`,
  `msoShapeRegularPentagon`, `msoShapeCross`, `msoShapeCan`, `msoShapeCube`, `msoShapeBevel`,
  `msoShapeFoldedCorner`, `msoShapeSmileyFace`, `msoShapeDonut`, `msoShapeNoSymbol`,
  `msoShapeBlockArc`, `msoShapeHeart`, `msoShapeLightningBolt`, `msoShapeSun`, `msoShapeMoon`,
  the arrows (`Right`, `Left`, `Up`, `Down`, `LeftRight`, `UpDown`), `msoShapePentagon`,
  `msoShapeChevron`, `msoShapeFlowchartProcess` and the 4, 5, 8, 16, 24 and 32-point stars.
- Constants: `msoTrue`/`msoFalse`/`msoTriStateToggle`, `msoTextOrientation…`, `msoShape…`,
  shape types (`msoAutoShape`, `msoTextBox`, …), `ppLayout…`, `ppSelection…`, `ppAutoSize…`.
- Tri-states read as `msoTrue` (-1), `msoFalse` (0) or `msoTriStateMixed` (-2); colours are
  VBA's numbers. `Slides(99)` is run-time error 9, as a missing slide or shape is.

A macro is one undo step in the app. Reads are cached (a sheet's used range is read once), writes
are sent in batches.

## Limits

- PowerPoint: not yet tables, pictures, groups, placeholders by type, `Characters`/`Paragraphs`/
  `Runs` (text formatting applies to a shape's whole text), `Rotation`, `ZOrder`, `Duplicate`,
  `Copy`/`Paste`, animations, slide shows, `Shapes.Range`, notes, `TextFrame2`, `DashStyle`;
  `AddTextbox` makes horizontal text only; `AddLine` keeps the line's box but not its direction
  (a line from bottom-left to top-right is drawn top-left to bottom-right); `Slides.Add` finds
  the layout by its standard name (Title Slide, Title and Content, Blank, …) in the first
  master's layouts and uses the app's default layout when it can't; several selected shapes
  answer only the members listed above.
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
