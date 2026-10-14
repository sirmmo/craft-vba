//! PowerPoint macros end to end against a presentation in memory that answers as DeckCraft does.

use serde_json::Value as Json;

use crate::fakedeck::Deck;
use crate::host::{Host, HostError};
use crate::model::App;
use crate::tests::big_stack;
use crate::vba::{Failure, Interp, Program};

struct FakeDeck(Deck);

impl Host for FakeDeck {
    fn call(&mut self, method: &str, params: Json) -> Result<Json, HostError> {
        self.0.handle(method, &params).map_err(|(code, message, data)| HostError { code, message, data })
    }
}

fn run_in(deck: Deck, src: &str) -> (Result<(), Failure>, Deck) {
    let src = src.to_string();
    big_stack(move || {
        let prog = Program::new(&[("Module1".to_string(), src)]);
        assert!(prog.broken.is_empty(), "{:?}", prog.broken);
        let mut host = FakeDeck(deck);
        let r = {
            let mut it = Interp::for_app(&prog, &mut host, App::PowerPoint);
            let (m, p) = it.find("Main").unwrap();
            it.run(m, p)
        };
        (r, host.0)
    })
}

fn run(src: &str) -> Deck {
    let (r, deck) = run_in(Deck::default(), src);
    if let Err(f) = r {
        panic!("{} (line {})", f.text(), f.line);
    }
    deck
}

fn fails(src: &str) -> Failure {
    run_in(Deck::default(), src).0.unwrap_err()
}

/// What the macro logged with Debug.Print.
fn logs(src: &str) -> Vec<String> {
    run(src).logs
}

#[test]
fn slides_are_counted_indexed_walked_added_and_deleted() {
    let d = run(r#"
Sub Main()
    Dim s As Slide, n As Long
    Debug.Print ActivePresentation.Slides.Count
    Debug.Print ActivePresentation.Slides(2).SlideIndex; ActivePresentation.Slides.Item(1).SlideID
    For Each s In ActivePresentation.Slides
        n = n + s.SlideIndex
    Next
    Debug.Print n
    Set s = ActivePresentation.Slides.Add(2, ppLayoutBlank)
    Debug.Print s.SlideIndex; ActivePresentation.Slides.Count; s.Layout = ppLayoutBlank; s.CustomLayout.Name
    Set s = Application.ActivePresentation.Slides.Add(4, ppLayoutTitle)
    Debug.Print s.Layout; ActivePresentation.Slides.FindBySlideID(257).SlideIndex
    ActivePresentation.Slides(1).Delete
    Debug.Print ActivePresentation.Slides.Count; ActivePresentation.Slides(1).SlideID = 300
End Sub"#);
    assert_eq!(d.logs, ["2", "2 256", "3", "2 3 True Blank", "1 3", "3 True"]);
    assert_eq!(d.slides.iter().map(|s| s.id).collect::<Vec<_>>(), [300, 257, 301]);
    assert_eq!(d.slides[0].layout, 2147483655, "ppLayoutBlank found the Blank layout");
    assert_eq!(d.slides[2].layout, 2147483649);
}

#[test]
fn slides_add_without_a_layout_list_uses_the_default_layout() {
    let deck = Deck { layouts_api: false, ..Deck::default() };
    let (r, d) =
        run_in(deck, "Sub Main()\r\n    ActivePresentation.Slides(1).Shapes(1).Name = \"T\"\r\n    ActivePresentation.Slides.Add 1, ppLayoutBlank\r\nEnd Sub");
    r.unwrap();
    assert_eq!(d.slides.len(), 3);
    assert_eq!(d.slides[0].layout, 2147483650);
    assert_eq!(d.slides[1].shapes[0].name, "T", "the write queued before the layout lookup wasn't lost");
}

#[test]
fn shapes_are_added_found_walked_and_deleted() {
    let d = run(r#"
Sub Main()
    Dim sh As Shape, names As String
    With ActivePresentation.Slides(2).Shapes
        .AddTextbox msoTextOrientationHorizontal, 10, 20, 300, 40
        .AddShape msoShapeRectangle, 0, 0, 50, 50
        .AddShape msoShapeRoundedRectangle, 0, 0, 50, 50
        .AddShape msoShapeOval, 0, 0, 50, 50
        .AddShape msoShapeIsoscelesTriangle, 0, 0, 50, 50
        .AddShape msoShapeRightArrow, 0, 0, 50, 50
        .AddShape(msoShape5pointStar, 1, 2, 3, 4).Name = "Star"
        .AddLine 100, 200, 40, 50
        Debug.Print .Count
        For Each sh In ActivePresentation.Slides(2).Shapes
            names = names & sh.Type & ","
        Next
        Debug.Print names
        Debug.Print .Item("star").Id; .Item(1).Left; .Item(1).Top; .Item(1).Width; .Item(1).Height
        .Item("Star").Delete
        .Item(1).Delete
        Debug.Print .Count; .Item(1).Name
    End With
End Sub"#);
    assert_eq!(d.logs[0], "8");
    assert_eq!(d.logs[1], "17,1,1,1,1,1,1,9,");
    assert_eq!(d.logs[2], "306 10 20 300 40");
    assert_eq!(d.logs[3], "6 Rectangle 2");
    let kinds: Vec<&str> = d.slides[1].shapes.iter().map(|s| s.geometry.as_str()).collect();
    assert_eq!(kinds, ["rectangle", "roundRectangle", "ellipse", "triangle", "rightArrow", "straight"]);
    let line = d.slides[1].shapes.last().unwrap();
    assert_eq!((line.left, line.top, line.width, line.height), (40.0, 50.0, 60.0, 150.0));
}

#[test]
fn shape_geometry_and_names_are_read_and_written() {
    let d = run(r#"
Sub Main()
    Dim sh As Shape
    Set sh = ActivePresentation.Slides(1).Shapes("Title 1")
    sh.Left = sh.Left + 15
    sh.Top = 30
    sh.Width = 400
    sh.Height = sh.Width / 4
    sh.Name = "Heading"
    Debug.Print sh.Name; sh.Left; sh.Height; sh.HasTextFrame = msoTrue; TypeName(sh)
    Debug.Print ActivePresentation.Slides(1).Shapes("Heading").Id
End Sub"#);
    let s = d.shape(0, "Heading");
    assert_eq!((s.left, s.top, s.width, s.height), (15.0, 30.0, 400.0, 100.0));
    assert_eq!(d.logs, ["Heading 15 100 True Shape", "260"]);
}

#[test]
fn text_and_fonts() {
    let d = run(r#"
Sub Main()
    Dim tr As TextRange
    With ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange
        Debug.Print Replace(.Text, vbCr, "|"); .Length
        .Text = "One" & vbCr & "Two"
        .Font.Bold = msoTrue
        .Font.Italic = True
        .Font.Size = 24
        .Font.Name = "Georgia"
        .Font.Color.RGB = RGB(255, 0, 0)
        Debug.Print .Font.Bold; .Font.Italic; .Font.Size; .Font.Name; .Font.Color.RGB; .Font.Underline
        .Font.Underline = msoTrue
        .Font.Bold = msoTriStateToggle
    End With
    Set tr = ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange
    tr = "Plain"
    tr.InsertAfter "!"
    Debug.Print tr; Len(tr); ActivePresentation.Slides(1).Shapes(1).TextFrame.HasText
    ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Font.Color = vbBlue
    ActivePresentation.Slides(1).Shapes(1).TextFrame.WordWrap = msoFalse
End Sub"#);
    assert_eq!(d.logs, ["First line|Second line 22", "-1 -1 24 Georgia 255 0", "Plain! 6 -1"]);
    let s = d.shape(0, "Subtitle 2");
    assert_eq!(s.text, "One\nTwo");
    assert!(!s.bold, "toggled back off");
    assert!(s.italic && s.underline);
    assert_eq!((s.size, s.font.as_str(), s.color.as_str()), (24.0, "Georgia", "#FF0000"));
    let t = d.shape(0, "Title 1");
    assert_eq!((t.text.as_str(), t.color.as_str(), t.word_wrap), ("Plain!", "#0000FF", false));
}

#[test]
fn fill_and_line() {
    let d = run(r#"
Sub Main()
    Dim sh As Shape
    Set sh = ActivePresentation.Slides(2).Shapes.AddShape(msoShapeOval, 0, 0, 10, 10)
    Debug.Print sh.Fill.Visible; sh.Fill.ForeColor.RGB; sh.Line.Visible; sh.Line.Weight
    sh.Fill.ForeColor.RGB = RGB(0, 128, 255)
    sh.Line.ForeColor.RGB = RGB(1, 2, 3)
    sh.Line.Weight = 2.5
    sh.Fill.Transparency = 0.25
    Debug.Print sh.Fill.ForeColor.RGB; sh.Line.ForeColor; sh.Fill.Transparency
    With ActivePresentation.Slides(1).Shapes(1)
        Debug.Print .Fill.Visible
        .Fill.Visible = msoTrue
        .Line.Visible = msoFalse
    End With
    ActivePresentation.Slides(2).Shapes(1).Line.Visible = msoFalse
    ActivePresentation.Slides(2).Shapes(1).Fill.Visible = msoFalse
End Sub"#);
    assert_eq!(d.logs, ["-1 12874308 -1 1", "16744448 197121 0.25", "0"]);
    let s = &d.slides[1].shapes[0];
    assert_eq!((s.fill.as_deref(), s.line_color.as_str(), s.line_weight, s.line_visible), (None, "#010203", 2.5, false));
    assert_eq!(s.transparency, 0.25);
    let t = d.shape(0, "Title 1");
    assert_eq!(t.fill.as_deref(), Some("#FFFFFF"), "a fill made visible is solid");
    assert!(!t.line_visible);
}

#[test]
fn colours_are_vbas_numbers() {
    let d = logs(
        r#"
Sub Main()
    With ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Font.Color
        .RGB = RGB(255, 0, 0)
        Debug.Print .RGB
        .RGB = &HFF0000
        Debug.Print .RGB; .Type
        .RGB = vbGreen
        Debug.Print .RGB
    End With
End Sub"#,
    );
    assert_eq!(d, ["255", "16711680 1", "65280"]);
    let e = fails("Sub Main()\r\n    ActivePresentation.Slides(1).Shapes(1).Fill.ForeColor.RGB = -5\r\nEnd Sub");
    assert_eq!(e.number, 5);
}

#[test]
fn the_selection() {
    let deck = Deck { selected_slides: vec![256], selected_shapes: vec![261], ..Deck::default() };
    let (r, d) = run_in(
        deck,
        r#"
Sub Main()
    Dim s As Shape
    With ActiveWindow.Selection
        Debug.Print .Type = ppSelectionShapes; .ShapeRange.Count; .ShapeRange(1).Name; .ShapeRange.Name; .SlideRange.SlideIndex
        .ShapeRange.Left = 77
        .ShapeRange(1).TextFrame.TextRange.Text = "picked"
    End With
    For Each s In ActiveWindow.Selection.ShapeRange
        s.Top = 5
    Next
    ActivePresentation.Slides(1).Shapes(1).Select
    ActivePresentation.Slides(1).Shapes(2).Select msoFalse
    Debug.Print ActiveWindow.Selection.ShapeRange.Count
    ActivePresentation.Slides(2).Select
    Debug.Print ActiveWindow.Selection.Type; ActiveWindow.Selection.SlideRange(1).SlideID; ActiveWindow.View.Slide.SlideIndex
    ActiveWindow.View.GotoSlide 1
    Debug.Print ActiveWindow.Selection.SlideRange.SlideIndex
End Sub"#,
    );
    r.unwrap();
    assert_eq!(d.logs, ["True 1 Subtitle 2 Subtitle 2 1", "2", "1 257 2", "1"]);
    let s = d.shape(0, "Subtitle 2");
    assert_eq!((s.left, s.top, s.text.as_str()), (77.0, 5.0, "picked"));
}

#[test]
fn shapes_reached_through_the_selection_without_their_slide() {
    let deck = Deck { selected_slides: vec![], selected_shapes: vec![260], ..Deck::default() };
    let (r, d) = run_in(deck, "Sub Main()\r\n    ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.Bold = True\r\nEnd Sub");
    r.unwrap();
    assert!(d.shape(0, "Title 1").bold);
    let empty = Deck { selected_shapes: vec![], ..Deck::default() };
    let (r, _) = run_in(empty, "Sub Main()\r\n    ActiveWindow.Selection.ShapeRange.Left = 1\r\nEnd Sub");
    assert_eq!(r.unwrap_err().number, 91);
}

#[test]
fn errors_are_vba_run_time_errors() {
    let e = fails("Sub Main()\r\n    Debug.Print ActivePresentation.Slides(99).SlideIndex\r\nEnd Sub");
    assert_eq!((e.number, e.text().as_str()), (9, "Run-time error '9': Subscript out of range"));
    assert_eq!(fails("Sub Main()\r\n    ActivePresentation.Slides(1).Shapes(\"Nope\").Left = 1\r\nEnd Sub").number, 9);
    assert_eq!(fails("Sub Main()\r\n    ActivePresentation.Slides.Add 9, ppLayoutText\r\nEnd Sub").number, 9);
    assert_eq!(fails("Sub Main()\r\n    ActivePresentation.Slides(1).Shapes.AddShape 999, 0, 0, 1, 1\r\nEnd Sub").number, 5);
    assert_eq!(fails("Sub Main()\r\n    ActivePresentation.Slides(1).Shapes.AddShape msoShapeOval\r\nEnd Sub").number, 449);
    assert_eq!(fails("Sub Main()\r\n    ActivePresentation.Slides(1).Frobnicate\r\nEnd Sub").number, 438);
    assert!(fails("Sub Main()\r\n    Range(\"A1\").Value = 1\r\nEnd Sub").compile, "no Excel in DeckCraft");
    // A shape deleted by the macro is gone: the app's ItemNotFound is error 9.
    let e = fails("Sub Main()\r\n    Dim s As Shape\r\n    Set s = ActivePresentation.Slides(1).Shapes(1)\r\n    s.Delete\r\n    s.Left = 4\r\nEnd Sub");
    assert_eq!(e.number, 9);
    // A line has no text frame.
    let e = fails("Sub Main()\r\n    ActivePresentation.Slides(2).Shapes.AddLine(0, 0, 1, 1).TextFrame.TextRange.Text = \"x\"\r\nEnd Sub");
    assert_eq!(e.number, 5);
    let (r, d) =
        run_in(Deck::default(), "Sub Main()\r\n    On Error Resume Next\r\n    ActivePresentation.Slides(42).Delete\r\n    Debug.Print Err.Number\r\nEnd Sub");
    r.unwrap();
    assert_eq!(d.logs, ["9"]);
}

#[test]
fn writes_are_batched_and_reads_cached() {
    let d = run(r#"
Sub Main()
    Dim i As Long, sh As Shape
    Set sh = ActivePresentation.Slides(1).Shapes(1)
    For i = 1 To 50
        sh.Left = i
    Next
    Debug.Print sh.Left; sh.Top; sh.Width; sh.Height; sh.Name
End Sub"#);
    assert_eq!(d.logs, ["50 0 100 50 Title 1"]);
    assert!(d.syncs <= 4, "{} syncs", d.syncs);
}

#[test]
fn messages_constants_and_application() {
    let d = run(r#"
Sub Main()
    MsgBox "Slides: " & ActivePresentation.Slides.Count, vbInformation, "Deck"
    Debug.Print Application.Name; ActivePresentation.Name; msoTrue; msoFalse; ppLayoutText; TypeName(ActivePresentation.Slides(1))
    Application.Run "Other", 3
End Sub
Sub Other(n)
    Debug.Print "other"; n
End Sub"#);
    assert_eq!(d.messages, vec![("Slides: 2".to_string(), Some("Deck".to_string()))]);
    assert_eq!(d.logs, ["DeckCraft Deck1 -1 0 2 Slide", "other 3"]);
}

#[test]
fn every_member_survives_odd_arguments() {
    // No member of the model may panic, whatever it's given.
    let members = [
        "Slides.Add",
        "Slides.Item",
        "Slides(0)",
        "Slides(-1)",
        "Slides(1).Shapes.AddTextbox",
        "Slides(1).Shapes.AddShape",
        "Slides(1).Shapes.AddLine",
        "Slides(1).Shapes.Item",
        "Slides(1).Shapes(1).Fill.ForeColor.RGB",
        "Slides(1).Shapes(1).Select",
        "Slides(1).Layout",
        "Slides(1).Shapes(1).TextFrame.AutoSize",
        "Slides(1).Shapes(1).TextFrame.TextRange.Start",
        "Slides(1).Shapes.FindById",
        "Slides(1).Shapes(1).IncrementLeft",
        "Slides(1).Shapes(1).IncrementTop",
        "Windows.Selection.TextRange.Text",
        "Windows.Selection.Unselect",
        "Windows.Selection.ShapeRange.IncrementLeft",
        "Windows.Selection.SlideRange.Delete",
    ];
    for m in members {
        for args in ["", " 1e300", " \"x\"", " -1, -1, -1, -1, -1", " 1, 1e308, 1, 1, 1"] {
            let src =
                format!("Sub Main()\r\n    On Error Resume Next\r\n    ActivePresentation.{m}{args}\r\n    Debug.Print ActivePresentation.{m}\r\nEnd Sub");
            let _ = run_in(Deck::default(), &src);
        }
    }
}

#[test]
fn recorded_deckcraft_steps_run_back() {
    use serde_json::json;
    let steps = [
        json!({"command": "slide.new", "params": {}}),
        json!({"command": "slide.new", "params": {"layout": "blank"}}),
        json!({"command": "shape.insert", "params": {"preset": "ellipse", "rect": [10, 20, 100, 50], "text": "Hi"}}),
        json!({"command": "format.bold", "params": {"on": true}}),
        json!({"command": "format.italic", "params": {}}),
        json!({"command": "format.underline", "params": {"on": true}}),
        json!({"command": "format.size", "params": {"size": 24}}),
        json!({"command": "format.font", "params": {"family": "Georgia"}}),
        json!({"command": "format.color", "params": {"color": "#FF0000"}}),
        json!({"command": "text.edit", "params": {}}),
        json!({"command": "text.insert", "params": {"text": "!"}}),
        json!({"command": "text.insert", "params": {"text": "\nnext"}}),
        json!({"command": "text.exit", "params": {}}),
        json!({"command": "shape.fill", "params": {"color": "#00FF00", "transparency": 0.5}}),
        json!({"command": "shape.line", "params": {"width": 3, "color": "#010203"}}),
        json!({"command": "shape.move", "params": {"dx": 5, "dy": -10}}),
        json!({"command": "arrange.nudge", "params": {"dx": 1, "dy": 0}}),
        json!({"command": "shape.resize", "params": {"w": 200, "lockAspect": true}}),
        json!({"command": "shape.rename", "params": {"name": "Sun"}}),
        json!({"command": "insert.textBox", "params": {"rect": [0, 0, 300, 40], "text": "Box"}}),
        json!({"command": "shape.setBounds", "params": {"x": 1, "y": 2, "w": 30, "h": 40}}),
        json!({"command": "edit.deselect", "params": {}}),
        json!({"command": "slide.go", "params": {"index": 0}}),
        json!({"command": "selection.select", "params": {"id": 261}}),
        json!({"command": "text.set", "params": {"text": "Replaced"}}),
        json!({"command": "shape.fill", "params": {"none": true, "ids": [260]}}),
        json!({"command": "edit.select", "params": {"ids": [260, 261]}}),
        json!({"command": "shape.move", "params": {"dx": 7}}),
        json!({"command": "edit.delete", "params": {"ids": [260]}}),
        json!({"command": "transition.apply", "params": {"kind": "fade"}}),
        json!({"command": "slide.next", "params": {}}),
        json!({"command": "slide.last", "params": {}}),
        json!({"command": "slide.next", "params": {}}),
        json!({"command": "slide.delete", "params": {}}),
        json!({"command": "slide.first", "params": {}}),
    ];
    let src = crate::record::record_powerpoint("Main", &steps);
    assert!(src.contains("    ' (step not recorded: transition.apply)\r\n"), "{src}");
    let (r, d) = run_in(Deck::default(), &src);
    if let Err(f) = r {
        panic!("{} (line {})\n{src}", f.text(), f.line);
    }
    // Slide 1, the new Title and Content slide (after a title slide), then the blank one; the
    // last slide (the old blank one) was deleted.
    assert_eq!(d.slides.iter().map(|s| (s.id, s.layout)).collect::<Vec<_>>(), [(256, 2147483649), (300, 2147483650), (301, 2147483655)]);
    let sun = d.shape(2, "Sun");
    assert_eq!(sun.geometry, "ellipse");
    assert_eq!((sun.left, sun.top, sun.width, sun.height), (16.0, 10.0, 200.0, 100.0));
    assert_eq!(sun.text, "Hi!\nnext");
    assert!(sun.bold && sun.italic && sun.underline);
    assert_eq!((sun.size, sun.font.as_str(), sun.color.as_str()), (24.0, "Georgia", "#FF0000"));
    assert_eq!((sun.fill.as_deref(), sun.transparency, sun.line_weight, sun.line_color.as_str()), (Some("#00FF00"), 0.5, 3.0, "#010203"));
    let boxed = &d.slides[2].shapes[1];
    assert_eq!((boxed.kind.as_str(), boxed.text.as_str()), ("textBox", "Box"));
    assert_eq!((boxed.left, boxed.top, boxed.width, boxed.height), (1.0, 2.0, 30.0, 40.0));
    let first = &d.slides[0].shapes;
    assert_eq!(first.len(), 1, "Title 1 was deleted");
    assert_eq!((first[0].name.as_str(), first[0].text.as_str(), first[0].left), ("Subtitle 2", "Replaced", 7.0));
    assert_eq!(d.selected_slides, [256], "slide.first went back");
}

#[test]
fn the_recorders_extras_in_the_model() {
    let deck = Deck { selected_slides: vec![256], selected_shapes: vec![260, 261], ..Deck::default() };
    let (r, d) = run_in(
        deck,
        r#"
Sub Main()
    With ActiveWindow.Selection.ShapeRange
        .IncrementLeft 10
        .IncrementTop -5
    End With
    Debug.Print ActivePresentation.Slides(1).Shapes.FindById(261).Name; ActivePresentation.Slides(1).Shapes.FindById("260").Left
    ActivePresentation.Slides(1).Shapes.FindById(261).Select
    ActiveWindow.Selection.TextRange.Font.Bold = msoTrue
    Debug.Print ActiveWindow.Selection.TextRange.Text = "First line" & vbCr & "Second line"
    ActiveWindow.Selection.Unselect
    Debug.Print ActiveWindow.Selection.Type
    ActiveWindow.Selection.Unselect
    ActivePresentation.Slides(1).Select
End Sub"#,
    );
    r.unwrap();
    assert_eq!(d.logs, ["Subtitle 2 10", "True", "1"]);
    let s = d.shape(0, "Subtitle 2");
    assert_eq!((s.left, s.top), (10.0, 115.0));
    assert!(s.bold);
    assert_eq!(fails("Sub Main()\r\n    ActivePresentation.Slides(1).Shapes.FindById(999).Select\r\nEnd Sub").number, 9);
    assert_eq!(fails("Sub Main()\r\n    ActiveWindow.Selection.TextRange.Text = \"x\"\r\nEnd Sub").number, 91);
    let two = Deck { selected_shapes: vec![260, 261], ..Deck::default() };
    assert_eq!(run_in(two, "Sub Main()\r\n    ActiveWindow.Selection.TextRange.Text = \"x\"\r\nEnd Sub").0.unwrap_err().number, 438);
    // Several slides selected are deleted together.
    let both = Deck { selected_slides: vec![256, 257], ..Deck::default() };
    let (r, d) = run_in(both, "Sub Main()\r\n    ActivePresentation.Slides.Add 3, ppLayoutBlank\r\n    ActiveWindow.Selection.SlideRange.Delete\r\nEnd Sub");
    r.unwrap();
    assert_eq!(d.slides.iter().map(|s| s.id).collect::<Vec<_>>(), [300]);
}

#[test]
fn shape_constants_use_the_vba_names_first() {
    let out = logs("Sub Main()\r\nDebug.Print msoShapePentagon\r\nDebug.Print msoShapeRegularPentagon\r\nDebug.Print msoShapeRightTriangle\r\nEnd Sub\r\n");
    assert_eq!(out, ["51", "12", "8"]);
}

#[test]
fn right_triangles_and_home_plates_are_recorded() {
    let steps = [
        serde_json::json!({"command": "shape.insert", "params": {"preset": "rtTriangle", "rect": [10, 20, 30, 40]}}),
        serde_json::json!({"command": "shape.insert", "params": {"preset": "homePlate", "rect": [10, 20, 30, 40]}}),
    ];
    let src = crate::record::record_powerpoint("Macro1", &steps);
    assert!(src.contains("AddShape(msoShapeRightTriangle, 10, 20, 30, 40)"), "{src}");
    assert!(src.contains("AddShape(msoShapePentagon, 10, 20, 30, 40)"), "{src}");
}
