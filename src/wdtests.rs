//! Word macros end to end against a document in memory that answers as WordCraft does.

use serde_json::Value as Json;

use crate::fakeword::Doc;
use crate::host::{Host, HostError};
use crate::model::App;
use crate::tests::big_stack;
use crate::vba::{Failure, Interp, Program};

struct FakeWord(Doc);

impl Host for FakeWord {
    fn call(&mut self, method: &str, params: Json) -> Result<Json, HostError> {
        self.0.handle(method, &params).map_err(|(code, message, data)| HostError { code, message, data })
    }
}

fn run_modules(doc: Doc, modules: Vec<(String, String)>) -> (Result<(), Failure>, Doc) {
    big_stack(move || {
        let prog = Program::new(&modules);
        assert!(prog.broken.is_empty(), "{:?}", prog.broken);
        let mut host = FakeWord(doc);
        let r = {
            let mut it = Interp::for_app(&prog, &mut host, App::Word);
            let (m, p) = it.find("Main").unwrap();
            it.run(m, p)
        };
        (r, host.0)
    })
}

fn run_in(doc: Doc, src: &str) -> (Result<(), Failure>, Doc) {
    run_modules(doc, vec![("Module1".to_string(), src.to_string())])
}

fn run(src: &str) -> Doc {
    let (r, doc) = run_in(Doc::default(), src);
    if let Err(f) = r {
        panic!("{} (line {})", f.text(), f.line);
    }
    doc
}

fn fails(src: &str) -> Failure {
    run_in(Doc::default(), src).0.unwrap_err()
}

#[test]
fn the_application_and_the_document() {
    let d = run(r#"
Sub Main()
    Debug.Print Application.Name; ActiveDocument.Name; ThisDocument.Name; Documents.Count; Documents(1).Name; TypeName(ActiveDocument)
    Debug.Print Documents("report.docx").Paragraphs.Count; Application.ActiveDocument.Content.Paragraphs.Count; TypeName(Selection.Font)
    Debug.Print Replace(ActiveDocument.Content.Text, vbCr, "|"); "/"; Replace(ActiveDocument.Range.Text, vbCr, "|")
    Debug.Print ActiveDocument.Paragraphs(1).Range; "/"; Len(ActiveDocument.Paragraphs(2).Range)
    ActiveDocument.Save
    MsgBox "Paragraphs: " & ActiveDocument.Paragraphs.Count, vbInformation, "Doc"
    Application.Run "Other", 3
End Sub
Sub Other(n)
    Debug.Print "other"; n
End Sub"#);
    assert_eq!(
        d.logs,
        [
            "WordCraft Report.docx Report.docx 1 Report.docx Document",
            "3 3 Font",
            "Hello world|Second paragraph.|Third.| / Hello world|Second paragraph.|Third.|",
            "Hello world\r / 18",
            "other 3"
        ]
    );
    assert_eq!(d.saves, 1);
    assert_eq!(d.messages, vec![("Paragraphs: 3".to_string(), Some("Doc".to_string()))]);
    // An app that doesn't answer the document's name: Word's name for a new document.
    let (r, d) = run_in(
        Doc { name: None, ..Doc::default() },
        "Sub Main()\r\n    ActiveDocument.Paragraphs(1).Alignment = 2\r\n    Debug.Print ActiveDocument.Name\r\nEnd Sub",
    );
    r.unwrap();
    assert_eq!(d.logs, ["Document1"]);
    assert_eq!(d.paras[0].alignment, "Right", "the write queued before the name lookup wasn't lost");
}

#[test]
fn paragraphs_are_counted_indexed_walked_and_added() {
    let d = run(r#"
Sub Main()
    Dim p As Paragraph, s As String
    With ActiveDocument.Paragraphs
        Debug.Print .Count; TypeName(.Item(1)); .Item(1).Style; ActiveDocument.Paragraphs(3).Style; .First.Style; .Last.Style
    End With
    For Each p In ActiveDocument.Paragraphs
        s = s & Len(p.Range.Text) & ","
    Next
    Debug.Print s
    Debug.Print Replace(ActiveDocument.Paragraphs(1).Next.Range.Text, vbCr, ""); ActiveDocument.Paragraphs(1).Previous Is Nothing; ActiveDocument.Paragraphs(3).Next Is Nothing
    Debug.Print Replace(ActiveDocument.Paragraphs(3).Previous.Previous.Range.Text, vbCr, "")
    Set p = ActiveDocument.Paragraphs.Add
    p.Range.Text = "Fourth"
    Debug.Print ActiveDocument.Paragraphs.Count; Replace(ActiveDocument.Paragraphs.Last.Range.Text, vbCr, "")
    Set p = ActiveDocument.Paragraphs(1)
    Do Until p Is Nothing
        p.Range.Font.Italic = True
        Set p = p.Next
    Loop
End Sub"#);
    assert_eq!(d.logs, ["3 Paragraph Title Normal Title Normal", "12,18,7,", "Second paragraph. True True", "Hello world", "4 Fourth"]);
    assert_eq!(d.texts(), ["Hello world", "Second paragraph.", "Third.", "Fourth"]);
    assert!(d.paras.iter().all(|p| p.chars.iter().all(|(_, f)| f.italic)));
}

#[test]
fn paragraph_formats() {
    let d = run(r#"
Sub Main()
    With ActiveDocument.Paragraphs(2)
        .Style = wdStyleHeading1
        .Alignment = wdAlignParagraphRight
        .SpaceAfter = 12
        .SpaceBefore = 6
        .LeftIndent = 36
        .RightIndent = 18
        .FirstLineIndent = -18
        Debug.Print .Style; .Alignment; .SpaceAfter; .SpaceBefore; .LeftIndent; .RightIndent; .FirstLineIndent; .Format.Alignment
    End With
    ActiveDocument.Paragraphs(3).Style = "Quote"
    ActiveDocument.Paragraphs(3).Style = ActiveDocument.Styles("Intense Quote")
    ActiveDocument.Paragraphs(1).Format.Alignment = wdAlignParagraphJustify
    ActiveDocument.Paragraphs(1).Range.ParagraphFormat.SpaceAfter = 3
    Debug.Print ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphJustify; ActiveDocument.Paragraphs(3).Style; ActiveDocument.Styles(wdStyleTitle).NameLocal
End Sub"#);
    assert_eq!(d.logs, ["Heading 1 2 12 6 36 18 -18 2", "True Intense Quote Title"]);
    let p = &d.paras[1];
    assert_eq!((p.style.as_str(), p.alignment.as_str()), ("Heading 1", "Right"));
    assert_eq!((p.space_after, p.space_before, p.left_indent, p.right_indent, p.first_line_indent), (12.0, 6.0, 36.0, 18.0, -18.0));
    assert_eq!((d.paras[0].alignment.as_str(), d.paras[0].space_after), ("Justified", 3.0));
    assert_eq!(d.paras[2].style, "Intense Quote");
}

#[test]
fn fonts_colours_and_mixed_reads() {
    let d = run(r#"
Sub Main()
    With ActiveDocument.Paragraphs(1).Range.Font
        Debug.Print .Bold = wdUndefined; .Italic; .Size; .Name; .Color; .Underline
        .Bold = True
        .Italic = wdToggle
        .Size = 14
        .Name = "Georgia"
        .Color = wdColorRed
        .Underline = wdUnderlineDouble
        Debug.Print .Bold; .Italic; .Size; .Name; .Color; .Underline
    End With
    ActiveDocument.Paragraphs(2).Range.Font.Color = RGB(0, 128, 255)
    Debug.Print ActiveDocument.Paragraphs(2).Range.Font.Color; ActiveDocument.Paragraphs(3).Range.Bold
    With ActiveDocument.Content
        Debug.Print .Font.Size = wdUndefined; "["; .Font.Name; "]"; .Font.Color = wdUndefined; .Font.Bold = wdUndefined; .Font.Underline = wdUndefined
        Debug.Print .ParagraphFormat.Alignment = wdUndefined; IsNull(.Style); ActiveDocument.Paragraphs(2).Range.Font.Underline
    End With
    ActiveDocument.Paragraphs(3).Range.Bold = True
    ActiveDocument.Paragraphs(3).Range.Font.Underline = True
End Sub"#);
    assert_eq!(d.logs, ["True 0 11 Calibri 0 0", "-1 -1 14 Georgia 255 3", "16744448 0", "True [  ] True True True", "True True 0"]);
    for ci in [0, 10] {
        let f = d.fmt(0, ci);
        assert!(f.bold && f.italic);
        assert_eq!((f.size, f.name.as_str(), f.color.as_str(), f.underline.as_str()), (14.0, "Georgia", "#FF0000", "Double"));
    }
    assert_eq!(d.fmt(1, 0).color, "#0080FF");
    assert!(!d.fmt(1, 0).bold, "a paragraph's font stops at its mark");
    assert!(d.fmt(2, 0).bold);
    assert_eq!(d.fmt(2, 0).underline, "Single");
}

#[test]
fn colour_conversion() {
    let d = run(r#"
Sub Main()
    With ActiveDocument.Paragraphs(2).Range.Font
        .Color = RGB(255, 0, 0)
        Debug.Print .Color
        .Color = &HFF0000
        Debug.Print .Color; .Color = wdColorBlue
        .Color = wdColorBrightGreen
        Debug.Print .Color
    End With
End Sub"#);
    assert_eq!(d.logs, ["255", "16711680 True", "65280"]);
    assert_eq!(d.fmt(1, 0).color, "#00FF00");
    assert_eq!(fails("Sub Main()\r\n    ActiveDocument.Content.Font.Color = -5\r\nEnd Sub").number, 5);
    assert_eq!(fails("Sub Main()\r\n    ActiveDocument.Content.Font.Color = wdColorAutomatic\r\nEnd Sub").number, 5);
}

#[test]
fn ranges_insert_delete_and_style() {
    let d = run(r#"
Sub Main()
    ActiveDocument.Content.InsertAfter "End."
    ActiveDocument.Content.InsertBefore "Start "
    ActiveDocument.Paragraphs(2).Range.InsertAfter "!"
    ActiveDocument.Paragraphs(2).Range.InsertBefore "> "
    ActiveDocument.Paragraphs(2).Range.InsertParagraphAfter
    ActiveDocument.Paragraphs(1).Range.InsertParagraphBefore
    Debug.Print ActiveDocument.Paragraphs.Count
    ActiveDocument.Paragraphs(4).Range.Text = "New" & vbCr & "Lines" & vbCr
    ActiveDocument.Paragraphs(1).Range.Delete
    ActiveDocument.Range(0, 0).InsertBefore "<"
    ActiveDocument.Paragraphs(1).Range.Style = "Heading 2"
    Debug.Print ActiveDocument.Paragraphs(1).Range.Style; ActiveDocument.Content.Paragraphs.Count; ActiveDocument.Paragraphs(2).Range.Paragraphs.Count
    ActiveDocument.Paragraphs(2).Range.HighlightColorIndex = wdYellow
    Debug.Print ActiveDocument.Paragraphs(2).Range.HighlightColorIndex; ActiveDocument.Paragraphs(3).Range.HighlightColorIndex
    ActiveDocument.Paragraphs(3).Range.HighlightColorIndex = wdBrightGreen
    ActiveDocument.Paragraphs(3).Range.HighlightColorIndex = wdNoHighlight
End Sub"#);
    assert_eq!(d.logs, ["5", "Heading 2 5 1", "7 0"]);
    assert_eq!(d.texts(), ["<Start Hello world", "> Second paragraph.!", "New", "Lines", "Third.End."]);
    assert_eq!(d.paras[0].style, "Heading 2");
    assert_eq!(d.fmt(1, 0).highlight.as_deref(), Some("Yellow"));
    assert_eq!(d.fmt(2, 0).highlight, None);
    // Content.Text = … replaces the whole document.
    let d = run("Sub Main()\r\n    ActiveDocument.Content.Text = \"One\" & vbCr & \"Two\"\r\n    Debug.Print ActiveDocument.Paragraphs.Count\r\nEnd Sub");
    assert_eq!(d.texts(), ["One", "Two"]);
    assert_eq!(d.logs, ["2"]);
}

#[test]
fn the_selection() {
    let d = run(r#"
Sub Main()
    Selection.TypeText "!"
    Selection.TypeParagraph
    Selection.TypeText Text:="Typed " & "more"
    Debug.Print Replace(Selection.Paragraphs(1).Range.Text, vbCr, "|"); Selection.Range.Text = ""
    Selection.HomeKey Unit:=wdStory
    Selection.TypeText "A: "
    Selection.EndKey wdStory
    Selection.TypeText "x" & vbCr & "y"
    ActiveDocument.Paragraphs(2).Range.Select
    Debug.Print Replace(Selection.Text, vbCr, "|"); Replace(Selection, vbCr, "|")
    Selection.Font.Bold = True
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Style = ActiveDocument.Styles("Quote")
    Debug.Print Selection.Font.Bold; Selection.ParagraphFormat.Alignment; Selection.Style; Selection.Range.Paragraphs.Count
    Selection.WholeStory
    Debug.Print Selection.Paragraphs.Count; Selection.ParagraphFormat.Alignment = wdUndefined
    ActiveDocument.Paragraphs(3).Range.Select
    Selection.Text = "Replaced"
    Selection.Font.Italic = True
    Debug.Print Selection.Text
    Selection.Range.InsertAfter "+"
End Sub"#);
    assert_eq!(d.logs, ["Typed more| True", "Second paragraph.| Second paragraph.|", "-1 1 Quote 1", "5 True", "Replaced"]);
    assert_eq!(d.texts(), ["A: Hello world", "Second paragraph.", "Replaced+", "Typed morex", "y"]);
    assert!(d.fmt(1, 0).bold && d.fmt(1, 16).bold && !d.fmt(0, 0).italic);
    assert_eq!((d.paras[1].alignment.as_str(), d.paras[1].style.as_str()), ("Centered", "Quote"));
    assert!(d.fmt(2, 0).italic, "the selection is the new text");
}

#[test]
fn ranges_by_id_follow_their_text() {
    let doc = Doc { sel: ((1, 0), (1, 6)), ..Doc::default() };
    let (r, d) = run_in(
        doc,
        r#"
Sub Main()
    Dim r As Range, r2 As Range
    Set r = Selection.Range
    Set r2 = r
    Debug.Print r.Text; TypeName(r); "/"; r.Font.Bold
    r.Text = "First"
    r.Font.Italic = True
    Debug.Print r2.Text
    r.InsertAfter "!"
    r.InsertBefore "<"
    On Error Resume Next
    Debug.Print r.Text
    Debug.Print Err.Number
End Sub"#,
    );
    r.unwrap();
    // Under On Error Resume Next the failing Debug.Print is skipped.
    assert_eq!(d.logs, ["Second String / 0", "First", "5"]);
    assert_eq!(d.texts()[1], "<First! paragraph.");
    assert!(d.fmt(1, 1).italic && d.fmt(1, 5).italic && !d.fmt(1, 8).italic);
}

#[test]
fn document_properties() {
    let d = run(r#"
Sub Main()
    Debug.Print ActiveDocument.BuiltInDocumentProperties("Title"); ActiveDocument.BuiltInDocumentProperties(wdPropertyAuthor).Value; ActiveDocument.BuiltInDocumentProperties("title").Name
    ActiveDocument.BuiltInDocumentProperties("Title") = "New title"
    ActiveDocument.BuiltInDocumentProperties("Subject").Value = "Sub"
    ActiveDocument.BuiltInDocumentProperties("Keywords") = "a, b"
    ActiveDocument.BuiltInDocumentProperties("Comments") = "c"
    ActiveDocument.BuiltInDocumentProperties("Author").Value = "Grace"
    Debug.Print ActiveDocument.BuiltInDocumentProperties("Title").Value
End Sub"#);
    assert_eq!(d.logs, ["Report Ada Title", "New title"]);
    assert_eq!((d.title.as_str(), d.subject.as_str(), d.keywords.as_str(), d.comments.as_str(), d.author.as_str()), ("New title", "Sub", "a, b", "c", "Grace"));
    assert_eq!(fails("Sub Main()\r\n    Debug.Print ActiveDocument.BuiltInDocumentProperties(\"Company\")\r\nEnd Sub").number, 5);
}

#[test]
fn errors_are_vba_run_time_errors() {
    let e = fails("Sub Main()\r\n    Debug.Print ActiveDocument.Paragraphs(99).Range.Text\r\nEnd Sub");
    assert_eq!((e.number, e.text().as_str()), (9, "Run-time error '9': Subscript out of range"));
    assert_eq!(fails("Sub Main()\r\n    ActiveDocument.Paragraphs(0).Alignment = 1\r\nEnd Sub").number, 9);
    assert_eq!(fails("Sub Main()\r\n    Documents(2).Save\r\nEnd Sub").number, 9);
    assert_eq!(fails("Sub Main()\r\n    ActiveDocument.Frobnicate\r\nEnd Sub").number, 438);
    assert_eq!(fails("Sub Main()\r\n    Debug.Print ActiveDocument.Content.Start\r\nEnd Sub").number, 438);
    assert_eq!(fails("Sub Main()\r\n    ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphDistribute\r\nEnd Sub").number, 5);
    assert_eq!(fails("Sub Main()\r\n    ActiveDocument.Paragraphs(1).Range.Font.Size = 0\r\nEnd Sub").number, 5);
    assert_eq!(fails("Sub Main()\r\n    Selection.HomeKey Unit:=wdLine\r\nEnd Sub").number, 5);
    assert_eq!(fails("Sub Main()\r\n    Selection.TypeText\r\nEnd Sub").number, 449);
    let e = fails("Sub Main()\r\n    Debug.Print ActiveDocument.Range(1, 5).Text\r\nEnd Sub");
    assert_eq!(e.number, 5);
    assert!(e.message.contains("character positions"), "{}", e.message);
    assert!(fails("Sub Main()\r\n    Range(\"A1\").Value = 1\r\nEnd Sub").compile, "no Excel in WordCraft");
    let (r, d) = run_in(
        Doc::default(),
        "Sub Main()\r\n    On Error Resume Next\r\n    ActiveDocument.Paragraphs(42).Range.Delete\r\n    Debug.Print Err.Number\r\nEnd Sub",
    );
    r.unwrap();
    assert_eq!(d.logs, ["9"]);
    // The app refuses a write: the run fails, and nothing of that batch was applied.
    let (r, d) = run_in(
        Doc::default(),
        "Sub Main()\r\n    ActiveDocument.Paragraphs(2).Alignment = wdAlignParagraphRight\r\n    ActiveDocument.Paragraphs(1).Style = \"\"\r\nEnd Sub",
    );
    assert_eq!(r.unwrap_err().number, 1004);
    assert_eq!(d.paras[1].alignment, "Left");
}

#[test]
fn writes_are_batched_and_reads_cached() {
    let d = run(r#"
Sub Main()
    Dim i As Long, p As Paragraph
    Set p = ActiveDocument.Paragraphs(1)
    For i = 1 To 50
        p.SpaceAfter = i
        p.Range.Font.Size = 10 + i
    Next
    Debug.Print p.SpaceAfter; p.SpaceBefore; p.Alignment; p.LeftIndent; p.Range.Font.Size; p.Range.Font.Name; p.Range.Font.Italic
End Sub"#);
    assert_eq!(d.logs, ["50 0 1 0 60 Calibri 0"]);
    assert!(d.syncs <= 4, "{} syncs", d.syncs);
}

#[test]
fn this_document_is_also_a_module() {
    let modules = vec![
        ("ThisDocument".to_string(), "Public Greeting As String\r\nSub Hello()\r\n    Greeting = \"hi\"\r\nEnd Sub\r\n".to_string()),
        (
            "Module1".to_string(),
            "Sub Main()\r\n    ThisDocument.Hello\r\n    Debug.Print ThisDocument.Greeting; ThisDocument.Paragraphs.Count; ThisDocument.Paragraphs(1).Style\r\nEnd Sub\r\n"
                .to_string(),
        ),
    ];
    let (r, d) = run_modules(Doc::default(), modules);
    r.unwrap();
    assert_eq!(d.logs, ["hi 3 Title"]);
}

#[test]
fn every_member_survives_odd_arguments() {
    // No member of the model may panic, whatever it's given.
    let members = [
        "Paragraphs.Item",
        "Paragraphs(0)",
        "Paragraphs(-1)",
        "Paragraphs.Add",
        "Paragraphs(1).Range.InsertAfter",
        "Paragraphs(1).Range.Text",
        "Paragraphs(1).Range.Font.Size",
        "Paragraphs(1).Range.Font.Color",
        "Paragraphs(1).Range.Font.Underline",
        "Paragraphs(1).Range.HighlightColorIndex",
        "Paragraphs(1).Style",
        "Paragraphs(1).Alignment",
        "Paragraphs(1).SpaceAfter",
        "Paragraphs(1).Next",
        "Range",
        "Styles",
        "BuiltInDocumentProperties",
        "Content.Paragraphs",
        "Application.Selection.TypeText",
        "Application.Selection.HomeKey",
        "Application.Documents",
    ];
    for m in members {
        for args in ["", " 1e300", " \"x\"", " -1, -1, -1, -1, -1", " 1, 1e308, 1, 1, 1", " 9999998", " -16777216"] {
            let src = format!(
                "Sub Main()\r\n    On Error Resume Next\r\n    ActiveDocument.{m}{args}\r\n    Debug.Print ActiveDocument.{m}\r\n    ActiveDocument.{m} = 1\r\nEnd Sub"
            );
            let _ = run_in(Doc::default(), &src);
        }
    }
}

#[test]
fn a_range_over_many_paragraphs_is_read_in_one_request() {
    let texts: Vec<String> = (0..300).map(|i| format!("Paragraph {i}")).collect();
    let refs: Vec<&str> = texts.iter().map(String::as_str).collect();
    let (r, d) = run_in(
        Doc::with_text(&refs),
        "Sub Main()\r\n    Debug.Print ActiveDocument.Content.ParagraphFormat.Alignment; ActiveDocument.Content.ParagraphFormat.SpaceAfter\r\n    ActiveDocument.Content.ParagraphFormat.SpaceAfter = 4\r\nEnd Sub",
    );
    r.unwrap();
    assert_eq!(d.logs, ["0 8"]);
    assert!(d.paras.iter().all(|p| p.space_after == 4.0));
    assert!(d.syncs <= 4, "{} syncs", d.syncs);
}
