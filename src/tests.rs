//! Macros end to end against a workbook in memory that answers as GridCraft does.

use serde_json::{Value as Json, json};

use crate::fakehost::Book;
use crate::host::{Host, HostError};
use crate::vba::{Failure, Interp, Program};

pub struct FakeHost(pub Book);

impl Host for FakeHost {
    fn call(&mut self, method: &str, params: Json) -> Result<Json, HostError> {
        self.0.handle(method, &params).map_err(|(code, message, data)| HostError { code, message, data })
    }
}

/// Macros run with a big stack (the add-in's own thread has one too).
pub fn big_stack<T: Send + 'static>(f: impl FnOnce() -> T + Send + 'static) -> T {
    std::thread::Builder::new().stack_size(crate::STACK).spawn(f).unwrap().join().unwrap()
}

fn run_in(book: Book, src: &str, name: &str) -> (Result<(), Failure>, Book) {
    let (src, name) = (src.to_string(), name.to_string());
    big_stack(move || run_in_here(book, &src, &name))
}

fn run_in_here(book: Book, src: &str, name: &str) -> (Result<(), Failure>, Book) {
    let prog = Program::new(&[("Module1".to_string(), src.to_string())]);
    assert!(prog.broken.is_empty(), "{:?}", prog.broken);
    let mut host = FakeHost(book);
    let r = {
        let mut it = Interp::new(&prog, &mut host);
        match it.find(name) {
            Ok((m, p)) => it.run(m, p),
            Err(e) => panic!("{e}"),
        }
    };
    (r, host.0)
}

fn run(src: &str) -> Book {
    let (r, book) = run_in(Book::default(), src, "Main");
    if let Err(f) = r {
        panic!("{} (line {})", f.text(), f.line);
    }
    book
}

fn fails(src: &str) -> Failure {
    run_in(Book::default(), src, "Main").0.unwrap_err()
}

#[test]
fn values_formulas_and_formats_reach_the_sheet() {
    let b = run(r#"
Sub Main()
    Range("A1").Value = 5
    Range("A2") = "text"
    Cells(3, 1).Value = 2.5
    Range("B1").Formula = "=A1*2"
    With Range("A1:B2")
        .Font.Bold = True
        .Interior.Color = RGB(255, 0, 0)
        .NumberFormat = "0.00"
    End With
    Range("C1").Font.Color = vbBlue
    Range("C2").Value = Range("A1").Value + Range("A3").Value
    ActiveSheet.Range("D1").Value = Range("A1").Address
End Sub"#);
    assert_eq!(b.get("Sheet1", "A1"), json!(5));
    assert_eq!(b.get("Sheet1", "A2"), json!("text"));
    assert_eq!(b.get("Sheet1", "A3"), json!(2.5));
    assert_eq!(b.cell("Sheet1", "B1").formula.as_deref(), Some("=A1*2"));
    let a1 = b.cell("Sheet1", "A1");
    assert!(a1.bold);
    assert_eq!(a1.fill, "#FF0000");
    assert_eq!(a1.numfmt.as_deref(), Some("0.00"));
    assert_eq!(b.cell("Sheet1", "C1").color, "#0000FF");
    assert_eq!(b.get("Sheet1", "C2"), json!(7.5));
    assert_eq!(b.get("Sheet1", "D1"), json!("$A$1"));
}

#[test]
fn loops_functions_and_byref() {
    let b = run(r#"
Option Explicit
Private total As Long

Function Twice(ByVal x As Double) As Double
    Twice = x * 2
End Function

Sub AddTo(ByRef acc As Long, n As Long)
    acc = acc + n
End Sub

Sub Main()
    Dim i As Long, s As String
    For i = 1 To 10
        Cells(i, 1).Value = Twice(i)
        AddTo total, i
    Next i
    Range("B1").Value = total
    i = 0
    Do While i < 5
        i = i + 1
        If i = 3 Then Exit Do
    Loop
    Range("B2").Value = i
    For i = 10 To 1 Step -3
        s = s & i & ","
    Next
    Range("B3").Value = s
    Dim arr(1 To 3) As Integer, k As Integer
    For k = 1 To 3: arr(k) = k * k: Next k
    Range("B4").Value = arr(1) + arr(2) + arr(3)
    Dim words() As String
    words = Split("a b c", " ")
    Range("B5").Value = UBound(words) & ":" & Join(words, "-")
    Select Case total
        Case Is > 100: Range("B6").Value = "big"
        Case 50 To 60: Range("B6").Value = "fifty-five"
        Case Else: Range("B6").Value = "other"
    End Select
End Sub"#);
    assert_eq!(b.get("Sheet1", "A10"), json!(20.0));
    assert_eq!(b.get("Sheet1", "B1"), json!(55));
    assert_eq!(b.get("Sheet1", "B2"), json!(3));
    assert_eq!(b.get("Sheet1", "B3"), json!("10,7,4,1,"));
    assert_eq!(b.get("Sheet1", "B4"), json!(14));
    assert_eq!(b.get("Sheet1", "B5"), json!("2:a-b-c"));
    assert_eq!(b.get("Sheet1", "B6"), json!("fifty-five"));
}

#[test]
fn reading_cells_reads_the_sheet_once() {
    let mut book = Book::default();
    for r in 1..=50 {
        book.put("Sheet1", &format!("A{r}"), json!(r));
    }
    let (r, b) = run_in(
        book,
        r#"
Sub Main()
    Dim i As Long, sum As Double
    For i = 1 To 50
        sum = sum + Cells(i, 1).Value
    Next
    For i = 1 To 50
        Cells(i, 2).Value = Cells(i, 1).Value * 10
    Next
    Range("C1").Value = sum
End Sub"#,
        "Main",
    );
    r.unwrap();
    assert_eq!(b.get("Sheet1", "C1"), json!(1275.0));
    assert_eq!(b.get("Sheet1", "B50"), json!(500.0));
    // One read of the used range, then the writes go together (values written are known).
    assert!(b.syncs <= 3, "{} requests", b.syncs);
}

#[test]
fn errors_are_vbas_and_handlers_catch_them() {
    let f = fails("Sub Main()\r\n    Dim x As Integer\r\n    x = 40000\r\nEnd Sub");
    assert_eq!((f.number, f.line), (6, 3));
    assert_eq!(f.text(), "Run-time error '6': Overflow");
    let f = fails("Sub Main()\r\n    x = 1 / 0\r\nEnd Sub");
    assert_eq!(f.number, 11);
    let f = fails("Sub Main()\r\n    Worksheets(\"Nope\").Range(\"A1\").Value = 1\r\nEnd Sub");
    assert_eq!(f.number, 9);
    assert_eq!(f.text(), "Run-time error '9': Subscript out of range");
    let f = fails("Option Explicit\r\nSub Main()\r\n    y = 1\r\nEnd Sub");
    assert!(f.compile && f.text().contains("Variable not defined"), "{}", f.text());
    let f = fails("Sub Main()\r\n    NoSuchThing 1\r\nEnd Sub");
    assert!(f.text().contains("Sub or Function not defined"), "{}", f.text());
    let f = fails("Sub Main()\r\n    Err.Raise 1000 + vbObjectError, , \"custom\"\r\nEnd Sub");
    assert_eq!(f.message, "custom");

    let b = run(r#"
Sub Main()
    On Error GoTo Fail
    Range("A1").Value = "before"
    Dim x As Integer
    x = 1 / 0
    Range("A2").Value = "not reached"
    Exit Sub
Fail:
    Range("A3").Value = Err.Number & " " & Err.Description
    Resume Done
Done:
    Range("A4").Value = Err.Number
End Sub"#);
    assert_eq!(b.get("Sheet1", "A1"), json!("before"));
    assert_eq!(b.get("Sheet1", "A2"), Json::Null);
    assert_eq!(b.get("Sheet1", "A3"), json!("11 Division by zero"));
    assert_eq!(b.get("Sheet1", "A4"), json!(0));

    let b = run(r#"
Sub Main()
    On Error Resume Next
    x = 1 / 0
    Range("A1").Value = Err.Number
    Err.Clear
    Range("A2").Value = Err.Number
    On Error GoTo 0
End Sub"#);
    assert_eq!(b.get("Sheet1", "A1"), json!(11));
    assert_eq!(b.get("Sheet1", "A2"), json!(0));
}

#[test]
fn runaway_macros_stop() {
    let f = fails("Sub Main()\r\n    Do\r\n    Loop\r\nEnd Sub");
    assert!(f.message.contains("ran too long"), "{}", f.text());
    let f = fails("Sub Main()\r\n    Main\r\nEnd Sub");
    assert_eq!(f.number, 28, "{}", f.text());
}

#[test]
fn sheets_messages_and_logs() {
    let (r, b) = run_in(
        Book::with_sheets(&["Data", "Other"]),
        r#"
Sub Main()
    Dim ws As Worksheet, n As Integer
    For Each ws In Worksheets
        n = n + 1
        ws.Range("A1").Value = ws.Name & n
    Next
    Set ws = Worksheets.Add
    ws.Name = "Report"
    ws.Range("A1").Value = Worksheets.Count
    Worksheets("Data").Activate
    ActiveSheet.Range("B1").Value = ActiveSheet.Index
    MsgBox "Done: " & Worksheets.Count, vbInformation, "Report"
    Debug.Print "count"; Worksheets.Count
End Sub"#,
        "Module1.Main",
    );
    r.unwrap();
    assert_eq!(b.get("Data", "A1"), json!("Data1"));
    assert_eq!(b.get("Other", "A1"), json!("Other2"));
    assert_eq!(b.get("Report", "A1"), json!(3));
    assert_eq!(b.get("Data", "B1"), json!(1));
    assert_eq!(b.messages, vec![("Done: 3".to_string(), Some("Report".to_string()))]);
    assert_eq!(b.logs, vec!["count 3".to_string()]);
}

#[test]
fn ranges_arrays_end_and_for_each() {
    let mut book = Book::default();
    for (i, v) in [3, 1, 4, 1, 5].iter().enumerate() {
        book.put("Sheet1", &format!("A{}", i + 1), json!(v));
    }
    let (r, b) = run_in(
        book,
        r#"
Sub Main()
    Dim last As Long, c As Range, total As Double, v As Variant
    last = Cells(Rows.Count, 1).End(xlUp).Row
    Range("B1").Value = last
    For Each c In Range("A1:A" & last)
        total = total + c.Value
    Next
    Range("B2").Value = total
    v = Range("A1:A3").Value
    Range("B3").Value = UBound(v, 1) & "x" & UBound(v, 2) & "=" & v(3, 1)
    Range("C1:D2").Value = Array(7, 8)
    Range("B4").Value = WorksheetFunction.Sum(Range("A1:A5"))
    Range("B5").Value = Range("A1").Offset(2, 0).Value
    Range("B6").Value = Range("A1").Resize(2, 3).Address
    Dim col As New Collection
    col.Add "x", "k1"
    col.Add "y"
    Range("B7").Value = col.Count & col("k1") & col(2)
End Sub"#,
        "Main",
    );
    r.unwrap();
    assert_eq!(b.get("Sheet1", "B1"), json!(5));
    assert_eq!(b.get("Sheet1", "B2"), json!(14.0));
    assert_eq!(b.get("Sheet1", "B3"), json!("3x1=4"));
    assert_eq!(b.get("Sheet1", "C1"), json!(7));
    assert_eq!(b.get("Sheet1", "D2"), json!(8), "a 1-D array repeats down the rows");
    assert_eq!(b.get("Sheet1", "B4"), json!(14.0));
    assert_eq!(b.get("Sheet1", "B5"), json!(4.0));
    assert_eq!(b.get("Sheet1", "B6"), json!("$A$1:$C$2"));
    assert_eq!(b.get("Sheet1", "B7"), json!("2xy"));
}

#[test]
fn the_selection_and_clearing() {
    let mut book = Book::default();
    book.put("Sheet1", "A1", json!(1));
    book.put("Sheet1", "A2", json!(2));
    book.selection = (0, 0, 0, 1, 0);
    let (r, b) = run_in(
        book,
        r#"
Sub Main()
    Dim c As Range
    For Each c In Selection
        c.Value = c.Value * 100
    Next
    Range("B1").Value = Selection.Address
    Range("A2").ClearContents
    Range("C1").Select
    ActiveCell.Value = "here"
End Sub"#,
        "Main",
    );
    r.unwrap();
    assert_eq!(b.get("Sheet1", "A1"), json!(100.0));
    assert_eq!(b.get("Sheet1", "A2"), Json::Null);
    assert_eq!(b.get("Sheet1", "B1"), json!("$A$1:$A$2"));
    assert_eq!(b.get("Sheet1", "C1"), json!("here"));
}

#[test]
fn macros_are_found_by_name() {
    let prog = Program::new(&[
        ("Module1".to_string(), "Sub Hello()\r\nEnd Sub\r\nPrivate Sub Hidden()\r\nEnd Sub".to_string()),
        ("Module2".to_string(), "Sub Hello()\r\nEnd Sub".to_string()),
        ("Broken".to_string(), "Sub X(\r\nEnd Sub".to_string()),
    ]);
    assert_eq!(prog.broken.len(), 1);
    let mut host = FakeHost(Book::default());
    let it = Interp::new(&prog, &mut host);
    assert_eq!(it.find("Module2.Hello"), Ok((1, 0)));
    assert!(it.find("Hello").unwrap_err().contains("Ambiguous"));
    assert!(it.find("Nope").is_err());
    assert!(it.find("Broken.X").unwrap_err().contains("Compile error"));
}
