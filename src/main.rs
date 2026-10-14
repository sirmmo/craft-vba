//! `craft-vba`: the VBA add-in. `craft-vba --manifest` prints its `plugin.json`; otherwise it
//! answers the app on standard input and output.
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable, clippy::indexing_slicing)]
#![forbid(unsafe_code)]

fn main() {
    if std::env::args().any(|a| a == "--manifest") {
        println!("{:#}", craft_vba::plugin::manifest());
        return;
    }
    // Macros run with a big stack: deep calls are an error VBA reports, never a crash.
    let worker = std::thread::Builder::new().name("craft-vba".into()).stack_size(craft_vba::STACK).spawn(|| {
        let stdin = std::io::stdin();
        craft_vba::plugin::serve(stdin.lock(), std::io::stdout());
    });
    match worker {
        Ok(h) => {
            let _ = h.join();
        }
        Err(e) => eprintln!("craft-vba: can't start: {e}"),
    }
}
