//! The app running the add-in, as macros see it: requests the plugin makes while the app waits
//! for one of its own (protocol 1.1's `host/…` methods).

use serde_json::Value as Json;

#[derive(Clone, Debug, PartialEq)]
pub struct HostError {
    pub code: i64,
    pub message: String,
    pub data: Json,
}

pub trait Host {
    /// Sends `method` to the app and waits for its answer.
    fn call(&mut self, method: &str, params: Json) -> Result<Json, HostError>;
}
