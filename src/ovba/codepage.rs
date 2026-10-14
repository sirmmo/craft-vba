//! The project's code page (`PROJECTCODEPAGE`): module names and sources are stored in it.
//! Windows-1252 and UTF-8 are decoded exactly; other code pages are read as 1252 (lossy).

/// Windows-1252's 0x80–0x9F (the rest is Latin-1). Unassigned bytes map to the same C1 control
/// point, as Windows' own conversion does.
const CP1252_HIGH: [char; 32] = [
    '\u{20AC}', '\u{0081}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}', '\u{2020}', '\u{2021}', '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}',
    '\u{008D}', '\u{017D}', '\u{008F}', '\u{0090}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}', '\u{2022}', '\u{2013}', '\u{2014}', '\u{02DC}', '\u{2122}',
    '\u{0161}', '\u{203A}', '\u{0153}', '\u{009D}', '\u{017E}', '\u{0178}',
];

pub const UTF8: u16 = 65001;
pub const WINDOWS_1252: u16 = 1252;

pub fn decode(bytes: &[u8], code_page: u16) -> String {
    if code_page == UTF8 {
        return String::from_utf8_lossy(bytes).into_owned();
    }
    bytes
        .iter()
        .map(|&b| match b {
            0x80..=0x9F => CP1252_HIGH.get(usize::from(b - 0x80)).copied().unwrap_or('?'),
            _ => char::from(b),
        })
        .collect()
}

/// Encodes `text`; characters the code page lacks become `?`.
pub fn encode(text: &str, code_page: u16) -> Vec<u8> {
    if code_page == UTF8 {
        return text.as_bytes().to_vec();
    }
    text.chars()
        .map(|c| {
            let u = u32::from(c);
            if u < 0x80 || (0xA0..=0xFF).contains(&u) { u as u8 } else { CP1252_HIGH.iter().position(|&x| x == c).map(|i| 0x80 + i as u8).unwrap_or(b'?') }
        })
        .collect()
}

pub fn utf16le(text: &str) -> Vec<u8> {
    text.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

pub fn from_utf16le(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes.chunks_exact(2).map(|p| u16::from_le_bytes([p.first().copied().unwrap_or(0), p.get(1).copied().unwrap_or(0)])).collect();
    String::from_utf16_lossy(&units)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn windows_1252_round_trips() {
        let s = "Café — “quoted” €5 Œuvre";
        assert_eq!(decode(&encode(s, WINDOWS_1252), WINDOWS_1252), s);
        assert_eq!(encode("日本", WINDOWS_1252), b"??");
        assert_eq!(decode(&[0x80, 0x41, 0xE9], 1252), "€Aé");
        assert_eq!(decode("ü".as_bytes(), UTF8), "ü");
        assert_eq!(from_utf16le(&utf16le("Modül1")), "Modül1");
    }
}
