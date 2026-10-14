//! The compression [MS-OVBA] §2.4.1 uses for the `dir` stream and module sources: a signature
//! byte, then chunks of at most 4096 decompressed bytes, each either stored raw (exactly 4096
//! bytes) or as token sequences of literals and copy tokens.

/// Largest decompressed size accepted (hostile input can't make us allocate more).
pub const MAX_DECOMPRESSED: usize = 64 << 20;

const CHUNK: usize = 4096;

/// How a copy token splits into offset and length, from the position in the chunk.
fn copy_token_help(decompressed_in_chunk: usize) -> (u16, u16, u32) {
    // The smallest bit count whose 2^count covers the distance back to the chunk's start; at least 4.
    let mut bit_count = 4u32;
    while bit_count < 12 && (1usize << bit_count) < decompressed_in_chunk {
        bit_count += 1;
    }
    let length_mask = 0xFFFFu16 >> bit_count;
    let offset_mask = !length_mask;
    (length_mask, offset_mask, bit_count)
}

/// Decompresses a compressed container.
pub fn decompress(data: &[u8]) -> Result<Vec<u8>, String> {
    let mut out = Vec::new();
    let Some((&sig, mut rest)) = data.split_first() else { return Err("empty compressed container".into()) };
    if sig != 1 {
        return Err("not a compressed container (bad signature byte)".into());
    }
    while let Some(&[h0, h1]) = rest.get(..2) {
        let header = u16::from_le_bytes([h0, h1]);
        let size = usize::from(header & 0x0FFF) + 3;
        let compressed = header & 0x8000 != 0;
        let chunk = rest.get(2..size.min(rest.len())).unwrap_or(&[]);
        rest = rest.get(size..).unwrap_or(&[]);
        let chunk_start = out.len();
        if !compressed {
            // A raw chunk is 4096 bytes; a short one at the end is taken as it is.
            out.extend_from_slice(chunk.get(..CHUNK.min(chunk.len())).unwrap_or(&[]));
        } else {
            let mut i = 0;
            'chunk: while let Some(&flags) = chunk.get(i) {
                i += 1;
                for bit in 0..8 {
                    let Some(&byte) = chunk.get(i) else { break 'chunk };
                    if flags & (1 << bit) == 0 {
                        out.push(byte);
                        i += 1;
                    } else {
                        let Some(&[t0, t1]) = chunk.get(i..i + 2) else { break 'chunk };
                        let token = u16::from_le_bytes([t0, t1]);
                        i += 2;
                        let (length_mask, offset_mask, bit_count) = copy_token_help(out.len() - chunk_start);
                        let length = usize::from(token & length_mask) + 3;
                        let offset = usize::from((token & offset_mask) >> (16 - bit_count)) + 1;
                        if offset > out.len() - chunk_start {
                            return Err("a copy token points before its chunk".into());
                        }
                        let from = out.len() - offset;
                        for k in 0..length {
                            let b = out.get(from + k).copied().unwrap_or(0);
                            out.push(b);
                        }
                    }
                    if out.len() > MAX_DECOMPRESSED {
                        return Err("decompresses to more than 64 MiB".into());
                    }
                }
            }
        }
        if out.len() > MAX_DECOMPRESSED {
            return Err("decompresses to more than 64 MiB".into());
        }
    }
    Ok(out)
}

/// Compresses `data` into a compressed container.
pub fn compress(data: &[u8]) -> Vec<u8> {
    let mut out = vec![1u8];
    for chunk in data.chunks(CHUNK) {
        let compressed = compress_chunk(chunk);
        // A chunk that doesn't shrink is stored raw, which only a full 4096-byte chunk may be.
        if compressed.len() > CHUNK && chunk.len() == CHUNK {
            let header: u16 = (0b011 << 12) | (CHUNK as u16 + 2 - 3);
            out.extend_from_slice(&header.to_le_bytes());
            out.extend_from_slice(chunk);
        } else {
            let header: u16 = 0x8000 | (0b011 << 12) | ((compressed.len() + 2 - 3) as u16 & 0x0FFF);
            out.extend_from_slice(&header.to_le_bytes());
            out.extend_from_slice(&compressed);
        }
    }
    out
}

/// The token sequences of one chunk (without its header).
fn compress_chunk(chunk: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(chunk.len() + chunk.len() / 8 + 1);
    let mut pos = 0;
    while pos < chunk.len() {
        let flag_at = out.len();
        out.push(0);
        let mut flags = 0u8;
        for bit in 0..8 {
            if pos >= chunk.len() {
                break;
            }
            let (length_mask, _, bit_count) = copy_token_help(pos);
            let max_length = usize::from(length_mask) + 3;
            let max_offset = 1usize << bit_count;
            // The longest earlier match within reach (greedy; the window is at most 4096 bytes).
            let (mut best_len, mut best_off) = (0usize, 0usize);
            let start = pos.saturating_sub(max_offset);
            for cand in start..pos {
                let mut len = 0;
                while len < max_length && chunk.get(pos + len).is_some_and(|b| chunk.get(cand + len) == Some(b)) {
                    len += 1;
                }
                if len > best_len {
                    best_len = len;
                    best_off = pos - cand;
                }
            }
            if best_len >= 3 {
                let token = (((best_off - 1) as u16) << (16 - bit_count)) | ((best_len - 3) as u16);
                out.extend_from_slice(&token.to_le_bytes());
                flags |= 1 << bit;
                pos += best_len;
            } else if let Some(&b) = chunk.get(pos) {
                out.push(b);
                pos += 1;
            }
        }
        if let Some(f) = out.get_mut(flag_at) {
            *f = flags;
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trips() {
        let mut big = Vec::new();
        for i in 0..20_000u32 {
            big.extend_from_slice(format!("Line {} of a module\r\n", i % 97).as_bytes());
        }
        let noise: Vec<u8> = (0..9000u32).map(|i| (i.wrapping_mul(2_654_435_761) >> 13) as u8).collect();
        for data in [&b""[..], b"a", b"abc", b"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa", &big, &noise] {
            let c = compress(data);
            assert_eq!(decompress(&c).unwrap(), data, "len {}", data.len());
        }
        assert!(compress(&big).len() < big.len() / 4, "text compresses");
    }

    #[test]
    fn copy_tokens_use_the_spec_bit_counts() {
        // The bit count grows with the position in the chunk: 4 bits up to 16 bytes, 12 at the end.
        assert_eq!(copy_token_help(1).2, 4);
        assert_eq!(copy_token_help(16).2, 4);
        assert_eq!(copy_token_help(17).2, 5);
        assert_eq!(copy_token_help(4096).2, 12);
        // "abcabcabc": three literals, then one copy token (offset 3, length 6) in a single
        // token sequence: flags 0b1000, a b c, token.
        let c = compress(b"abcabcabc");
        let (_, _, bits) = copy_token_help(3);
        let token = ((3u16 - 1) << (16 - bits)) | (6 - 3);
        let mut expected = vec![1u8];
        let body = [0b1000, b'a', b'b', b'c', token.to_le_bytes()[0], token.to_le_bytes()[1]];
        expected.extend_from_slice(&(0xB000u16 | (body.len() as u16 + 2 - 3)).to_le_bytes());
        expected.extend_from_slice(&body);
        assert_eq!(c, expected);
    }

    #[test]
    fn hostile_containers_are_errors() {
        assert!(decompress(&[]).is_err());
        assert!(decompress(&[2, 0, 0]).is_err());
        // A copy token before anything was written.
        assert!(decompress(&[1, 0x02, 0xB0, 0x01, 0x00, 0x00]).is_err());
        // Truncated chunks stop without panicking.
        let mut c = compress(b"hello hello hello hello");
        c.truncate(c.len() - 2);
        let _ = decompress(&c);
        for len in 0..64 {
            let junk: Vec<u8> = std::iter::once(1).chain((0..len).map(|i| (i * 37 + 11) as u8)).collect();
            let _ = decompress(&junk);
        }
    }
}
