pub(crate) fn zip(parts: &[(&str, &[u8])], deflate: bool) -> Vec<u8> {
    let mut output = Vec::new();
    let mut central = Vec::new();
    let method = u16::from(deflate) * 8;
    for (name, data) in parts {
        let offset = output.len() as u32;
        let crc = crate::zip::crc32(data);
        let compressed = if deflate {
            let mut compressed = Vec::new();
            let count = data.len().max(1).div_ceil(65535);
            for index in 0..count {
                let start = index * 65535;
                let end = (start + 65535).min(data.len());
                let len = (end - start) as u16;
                compressed.push(u8::from(index + 1 == count));
                compressed.extend_from_slice(&len.to_le_bytes());
                compressed.extend_from_slice(&(!len).to_le_bytes());
                compressed.extend_from_slice(&data[start..end]);
            }
            compressed
        } else {
            data.to_vec()
        };
        output.extend_from_slice(b"PK\x03\x04");
        for value in [20u16, 0, method, 0, 0] {
            output.extend_from_slice(&value.to_le_bytes());
        }
        for value in [crc, compressed.len() as u32, data.len() as u32] {
            output.extend_from_slice(&value.to_le_bytes());
        }
        output.extend_from_slice(&(name.len() as u16).to_le_bytes());
        output.extend_from_slice(&0u16.to_le_bytes());
        output.extend_from_slice(name.as_bytes());
        output.extend_from_slice(&compressed);
        central.extend_from_slice(b"PK\x01\x02");
        for value in [20u16, 20, 0, method, 0, 0] {
            central.extend_from_slice(&value.to_le_bytes());
        }
        for value in [crc, compressed.len() as u32, data.len() as u32] {
            central.extend_from_slice(&value.to_le_bytes());
        }
        for value in [name.len() as u16, 0, 0, 0, 0] {
            central.extend_from_slice(&value.to_le_bytes());
        }
        central.extend_from_slice(&0u32.to_le_bytes());
        central.extend_from_slice(&offset.to_le_bytes());
        central.extend_from_slice(name.as_bytes());
    }
    let offset = output.len() as u32;
    output.extend_from_slice(&central);
    output.extend_from_slice(b"PK\x05\x06");
    for value in [0u16, 0, parts.len() as u16, parts.len() as u16] {
        output.extend_from_slice(&value.to_le_bytes());
    }
    output.extend_from_slice(&(central.len() as u32).to_le_bytes());
    output.extend_from_slice(&offset.to_le_bytes());
    output.extend_from_slice(&0u16.to_le_bytes());
    output
}
