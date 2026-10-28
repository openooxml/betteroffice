use std::io;

pub fn synthetic(seed: u32) -> Vec<u8> {
    let mut state = seed;
    let mut block = Vec::with_capacity(16384);
    for _ in 0..16384 {
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        block.push(state as u8);
    }
    block.repeat(4)
}

pub fn compress(bytes: &[u8], dict: Option<&[u8]>) -> Vec<u8> {
    let mut compressor = zstd::bulk::Compressor::default();
    compressor
        .set_parameter(zstd::zstd_safe::CParameter::CompressionLevel(19))
        .unwrap();
    compressor.include_checksum(true).unwrap();
    compressor.include_contentsize(true).unwrap();
    compressor.window_log(20).unwrap();
    if let Some(dict) = dict {
        compressor
            .context_mut()
            .ref_prefix(dict)
            .map_err(|code| io::Error::other(zstd::zstd_safe::get_error_name(code)))
            .unwrap();
    }
    compressor.compress(bytes).unwrap()
}
