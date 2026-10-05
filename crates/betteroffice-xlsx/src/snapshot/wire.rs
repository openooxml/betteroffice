use super::{SnapshotError, SnapshotResult};

pub(crate) fn reserve<T>(values: &mut Vec<T>, count: usize) -> SnapshotResult<()> {
    values
        .try_reserve_exact(count)
        .map_err(|_| SnapshotError::new("cannot allocate snapshot records"))
}

pub(crate) const FORMAT_VERSION: u8 = 1;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub(crate) enum ChunkKind {
    Header = 1,
    AuthorityBase = 2,
    Yrs = 3,
    Model = 4,
    Cells = 5,
    Preserved = 6,
    Facts = 7,
    Source = 8,
    End = 9,
}

impl ChunkKind {
    fn from_u8(tag: u8) -> SnapshotResult<Self> {
        Ok(match tag {
            1 => Self::Header,
            2 => Self::AuthorityBase,
            3 => Self::Yrs,
            4 => Self::Model,
            5 => Self::Cells,
            6 => Self::Preserved,
            7 => Self::Facts,
            8 => Self::Source,
            9 => Self::End,
            _ => {
                return Err(SnapshotError::new(format!(
                    "unknown snapshot chunk kind {tag}"
                )));
            }
        })
    }
}

pub(crate) fn frame(kind: ChunkKind, ordinal: u64, payload: &[u8]) -> Vec<u8> {
    let mut w = Writer::with_capacity(payload.len() + 12);
    w.u8(FORMAT_VERSION);
    w.u8(kind as u8);
    w.var_u64(ordinal);
    w.raw(payload);
    w.into_bytes()
}

pub(crate) fn unframe(chunk: &[u8]) -> SnapshotResult<(ChunkKind, u64, &[u8])> {
    let mut r = Reader::new(chunk);
    let version = r.u8()?;
    if version != FORMAT_VERSION {
        return Err(SnapshotError::new(format!(
            "unsupported snapshot format {version}"
        )));
    }
    let kind = ChunkKind::from_u8(r.u8()?)?;
    let ordinal = r.var_u64()?;
    Ok((kind, ordinal, r.rest()))
}

#[derive(Debug, Default)]
pub(crate) struct Writer {
    buf: Vec<u8>,
}

impl Writer {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    pub(crate) fn with_capacity(capacity: usize) -> Self {
        Self {
            buf: Vec::with_capacity(capacity),
        }
    }

    pub(crate) fn len(&self) -> usize {
        self.buf.len()
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.buf.is_empty()
    }

    pub(crate) fn into_bytes(self) -> Vec<u8> {
        self.buf
    }

    pub(crate) fn u8(&mut self, value: u8) {
        self.buf.push(value);
    }

    pub(crate) fn bool(&mut self, value: bool) {
        self.buf.push(u8::from(value));
    }

    pub(crate) fn var_u64(&mut self, mut value: u64) {
        while value >= 0x80 {
            self.buf.push((value as u8) | 0x80);
            value >>= 7;
        }
        self.buf.push(value as u8);
    }

    pub(crate) fn var_i64(&mut self, value: i64) {
        self.var_u64(((value << 1) ^ (value >> 63)) as u64);
    }

    pub(crate) fn var_u32(&mut self, value: u32) {
        self.var_u64(u64::from(value));
    }

    pub(crate) fn var_usize(&mut self, value: usize) {
        self.var_u64(value as u64);
    }

    pub(crate) fn f64(&mut self, value: f64) {
        self.buf.extend_from_slice(&value.to_bits().to_le_bytes());
    }

    pub(crate) fn bytes(&mut self, value: &[u8]) {
        self.var_usize(value.len());
        self.buf.extend_from_slice(value);
    }

    pub(crate) fn str(&mut self, value: &str) {
        self.bytes(value.as_bytes());
    }

    pub(crate) fn raw(&mut self, value: &[u8]) {
        self.buf.extend_from_slice(value);
    }

    pub(crate) fn option<T>(&mut self, value: Option<T>, write: impl FnOnce(&mut Self, T)) {
        match value {
            None => self.u8(0),
            Some(value) => {
                self.u8(1);
                write(self, value);
            }
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct Reader<'a> {
    buf: &'a [u8],
    pos: usize,
}

impl<'a> Reader<'a> {
    pub(crate) fn new(buf: &'a [u8]) -> Self {
        Self { buf, pos: 0 }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.pos == self.buf.len()
    }

    pub(crate) fn finish(self) -> SnapshotResult<()> {
        if self.is_empty() {
            Ok(())
        } else {
            Err(SnapshotError::new(format!(
                "snapshot payload has {} trailing bytes",
                self.buf.len() - self.pos
            )))
        }
    }

    pub(crate) fn rest(self) -> &'a [u8] {
        &self.buf[self.pos..]
    }

    fn take(&mut self, n: usize) -> SnapshotResult<&'a [u8]> {
        let end = self
            .pos
            .checked_add(n)
            .filter(|end| *end <= self.buf.len())
            .ok_or_else(|| SnapshotError::new("snapshot payload is truncated"))?;
        let out = &self.buf[self.pos..end];
        self.pos = end;
        Ok(out)
    }

    pub(crate) fn u8(&mut self) -> SnapshotResult<u8> {
        Ok(self.take(1)?[0])
    }

    pub(crate) fn bool(&mut self) -> SnapshotResult<bool> {
        match self.u8()? {
            0 => Ok(false),
            1 => Ok(true),
            tag => Err(SnapshotError::new(format!("invalid snapshot bool {tag}"))),
        }
    }

    pub(crate) fn var_u64(&mut self) -> SnapshotResult<u64> {
        let mut value = 0u64;
        for shift in (0..64).step_by(7) {
            let byte = self.u8()?;
            let bits = u64::from(byte & 0x7f);
            if shift == 63 && bits > 1 {
                return Err(SnapshotError::new("snapshot varint overflows u64"));
            }
            value |= bits << shift;
            if byte & 0x80 == 0 {
                return Ok(value);
            }
        }
        Err(SnapshotError::new("snapshot varint overflows u64"))
    }

    pub(crate) fn var_i64(&mut self) -> SnapshotResult<i64> {
        let value = self.var_u64()?;
        Ok(((value >> 1) as i64) ^ -((value & 1) as i64))
    }

    pub(crate) fn var_u32(&mut self) -> SnapshotResult<u32> {
        u32::try_from(self.var_u64()?)
            .map_err(|_| SnapshotError::new("snapshot varint overflows u32"))
    }

    pub(crate) fn var_usize(&mut self) -> SnapshotResult<usize> {
        usize::try_from(self.var_u64()?)
            .map_err(|_| SnapshotError::new("snapshot varint overflows usize"))
    }

    pub(crate) fn f64(&mut self) -> SnapshotResult<f64> {
        let bytes: [u8; 8] = self.take(8)?.try_into().expect("eight bytes");
        Ok(f64::from_bits(u64::from_le_bytes(bytes)))
    }

    pub(crate) fn bytes(&mut self) -> SnapshotResult<&'a [u8]> {
        let len = self.var_usize()?;
        self.take(len)
    }

    pub(crate) fn str(&mut self) -> SnapshotResult<&'a str> {
        std::str::from_utf8(self.bytes()?)
            .map_err(|_| SnapshotError::new("snapshot string is not UTF-8"))
    }

    pub(crate) fn option<T>(
        &mut self,
        read: impl FnOnce(&mut Self) -> SnapshotResult<T>,
    ) -> SnapshotResult<Option<T>> {
        match self.u8()? {
            0 => Ok(None),
            1 => read(self).map(Some),
            tag => Err(SnapshotError::new(format!(
                "invalid snapshot option tag {tag}"
            ))),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn primitives_roundtrip_exactly() {
        let mut w = Writer::new();
        w.var_u64(u64::MAX);
        w.var_i64(i64::MIN);
        w.var_i64(-1);
        w.var_u32(300);
        w.f64(-0.0);
        w.f64(f64::from_bits(0x7ff8_0000_0000_0001));
        w.str("é");
        w.option(None::<u32>, |w, v| w.var_u32(v));
        w.option(Some(0u32), |w, v| w.var_u32(v));
        w.bool(true);
        let bytes = w.into_bytes();
        let mut r = Reader::new(&bytes);
        assert_eq!(r.var_u64().unwrap(), u64::MAX);
        assert_eq!(r.var_i64().unwrap(), i64::MIN);
        assert_eq!(r.var_i64().unwrap(), -1);
        assert_eq!(r.var_u32().unwrap(), 300);
        assert_eq!(r.f64().unwrap().to_bits(), (-0.0f64).to_bits());
        assert_eq!(r.f64().unwrap().to_bits(), 0x7ff8_0000_0000_0001);
        assert_eq!(r.str().unwrap(), "é");
        assert_eq!(r.option(|r| r.var_u32()).unwrap(), None);
        assert_eq!(r.option(|r| r.var_u32()).unwrap(), Some(0));
        assert!(r.bool().unwrap());
        r.finish().unwrap();
    }

    #[test]
    fn truncated_and_malformed_input_is_rejected() {
        assert!(Reader::new(&[0x80]).var_u64().is_err());
        assert!(Reader::new(&[0xff; 11]).var_u64().is_err());
        assert!(Reader::new(&[5, b'a']).bytes().is_err());
        assert!(Reader::new(&[2]).bool().is_err());
        assert!(Reader::new(&[1, 2]).finish().is_err());
        assert!(unframe(&[FORMAT_VERSION + 1, 1, 0]).is_err());
        assert!(unframe(&[FORMAT_VERSION, 0, 0]).is_err());
    }

    #[test]
    fn frames_roundtrip() {
        let chunk = frame(ChunkKind::Cells, 7, b"abc");
        let (kind, ordinal, payload) = unframe(&chunk).unwrap();
        assert_eq!((kind, ordinal, payload), (ChunkKind::Cells, 7, &b"abc"[..]));
    }
}
