use std::cell::Cell;
use std::future::poll_fn;
use std::rc::Rc;
use std::task::Poll;

#[derive(Clone, Default)]
pub struct WorkBudget {
    remaining: Rc<Cell<usize>>,
    touched: Rc<Cell<usize>>,
}

impl WorkBudget {
    pub fn reset(&self, units: usize) {
        self.remaining.set(units);
        self.touched.set(0);
    }

    pub fn touched(&self) -> usize {
        self.touched.get()
    }

    pub async fn step(&self) {
        self.take(1).await;
    }

    pub async fn copy_bytes(&self, source: &[u8]) -> Vec<u8> {
        let mut copied = Vec::with_capacity(source.len());
        self.append_bytes(&mut copied, source).await;
        copied
    }

    pub async fn append_bytes(&self, target: &mut Vec<u8>, source: &[u8]) {
        let mut offset = 0;
        while offset < source.len() {
            let count = (self.take(256).await * 64).min(source.len() - offset);
            target.extend_from_slice(&source[offset..offset + count]);
            offset += count;
        }
    }

    pub async fn equal_bytes(&self, left: &[u8], right: &[u8]) -> bool {
        if left.len() != right.len() {
            return false;
        }
        let mut offset = 0;
        while offset < left.len() {
            let count = (self.take(256).await * 64).min(left.len() - offset);
            if left[offset..offset + count] != right[offset..offset + count] {
                return false;
            }
            offset += count;
        }
        true
    }

    pub async fn clone_slice<T: Clone>(&self, source: &[T]) -> Vec<T> {
        let mut copied = Vec::with_capacity(source.len());
        for value in source {
            self.step().await;
            copied.push(value.clone());
        }
        copied
    }

    pub async fn take(&self, limit: usize) -> usize {
        poll_fn(|_| {
            let count = self.remaining.get().min(limit);
            if count == 0 {
                return Poll::Pending;
            }
            self.remaining.set(self.remaining.get() - count);
            self.touched.set(self.touched.get() + count);
            Poll::Ready(count)
        })
        .await
    }
}
