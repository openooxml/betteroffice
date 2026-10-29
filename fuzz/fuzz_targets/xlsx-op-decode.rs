#![no_main]

use libfuzzer_sys::fuzz_target;
use xlsx_ops::{Op, OracleOp};

fn check_op(direct: serde_json::Result<Op>, oracle: serde_json::Result<OracleOp>) {
    let direct = direct.map(|op| (serde_json::to_string(&op).unwrap(), op));
    let oracle = oracle.map(|op| (serde_json::to_string(&op).unwrap(), Op::from(op)));
    if let (Ok(direct), Ok(oracle)) = (&direct, &oracle) {
        assert_eq!(direct, oracle);
    }
    for (canonical, op) in [direct.ok(), oracle.ok()].into_iter().flatten() {
        assert_eq!(serde_json::from_str::<Op>(&canonical).unwrap(), op);
        assert_eq!(
            Op::from(serde_json::from_str::<OracleOp>(&canonical).unwrap()),
            op
        );
    }
}

fn check_ops(direct: serde_json::Result<Vec<Op>>, oracle: serde_json::Result<Vec<OracleOp>>) {
    let direct = direct.map(|ops| (serde_json::to_string(&ops).unwrap(), ops));
    let oracle = oracle.map(|ops| {
        (
            serde_json::to_string(&ops).unwrap(),
            ops.into_iter().map(Op::from).collect::<Vec<_>>(),
        )
    });
    if let (Ok(direct), Ok(oracle)) = (&direct, &oracle) {
        assert_eq!(direct, oracle);
    }
    for (canonical, ops) in [direct.ok(), oracle.ok()].into_iter().flatten() {
        assert_eq!(serde_json::from_str::<Vec<Op>>(&canonical).unwrap(), ops);
        assert_eq!(
            serde_json::from_str::<Vec<OracleOp>>(&canonical)
                .unwrap()
                .into_iter()
                .map(Op::from)
                .collect::<Vec<_>>(),
            ops
        );
    }
}

fuzz_target!(|data: &[u8]| {
    check_op(serde_json::from_slice(data), serde_json::from_slice(data));
    check_ops(serde_json::from_slice(data), serde_json::from_slice(data));
    if let Ok(text) = std::str::from_utf8(data) {
        check_op(serde_json::from_str(text), serde_json::from_str(text));
        check_ops(serde_json::from_str(text), serde_json::from_str(text));
    }
    if let Ok(value) = serde_json::from_slice::<serde_json::Value>(data) {
        check_op(
            serde_json::from_value(value.clone()),
            serde_json::from_value(value.clone()),
        );
        check_ops(
            serde_json::from_value(value.clone()),
            serde_json::from_value(value),
        );
    }
});
