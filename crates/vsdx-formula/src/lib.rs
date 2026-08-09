//! Shared bounded ShapeSheet formula core.

use std::collections::HashSet;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Unit {
    Number,
    Bool,
    Inches,
    Radians,
    Seconds,
}

pub fn unit(s: &str) -> Option<(Unit, f64)> {
    match s.to_ascii_lowercase().as_str() {
        "" => Some((Unit::Number, 1.)),
        "in" | "dl" => Some((Unit::Inches, 1.)),
        "cm" => Some((Unit::Inches, 1. / 2.54)),
        "mm" => Some((Unit::Inches, 1. / 25.4)),
        "pt" => Some((Unit::Inches, 1. / 72.)),
        "pica" => Some((Unit::Inches, 1. / 6.)),
        "ft" => Some((Unit::Inches, 12.)),
        "m" => Some((Unit::Inches, 100. / 2.54)),
        "deg" => Some((Unit::Radians, std::f64::consts::PI / 180.)),
        "rad" | "da" => Some((Unit::Radians, 1.)),
        "es" => Some((Unit::Seconds, 1.)),
        "em" => Some((Unit::Seconds, 60.)),
        "ed" => Some((Unit::Seconds, 86_400.)),
        "ew" => Some((Unit::Seconds, 604_800.)),
        "bool" => Some((Unit::Bool, 1.)),
        _ => None,
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum Expr {
    Number(f64, Unit),
    String(String),
    Reference(String),
    Unary(Box<Expr>),
    Binary(Box<Expr>, Op, Box<Expr>),
    Call(String, Vec<Expr>),
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Op {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Diagnostic {
    pub message: String,
}
#[derive(Clone, Copy, Debug)]
pub struct Limits {
    pub max_depth: usize,
    pub max_nodes: usize,
    pub max_tokens: usize,
}

#[derive(Clone, Debug, PartialEq)]
enum Tok {
    Number(f64, Unit),
    String(String),
    Ident(String),
    Op(Op),
    L,
    R,
    Comma,
    End,
    Invalid(String),
}
struct Parser<'a> {
    chars: std::iter::Peekable<std::str::Chars<'a>>,
    current: Tok,
    depth: usize,
    limits: Limits,
    nodes: usize,
    tokens: usize,
}
impl<'a> Parser<'a> {
    fn new(input: &'a str, limits: Limits) -> Self {
        let mut parser = Self {
            chars: input.chars().peekable(),
            current: Tok::End,
            depth: 0,
            limits,
            nodes: 0,
            tokens: 0,
        };
        parser.next();
        parser
    }
    fn parse(mut self) -> Result<Expr, Diagnostic> {
        let expr = self.cmp()?;
        if self.current != Tok::End {
            return Err(Diagnostic {
                message: "unexpected token".into(),
            });
        }
        Ok(expr)
    }
    fn node(&mut self) -> Result<(), Diagnostic> {
        self.nodes += 1;
        (self.nodes <= self.limits.max_nodes)
            .then_some(())
            .ok_or_else(|| Diagnostic {
                message: "formula AST node limit exceeded".into(),
            })
    }
    fn next(&mut self) {
        self.current = self.lex();
    }
    fn lex(&mut self) -> Tok {
        while self.chars.peek().is_some_and(|value| value.is_whitespace()) {
            self.chars.next();
        }
        let Some(character) = self.chars.next() else {
            return Tok::End;
        };
        self.tokens += 1;
        if self.tokens > self.limits.max_tokens {
            return Tok::Invalid("formula token limit exceeded".into());
        }
        match character {
            '(' => Tok::L,
            ')' => Tok::R,
            ',' => Tok::Comma,
            '+' => Tok::Op(Op::Add),
            '-' => Tok::Op(Op::Sub),
            '*' => Tok::Op(Op::Mul),
            '/' => Tok::Op(Op::Div),
            '^' => Tok::Op(Op::Pow),
            '=' => Tok::Op(Op::Eq),
            '<' => {
                if self.chars.next_if_eq(&'=').is_some() {
                    Tok::Op(Op::Le)
                } else if self.chars.next_if_eq(&'>').is_some() {
                    Tok::Op(Op::Ne)
                } else {
                    Tok::Op(Op::Lt)
                }
            }
            '>' => {
                if self.chars.next_if_eq(&'=').is_some() {
                    Tok::Op(Op::Ge)
                } else {
                    Tok::Op(Op::Gt)
                }
            }
            '"' => Tok::String(
                self.chars
                    .by_ref()
                    .take_while(|value| *value != '"')
                    .collect(),
            ),
            value if value.is_ascii_digit() || value == '.' => self.number(value),
            value => {
                let mut ident = value.to_string();
                while self.chars.peek().is_some_and(|next| {
                    next.is_ascii_alphanumeric() || matches!(*next, '.' | '!' | '_')
                }) {
                    ident.push(self.chars.next().unwrap());
                }
                Tok::Ident(ident)
            }
        }
    }
    fn number(&mut self, first: char) -> Tok {
        let mut value = first.to_string();
        while self
            .chars
            .peek()
            .is_some_and(|next| next.is_ascii_digit() || *next == '.')
        {
            value.push(self.chars.next().unwrap());
        }
        if self
            .chars
            .peek()
            .is_some_and(|next| matches!(*next, 'e' | 'E'))
        {
            let mut exponent = self.chars.clone();
            exponent.next();
            let digit = match exponent.next() {
                Some('+' | '-') => exponent.next(),
                other => other,
            };
            if digit.is_some_and(|next| next.is_ascii_digit()) {
                value.push(self.chars.next().unwrap());
                if self
                    .chars
                    .peek()
                    .is_some_and(|next| matches!(*next, '+' | '-'))
                {
                    value.push(self.chars.next().unwrap());
                }
                while self.chars.peek().is_some_and(|next| next.is_ascii_digit()) {
                    value.push(self.chars.next().unwrap());
                }
            }
        }
        let mut number = value.parse().unwrap_or(f64::NAN);
        while self.chars.peek().is_some_and(|next| next.is_whitespace()) {
            self.chars.next();
        }
        let mut suffix = String::new();
        while self
            .chars
            .peek()
            .is_some_and(|next| next.is_ascii_alphabetic())
        {
            suffix.push(self.chars.next().unwrap());
        }
        let Some((unit, scale)) = unit(&suffix) else {
            return Tok::Invalid(format!("unknown unit suffix {suffix}"));
        };
        number *= scale;
        Tok::Number(number, unit)
    }
    fn cmp(&mut self) -> Result<Expr, Diagnostic> {
        self.chain(Self::add, &[Op::Eq, Op::Ne, Op::Lt, Op::Le, Op::Gt, Op::Ge])
    }
    fn add(&mut self) -> Result<Expr, Diagnostic> {
        self.chain(Self::mul, &[Op::Add, Op::Sub])
    }
    fn mul(&mut self) -> Result<Expr, Diagnostic> {
        self.chain(Self::pow, &[Op::Mul, Op::Div])
    }
    fn chain(
        &mut self,
        parse: fn(&mut Self) -> Result<Expr, Diagnostic>,
        allowed: &[Op],
    ) -> Result<Expr, Diagnostic> {
        let mut left = parse(self)?;
        while let Tok::Op(op) = self.current {
            if !allowed.contains(&op) {
                break;
            }
            self.next();
            self.node()?;
            left = Expr::Binary(Box::new(left), op, Box::new(parse(self)?));
        }
        Ok(left)
    }
    fn pow(&mut self) -> Result<Expr, Diagnostic> {
        let left = self.primary()?;
        if self.current == Tok::Op(Op::Pow) {
            self.next();
            self.depth += 1;
            let right = self.pow()?;
            self.depth -= 1;
            self.node()?;
            Ok(Expr::Binary(Box::new(left), Op::Pow, Box::new(right)))
        } else {
            Ok(left)
        }
    }
    fn primary(&mut self) -> Result<Expr, Diagnostic> {
        if self.depth >= self.limits.max_depth {
            return Err(Diagnostic {
                message: "formula depth limit exceeded".into(),
            });
        }
        match self.current.clone() {
            Tok::Number(number, unit) => {
                self.next();
                self.node()?;
                Ok(Expr::Number(number, unit))
            }
            Tok::String(value) => {
                self.next();
                self.node()?;
                Ok(Expr::String(value))
            }
            Tok::Op(Op::Sub) => {
                self.next();
                self.depth += 1;
                let value = self.primary()?;
                self.depth -= 1;
                self.node()?;
                Ok(Expr::Unary(Box::new(value)))
            }
            Tok::Ident(name) => {
                self.next();
                if self.current != Tok::L {
                    self.node()?;
                    return Ok(Expr::Reference(name));
                }
                self.depth += 1;
                self.next();
                let mut args = Vec::new();
                if self.current != Tok::R {
                    loop {
                        args.push(self.cmp()?);
                        if self.current != Tok::Comma {
                            break;
                        }
                        self.next();
                    }
                }
                if self.current != Tok::R {
                    return Err(Diagnostic {
                        message: "expected ')'".into(),
                    });
                }
                self.next();
                self.depth -= 1;
                self.node()?;
                Ok(Expr::Call(name, args))
            }
            Tok::L => {
                self.depth += 1;
                self.next();
                let value = self.cmp()?;
                if self.current != Tok::R {
                    return Err(Diagnostic {
                        message: "expected ')'".into(),
                    });
                }
                self.next();
                self.depth -= 1;
                Ok(value)
            }
            Tok::Invalid(message) => Err(Diagnostic { message }),
            _ => Err(Diagnostic {
                message: "expected expression".into(),
            }),
        }
    }
}

pub fn parse(input: &str, limits: Limits) -> Result<Expr, Diagnostic> {
    if input.trim().eq_ignore_ascii_case("No Formula") {
        return Ok(Expr::Call("No Formula".into(), Vec::new()));
    }
    Parser::new(input, limits).parse()
}

pub fn evaluate_number(
    input: &str,
    limits: Limits,
    resolve: &mut impl FnMut(&str) -> Option<String>,
) -> Option<f64> {
    parse(input.trim_start_matches('='), limits)
        .ok()
        .and_then(|expr| NumberEngine::new(limits, resolve).evaluate(&expr))
}

struct NumberEngine<'a, F> {
    limits: Limits,
    resolve: &'a mut F,
    active: HashSet<String>,
}

impl<'a, F: FnMut(&str) -> Option<String>> NumberEngine<'a, F> {
    fn new(limits: Limits, resolve: &'a mut F) -> Self {
        Self {
            limits,
            resolve,
            active: HashSet::new(),
        }
    }

    fn evaluate(&mut self, expr: &Expr) -> Option<f64> {
        self.evaluate_at(expr, 0)
    }

    fn evaluate_at(&mut self, expr: &Expr, depth: usize) -> Option<f64> {
        if depth > self.limits.max_depth {
            return None;
        }
        match expr {
            Expr::Number(value, Unit::Number) => Some(*value),
            Expr::Reference(name) => {
                if !self.active.insert(name.clone()) {
                    return None;
                }
                let value = (self.resolve)(name)
                    .and_then(|formula| parse(formula.trim_start_matches('='), self.limits).ok())
                    .and_then(|formula| self.evaluate_at(&formula, depth + 1));
                self.active.remove(name);
                value
            }
            Expr::Unary(value) => Some(-self.evaluate_at(value, depth + 1)?),
            Expr::Binary(left, op, right) => {
                let left = self.evaluate_at(left, depth + 1)?;
                let right = self.evaluate_at(right, depth + 1)?;
                match op {
                    Op::Add => Some(left + right),
                    Op::Sub => Some(left - right),
                    Op::Mul => Some(left * right),
                    Op::Div if right != 0.0 => Some(left / right),
                    Op::Pow => Some(left.powf(right)),
                    _ => None,
                }
            }
            _ => None,
        }
        .filter(|value| value.is_finite())
    }
}
