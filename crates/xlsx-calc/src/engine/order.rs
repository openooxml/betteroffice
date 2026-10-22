//! the order a recalc evaluates its formulas in. the range graph is condensed
//! into strongly connected components: a formula on no cycle is a step of its
//! own, and the members of a component run together once everything they
//! read has. a range reaches the formulas it covers through a segment tree
//! over each row or column of formulas it spans, so a range over many
//! formulas costs a few edges rather than one per formula.

use std::collections::{HashMap, HashSet, VecDeque};

use xlsx_model::{CellRange, SheetId};

use super::{Key, key};
use crate::graph::DepGraph;

/// one unit of the evaluation order.
#[derive(Debug, PartialEq)]
pub(super) enum Step {
    /// a formula no cycle of the range graph passes through.
    Cell(Key),
    /// the members of one strongly connected component, in key order.
    Component(Vec<Key>),
}

/// order `recompute` so that each step follows every step it reads. ties go
/// to the lowest key, so a set without cycles keeps the order kahn's sort over
/// its cells gives it.
pub(super) fn topo_order(graph: &DepGraph, recompute: &HashSet<Key>) -> Vec<Step> {
    let network = Network::build(graph, recompute);
    let components = Components::find(&network.reads);
    network.sequence(&components)
}

/// the range graph over a recompute set: a node per formula, numbered in key
/// order, then the internal nodes of the segment trees. an edge runs from a
/// reader to what it reads.
struct Network {
    keys: Vec<Key>,
    /// formulas whose ranges cover their own cell.
    self_reads: Vec<bool>,
    reads: Adjacency,
}

impl Network {
    fn build(graph: &DepGraph, recompute: &HashSet<Key>) -> Self {
        let mut keys: Vec<Key> = recompute.iter().copied().collect();
        keys.sort_unstable();
        let mut edges = Edges {
            nodes: keys.len() as u32,
            pairs: Vec::new(),
            self_reads: vec![false; keys.len()],
        };
        {
            let mut lines = Lines {
                keys: &keys,
                columns: None,
                trees: HashMap::new(),
            };
            for (sheet, range, reader_sheet, reader_cell) in graph.edges() {
                xlsx_model::progress::tick();
                let Some(reader) = lines.id(key(reader_sheet, reader_cell)) else {
                    continue;
                };
                lines.link(reader, sheet, range, &mut edges);
                for (anchor_sheet, anchor) in graph.spill_sources(sheet, range) {
                    if let Some(source) = lines.id(key(anchor_sheet, anchor))
                        && source != reader
                    {
                        edges.push(reader, source);
                    }
                }
            }
        }
        Self {
            keys,
            self_reads: edges.self_reads,
            reads: Adjacency::new(edges.nodes as usize, &edges.pairs),
        }
    }

    /// kahn's sort over the components: one is ready once every component it
    /// reads has run, and ready ones run lowest key first. a component of tree
    /// nodes alone is no step, so it releases its readers as soon as it is
    /// ready itself.
    fn sequence(&self, components: &Components) -> Vec<Step> {
        let formulas = self.keys.len() as u32;
        let readers = self.reads.reversed();
        let mut waiting = vec![0usize; components.len()];
        for (node, &component) in components.of.iter().enumerate() {
            for &target in self.reads.of(node) {
                if components.of[target as usize] != component {
                    waiting[component as usize] += 1;
                }
            }
        }
        let lead: Vec<Option<u32>> = (0..components.len())
            .map(|component| {
                let nodes = components.nodes(component).iter().copied();
                nodes.filter(|&node| node < formulas).min()
            })
            .collect();
        let (mut ready, mut released): (Vec<usize>, Vec<usize>) = (0..components.len())
            .filter(|&component| waiting[component] == 0)
            .partition(|&component| lead[component].is_some());
        let mut queue = VecDeque::new();
        let mut steps = Vec::new();
        loop {
            while let Some(done) = released.pop() {
                for &node in components.nodes(done) {
                    for &reader in readers.of(node as usize) {
                        let component = components.of[reader as usize] as usize;
                        if component == done {
                            continue;
                        }
                        waiting[component] -= 1;
                        if waiting[component] > 0 {
                            continue;
                        }
                        match lead[component] {
                            Some(_) => ready.push(component),
                            None => released.push(component),
                        }
                    }
                }
            }
            ready.sort_unstable_by_key(|&component| lead[component]);
            queue.extend(ready.drain(..));
            let Some(next) = queue.pop_front() else {
                break;
            };
            xlsx_model::progress::tick();
            steps.push(self.step(components.nodes(next)));
            released.push(next);
        }
        steps
    }

    fn step(&self, nodes: &[u32]) -> Step {
        if let &[node] = nodes
            && !self.self_reads[node as usize]
        {
            return Step::Cell(self.keys[node as usize]);
        }
        let mut members: Vec<Key> = nodes
            .iter()
            .filter_map(|&node| self.keys.get(node as usize).copied())
            .collect();
        members.sort_unstable();
        Step::Component(members)
    }
}

/// the edges found so far, and how many nodes are numbered.
struct Edges {
    nodes: u32,
    pairs: Vec<(u32, u32)>,
    self_reads: Vec<bool>,
}

impl Edges {
    fn push(&mut self, from: u32, to: u32) {
        if from == to {
            self.self_reads[from as usize] = true;
        } else {
            self.pairs.push((from, to));
        }
    }
}

/// a row or column of formulas: its sheet, whether it is a row, and its index.
type LineKey = (SheetId, bool, u32);

/// the formulas of a recompute set along each row and column, and the segment
/// trees over the lines some range covers more than one formula of. formulas
/// are numbered in key order, so each row is a run of numbers; the order by
/// column is sorted once a range first needs it.
struct Lines<'k> {
    keys: &'k [Key],
    columns: Option<Vec<u32>>,
    trees: HashMap<LineKey, u32>,
}

impl Lines<'_> {
    fn id(&self, at: Key) -> Option<u32> {
        self.keys.binary_search(&at).ok().map(|id| id as u32)
    }

    /// link `reader` to the formulas `range` covers on `sheet`, line by line
    /// along the rows or the columns it spans fewer of.
    fn link(&mut self, reader: u32, sheet: SheetId, range: CellRange, edges: &mut Edges) {
        let (start, end) = (range.start, range.end);
        let keys = self.keys;
        if (start.row, start.col) == (end.row, end.col) {
            if let Some(id) = self.id((sheet, start.row, start.col)) {
                edges.push(reader, id);
            }
            return;
        }
        if end.col - start.col > end.row - start.row {
            let mut at = keys.partition_point(|&(s, row, _)| (s, row) < (sheet, start.row));
            while let Some(&(s, row, _)) = keys.get(at)
                && s == sheet
                && row <= end.row
            {
                let len = keys[at..].partition_point(|&(s, r, _)| (s, r) == (sheet, row));
                let run = &keys[at..at + len];
                let span = (
                    run.partition_point(|&(.., col)| col < start.col),
                    run.partition_point(|&(.., col)| col <= end.col),
                );
                let first = at as u32;
                let line = Line::new((sheet, true, row), len, |index: u32| first + index);
                line.link(reader, span, &mut self.trees, edges);
                at += len;
            }
            return;
        }
        let columns: &[u32] = self.columns.get_or_insert_with(|| by_column(keys));
        let column = |id: u32| {
            let (s, row, col) = keys[id as usize];
            (s, col, row)
        };
        let mut at = columns.partition_point(|&id| column(id) < (sheet, start.col, 0));
        while let Some(&id) = columns.get(at)
            && column(id).0 == sheet
            && column(id).1 <= end.col
        {
            let col = column(id).1;
            let len = columns[at..].partition_point(|&id| column(id) <= (sheet, col, u32::MAX));
            let run = &columns[at..at + len];
            let span = (
                run.partition_point(|&id| column(id).2 < start.row),
                run.partition_point(|&id| column(id).2 <= end.row),
            );
            let line = Line::new((sheet, false, col), len, |index: u32| run[index as usize]);
            line.link(reader, span, &mut self.trees, edges);
            at += len;
        }
    }
}

/// formula numbers ordered by sheet, column and row.
fn by_column(keys: &[Key]) -> Vec<u32> {
    let mut ids: Vec<u32> = (0..keys.len() as u32).collect();
    ids.sort_unstable_by_key(|&id| {
        let (sheet, row, col) = keys[id as usize];
        (sheet, col, row)
    });
    ids
}

/// one line of formulas and the number of each, in line order, for its
/// segment tree: internal index `i` covers indices `2i` and `2i + 1`, and the
/// formulas are the leaves from index `len` on.
struct Line<F> {
    key: LineKey,
    len: u32,
    leaf: F,
}

impl<F: Fn(u32) -> u32> Line<F> {
    fn new(key: LineKey, len: usize, leaf: F) -> Self {
        Self {
            key,
            len: len as u32,
            leaf,
        }
    }

    /// link `reader` to the formulas from `first` up to `last` along this
    /// line, through the tree nodes the span covers whole.
    fn link(
        &self,
        reader: u32,
        (first, last): (usize, usize),
        trees: &mut HashMap<LineKey, u32>,
        edges: &mut Edges,
    ) {
        let (first, last) = (first as u32, last as u32);
        match last - first {
            0 => {}
            1 => edges.push(reader, (self.leaf)(first)),
            _ => {
                let tree = *trees.entry(self.key).or_insert_with(|| self.tree(edges));
                let (mut low, mut high) = (first + self.len, last + self.len);
                while low < high {
                    if low & 1 == 1 {
                        edges.push(reader, self.node(tree, low));
                        low += 1;
                    }
                    if high & 1 == 1 {
                        high -= 1;
                        edges.push(reader, self.node(tree, high));
                    }
                    low >>= 1;
                    high >>= 1;
                }
            }
        }
    }

    fn tree(&self, edges: &mut Edges) -> u32 {
        let tree = edges.nodes;
        edges.nodes += self.len - 1;
        for index in 1..self.len {
            for child in [2 * index, 2 * index + 1] {
                edges.push(self.node(tree, index), self.node(tree, child));
            }
        }
        tree
    }

    fn node(&self, tree: u32, index: u32) -> u32 {
        match index.checked_sub(self.len) {
            Some(leaf) => (self.leaf)(leaf),
            None => tree + index - 1,
        }
    }
}

/// a graph's edges grouped by the node they leave.
struct Adjacency {
    starts: Vec<usize>,
    targets: Vec<u32>,
}

impl Adjacency {
    fn new(nodes: usize, pairs: &[(u32, u32)]) -> Self {
        let mut starts = vec![0; nodes + 1];
        for &(from, _) in pairs {
            starts[from as usize] += 1;
        }
        let mut total = 0;
        for start in &mut starts {
            let count = *start;
            *start = total;
            total += count;
        }
        let mut filled = starts.clone();
        let mut targets = vec![0; pairs.len()];
        for &(from, to) in pairs {
            targets[filled[from as usize]] = to;
            filled[from as usize] += 1;
        }
        Self { starts, targets }
    }

    fn reversed(&self) -> Self {
        let pairs: Vec<(u32, u32)> = (0..self.len())
            .flat_map(|from| self.of(from).iter().map(move |&to| (to, from as u32)))
            .collect();
        Self::new(self.len(), &pairs)
    }

    fn len(&self) -> usize {
        self.starts.len() - 1
    }

    fn of(&self, node: usize) -> &[u32] {
        &self.targets[self.starts[node]..self.starts[node + 1]]
    }
}

/// the strongly connected components of a graph.
struct Components {
    /// each node's component.
    of: Vec<u32>,
    /// every node, grouped by component.
    nodes: Vec<u32>,
    starts: Vec<usize>,
}

impl Components {
    /// tarjan's walk, kept on an explicit stack so no graph is too deep for it.
    fn find(graph: &Adjacency) -> Self {
        const UNSEEN: u32 = u32::MAX;
        let count = graph.len();
        let mut index = vec![UNSEEN; count];
        let mut low = vec![0; count];
        let mut on_stack = vec![false; count];
        let mut stack: Vec<u32> = Vec::new();
        let mut walk: Vec<(u32, usize)> = Vec::new();
        let mut next = 0;
        let mut found = Self {
            of: vec![0; count],
            nodes: Vec::with_capacity(count),
            starts: vec![0],
        };
        for root in 0..count as u32 {
            if index[root as usize] != UNSEEN {
                continue;
            }
            walk.push((root, 0));
            while let Some(top) = walk.last_mut() {
                let node = top.0 as usize;
                if index[node] == UNSEEN {
                    xlsx_model::progress::tick();
                    index[node] = next;
                    low[node] = next;
                    next += 1;
                    stack.push(top.0);
                    on_stack[node] = true;
                }
                if let Some(&target) = graph.of(node).get(top.1) {
                    top.1 += 1;
                    let target = target as usize;
                    if index[target] == UNSEEN {
                        walk.push((target as u32, 0));
                    } else if on_stack[target] {
                        low[node] = low[node].min(index[target]);
                    }
                    continue;
                }
                walk.pop();
                if let Some(&(parent, _)) = walk.last() {
                    low[parent as usize] = low[parent as usize].min(low[node]);
                }
                if low[node] == index[node] {
                    let component = (found.starts.len() - 1) as u32;
                    while let Some(member) = stack.pop() {
                        on_stack[member as usize] = false;
                        found.of[member as usize] = component;
                        found.nodes.push(member);
                        if member as usize == node {
                            break;
                        }
                    }
                    found.starts.push(found.nodes.len());
                }
            }
        }
        found
    }

    fn len(&self) -> usize {
        self.starts.len() - 1
    }

    fn nodes(&self, component: usize) -> &[u32] {
        &self.nodes[self.starts[component]..self.starts[component + 1]]
    }
}
