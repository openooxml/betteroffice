//! the order a recalc evaluates its formulas in. the range graph is condensed
//! into strongly connected components: a formula on no cycle is a step of its
//! own, and the members of a component run together once everything they
//! read has. a range reaches the formulas it covers through a segment tree
//! over each row or column of formulas it spans, so a range over many
//! formulas costs a few edges rather than one per formula.

use std::collections::{BTreeMap, HashMap, HashSet, VecDeque};

use xlsx_model::SheetId;

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
        let ids: HashMap<Key, u32> = keys.iter().zip(0..).map(|(&k, id)| (k, id)).collect();
        let mut columns: BTreeMap<(SheetId, u32), Line> = BTreeMap::new();
        let mut rows: BTreeMap<(SheetId, u32), Line> = BTreeMap::new();
        for (&(sheet, row, col), id) in keys.iter().zip(0..) {
            columns.entry((sheet, col)).or_default().push(row, id);
            rows.entry((sheet, row)).or_default().push(col, id);
        }
        let mut edges = Edges {
            nodes: ids.len() as u32,
            pairs: Vec::new(),
            self_reads: vec![false; ids.len()],
        };
        for (sheet, range, reader_sheet, reader_cell) in graph.edges() {
            let Some(&reader) = ids.get(&key(reader_sheet, reader_cell)) else {
                continue;
            };
            let (start, end) = (range.start, range.end);
            let (lines, across, along) = if end.col - start.col <= end.row - start.row {
                (&mut columns, (start.col, end.col), (start.row, end.row))
            } else {
                (&mut rows, (start.row, end.row), (start.col, end.col))
            };
            for (_, line) in lines.range_mut((sheet, across.0)..=(sheet, across.1)) {
                line.link(reader, along, &mut edges);
            }
            for (anchor_sheet, anchor) in graph.spill_sources(sheet, range) {
                if let Some(&source) = ids.get(&key(anchor_sheet, anchor))
                    && source != reader
                {
                    edges.push(reader, source);
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
            steps.push(self.step(components.nodes(next)));
            released.push(next);
        }
        steps
    }

    fn step(&self, nodes: &[u32]) -> Step {
        let mut members: Vec<u32> = nodes
            .iter()
            .copied()
            .filter(|&node| (node as usize) < self.keys.len())
            .collect();
        if let ([member], [_]) = (members.as_slice(), nodes)
            && !self.self_reads[*member as usize]
        {
            return Step::Cell(self.keys[*member as usize]);
        }
        members.sort_unstable();
        Step::Component(members.iter().map(|&id| self.keys[id as usize]).collect())
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

/// the formulas along one row or column in order, and the first node of their
/// segment tree once a range needs one. internal index `i` covers indices `2i`
/// and `2i + 1`, and the formulas are the leaves from index `len` on.
#[derive(Default)]
struct Line {
    positions: Vec<u32>,
    formulas: Vec<u32>,
    tree: Option<u32>,
}

impl Line {
    fn push(&mut self, position: u32, formula: u32) {
        self.positions.push(position);
        self.formulas.push(formula);
    }

    /// link `reader` to the formulas from `first` to `last` along this line,
    /// through the tree nodes the span covers whole.
    fn link(&mut self, reader: u32, (first, last): (u32, u32), edges: &mut Edges) {
        let start = self.positions.partition_point(|&at| at < first);
        let end = self.positions.partition_point(|&at| at <= last);
        match end - start {
            0 => {}
            1 => edges.push(reader, self.formulas[start]),
            _ => {
                let tree = self.tree(edges);
                let leaves = self.formulas.len() as u32;
                let (mut low, mut high) = (start as u32 + leaves, end as u32 + leaves);
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

    fn tree(&mut self, edges: &mut Edges) -> u32 {
        if let Some(tree) = self.tree {
            return tree;
        }
        let tree = edges.nodes;
        let leaves = self.formulas.len() as u32;
        edges.nodes += leaves - 1;
        for index in 1..leaves {
            for child in [2 * index, 2 * index + 1] {
                edges.push(self.node(tree, index), self.node(tree, child));
            }
        }
        self.tree = Some(tree);
        tree
    }

    fn node(&self, tree: u32, index: u32) -> u32 {
        match index.checked_sub(self.formulas.len() as u32) {
            Some(leaf) => self.formulas[leaf as usize],
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
