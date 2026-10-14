"""Pending edits and rendered-preview metadata."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class ProposalChange:
    slide_id: str
    shape_id: str | None
    before: dict[str, Any] | None
    after: dict[str, Any] | None
    old_text: str
    new_text: str

    @classmethod
    def _from_dict(cls, value: dict[str, Any]) -> ProposalChange:
        return cls(
            value["slideId"], value["shapeId"], value["before"], value["after"],
            value["oldText"], value["newText"],
        )


@dataclass(frozen=True)
class Proposal:
    id: str
    agent_id: str
    note: str | None
    edits: tuple[dict[str, Any], ...]
    changes: tuple[ProposalChange, ...]
    stale_targets: tuple[str, ...]

    @classmethod
    def _from_dict(cls, value: dict[str, Any]) -> Proposal:
        return cls(
            value["id"], value["agentId"], value["note"], tuple(value["edits"]),
            tuple(ProposalChange._from_dict(change) for change in value["changes"]),
            tuple(value["staleTargets"]),
        )


@dataclass(frozen=True)
class ProposalPreview:
    proposal: Proposal
    snapshot: dict[str, Any]

    @classmethod
    def _from_dict(cls, value: dict[str, Any]) -> ProposalPreview:
        return cls(Proposal._from_dict(value["proposal"]), value["snapshot"])
