"""Pending edits and rendered-preview metadata."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class ProposalChange:
    slide_id: str
    shape_id: str | None
    before: dict[str, Any] | None
    after: dict[str, Any] | None
    old_text: str
    new_text: str
    kind: str | None = None
    before_index: int | None = None
    after_index: int | None = None
    before_index_present: bool = False
    after_index_present: bool = False

    @classmethod
    def _from_dict(cls, value: dict[str, Any]) -> ProposalChange:
        return cls(
            value["slideId"], value["shapeId"], value["before"], value["after"],
            value["oldText"], value["newText"],
            value.get("kind"), value.get("beforeIndex"), value.get("afterIndex"),
            "beforeIndex" in value, "afterIndex" in value,
        )


@dataclass(frozen=True)
class Proposal:
    id: str
    agent_id: str
    note: str | None
    edits: tuple[dict[str, Any], ...]
    changes: tuple[ProposalChange, ...]
    stale_targets: tuple[str, ...]
    refs: dict[str, str] = field(default_factory=dict)

    @classmethod
    def _from_dict(cls, value: dict[str, Any]) -> Proposal:
        return cls(
            value["id"], value["agentId"], value["note"], tuple(value["edits"]),
            tuple(ProposalChange._from_dict(change) for change in value["changes"]),
            tuple(value["staleTargets"]),
            dict(value.get("refs", {})),
        )


@dataclass(frozen=True)
class ProposalStripEntry:
    slide_id: str
    status: str
    live_index: int | None
    preview_index: int | None

    @classmethod
    def _from_dict(cls, value: dict[str, Any]) -> ProposalStripEntry:
        return cls(
            value["slideId"], value["status"], value["liveIndex"], value["previewIndex"],
        )


@dataclass(frozen=True)
class ProposalPreview:
    proposal: Proposal
    snapshot: dict[str, Any]

    @classmethod
    def _from_dict(cls, value: dict[str, Any]) -> ProposalPreview:
        return cls(Proposal._from_dict(value["proposal"]), value["snapshot"])
