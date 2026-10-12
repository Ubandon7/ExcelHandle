from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, ConfigDict, Field


def to_camel(value: str) -> str:
    head, *tail = value.split("_")
    return head + "".join(part.capitalize() for part in tail)


class APIModel(BaseModel):
    model_config = ConfigDict(
        alias_generator=to_camel,
        populate_by_name=True,
        extra="forbid",
    )


class ConditionOptions(APIModel):
    case_sensitive: bool = False
    trim_cell: bool = False
    trim_input: bool = True


class ColumnRef(APIModel):
    header_name: str
    fallback_column_index: int | None = Field(default=None, ge=1)
    occurrence: int = Field(default=1, ge=1)


class FilterCondition(APIModel):
    type: Literal["condition"] = "condition"
    column_ref: ColumnRef
    operator: Literal[
        "starts_with",
        "ends_with",
        "contains",
        "not_contains",
        "empty",
        "not_empty",
        "in",
        "not_in",
        "no_prefix",
        "no_suffix",
    ]
    value: str | None = None
    values: list[str] = Field(default_factory=list)
    options: ConditionOptions = Field(default_factory=ConditionOptions)


class ConditionGroup(APIModel):
    type: Literal["group"] = "group"
    relation: Literal["AND", "OR"] = "AND"
    children: list["ConditionNode"] = Field(default_factory=list)


ConditionNode = Annotated[Union[FilterCondition, ConditionGroup], Field(discriminator="type")]
ConditionGroup.model_rebuild()


class SourceScope(APIModel):
    mode: Literal["ALL", "SINGLE"] = "ALL"
    sheet_id: str | None = None


class FilterGroupCreate(APIModel):
    name: str | None = None
    source_scope: SourceScope = Field(default_factory=SourceScope)
    condition_tree: ConditionGroup = Field(default_factory=ConditionGroup)


class FilterGroupUpdate(APIModel):
    name: str | None = None
    source_scope: SourceScope | None = None
    condition_tree: ConditionGroup | None = None
    priority: int | None = Field(default=None, ge=1)
    saved: bool | None = None
    active: bool | None = None


class FilterRunRequest(APIModel):
    ordered_group_ids: list[str] = Field(min_length=1)


class ColumnSplitRunRequest(APIModel):
    column_ref: ColumnRef
    source_scope: SourceScope = Field(default_factory=SourceScope)


class AffixSplitRunRequest(APIModel):
    column_ref: ColumnRef
    source_scope: SourceScope = Field(default_factory=SourceScope)
    kind: Literal["PREFIX", "SUFFIX"]
    delimiter: str = Field(default="-", min_length=1, max_length=5)


class ColumnSelector(APIModel):
    match_mode: Literal["NORMALIZED_HEADER", "FIXED_POSITION"]
    header_name: str | None = None
    column_index: int | None = Field(default=None, ge=1)


class NumberPolicy(APIModel):
    parse_plain_numeric_text: bool = True
    strip_non_numeric_symbols: bool = True
    reject_text_characters: bool = True
    preserve_invalid_original_value: bool = True
    on_invalid: Literal["skip_and_report"] = "skip_and_report"


class SumSourceRef(APIModel):
    type: Literal["FILE", "FILTER_RESULT"]
    id: str = Field(min_length=1)


class SumRunRequest(APIModel):
    source_ref: SumSourceRef | None = None
    column_selector: ColumnSelector
    number_policy: NumberPolicy = Field(default_factory=NumberPolicy)
