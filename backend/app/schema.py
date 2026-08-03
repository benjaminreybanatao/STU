"""
Data model + field registry for CRE Deal Screen.

Every deal tracks three provenance buckets:
  - om_facts:      pulled from the OM PDF (narrative + basic deal facts)
  - model_facts:   pulled from named cells in the underwriting Excel model
  - analyst_inputs: free text typed by the user to fill remaining gaps

FIELD_REGISTRY is the single source of truth for what the two-pager needs.
Gap analysis and deck generation both walk this list so they can never
disagree about what's required or where a number is allowed to come from.
"""
from dataclasses import dataclass, field
from enum import Enum


class Provenance(str, Enum):
    OM = "om"
    MODEL = "model"
    ANALYST = "analyst"
    MISSING = "missing"


class FieldGroup(str, Enum):
    PROPERTY = "property"
    TENANCY = "tenancy"
    PRICING = "pricing"
    CAPITAL_STACK = "capital_stack"
    RETURNS = "returns"
    ASSUMPTIONS = "assumptions"
    SOURCES_USES = "sources_uses"
    PER_UNIT = "per_unit"


@dataclass
class FieldSpec:
    key: str
    label: str
    group: FieldGroup
    # Ordered list of buckets to check, first non-empty wins.
    # Narrative/property facts prefer the OM; the firm's own underwritten
    # numbers prefer the model (it's the authoritative source for returns).
    source_priority: tuple = ("om", "model", "analyst")
    unit: str = ""  # display hint: "$", "%", "x", "yrs", etc.
    required: bool = True


FIELD_REGISTRY = [
    # --- Property ---
    FieldSpec("address", "Property Address", FieldGroup.PROPERTY, ("om", "analyst")),
    FieldSpec("year_built", "Year Built / Renovated", FieldGroup.PROPERTY, ("om", "analyst")),
    FieldSpec("sf_or_units", "Size (SF or Units)", FieldGroup.PROPERTY, ("om", "model", "analyst")),
    FieldSpec("property_type", "Property Type", FieldGroup.PROPERTY, ("om", "analyst")),
    FieldSpec("submarket_desc", "Submarket Description", FieldGroup.PROPERTY, ("om", "analyst"), required=False),

    # --- Tenancy ---
    FieldSpec("tenant_summary", "Tenant / Rent Roll Summary", FieldGroup.TENANCY, ("om", "model", "analyst")),
    FieldSpec("occupancy", "Occupancy", FieldGroup.TENANCY, ("om", "model", "analyst")),
    FieldSpec("occupancy_at_exit", "Occupancy at Exit", FieldGroup.TENANCY, ("model", "analyst"),
              unit="%", required=False),
    FieldSpec("walt", "WALT", FieldGroup.TENANCY, ("model", "om", "analyst"), unit="yrs", required=False),

    # --- Pricing ---
    FieldSpec("purchase_price", "Purchase Price", FieldGroup.PRICING, ("model", "om", "analyst"), unit="$"),
    FieldSpec("peak_cost", "Peak Cost", FieldGroup.PRICING, ("model", "analyst"), unit="$", required=False),
    FieldSpec("exit_price", "Exit Price (Gross)", FieldGroup.PRICING, ("model", "analyst"), unit="$", required=False),
    FieldSpec("going_in_cap", "Going-In Cap Rate", FieldGroup.PRICING, ("model", "om", "analyst"), unit="%"),
    FieldSpec("market_cap", "Cap Rate on Market Rents", FieldGroup.PRICING, ("model", "analyst"), unit="%",
              required=False),
    FieldSpec("exit_cap", "Exit Cap Rate", FieldGroup.PRICING, ("model", "analyst"), unit="%"),

    # --- Capital stack ---
    FieldSpec("leverage", "Leverage (LTV)", FieldGroup.CAPITAL_STACK, ("model", "analyst"), unit="%"),
    FieldSpec("debt_rate", "Debt Rate / Spread", FieldGroup.CAPITAL_STACK, ("model", "analyst"), unit="%"),
    FieldSpec("gross_debt_proceeds", "Gross Debt Proceeds", FieldGroup.CAPITAL_STACK, ("model", "analyst"),
              unit="$", required=False),
    FieldSpec("initial_equity", "Initial Equity", FieldGroup.CAPITAL_STACK, ("model", "analyst"), unit="$"),
    FieldSpec("peak_equity", "Peak Equity", FieldGroup.CAPITAL_STACK, ("model", "analyst"), unit="$", required=False),

    # --- Returns ---
    FieldSpec("hold_period", "Hold Period", FieldGroup.RETURNS, ("model", "om", "analyst"), unit="yrs"),
    FieldSpec("unlevered_irr", "Unlevered IRR", FieldGroup.RETURNS, ("model", "analyst"), unit="%"),
    FieldSpec("unlevered_em", "Unlevered Equity Multiple", FieldGroup.RETURNS, ("model", "analyst"), unit="x"),
    FieldSpec("levered_irr", "Levered IRR", FieldGroup.RETURNS, ("model", "analyst"), unit="%"),
    FieldSpec("levered_em", "Levered Equity Multiple", FieldGroup.RETURNS, ("model", "analyst"), unit="x"),
    FieldSpec("cash_on_cash", "Avg Cash-on-Cash", FieldGroup.RETURNS, ("model", "analyst"), unit="%", required=False),

    # --- Sources & Uses at Close (right-hand exhibit table) ---
    FieldSpec("total_sources", "Total Sources", FieldGroup.SOURCES_USES, ("model", "analyst"), unit="$", required=False),
    FieldSpec("dd_closing_costs", "DD / Closing Costs", FieldGroup.SOURCES_USES, ("model", "analyst"), unit="$", required=False),
    FieldSpec("working_capital", "Working Capital", FieldGroup.SOURCES_USES, ("model", "analyst"), unit="$", required=False),
    FieldSpec("equity_subtotal", "Equity Subtotal", FieldGroup.SOURCES_USES, ("model", "analyst"), unit="$", required=False),
    FieldSpec("financing_cost", "Financing Cost", FieldGroup.SOURCES_USES, ("model", "analyst"), unit="$", required=False),
    FieldSpec("total_uses", "Total Uses", FieldGroup.SOURCES_USES, ("model", "analyst"), unit="$", required=False),

    # --- Assumptions (woven into the slide 1 narrative) ---
    FieldSpec("lease_term_assumption", "Lease Term Assumption", FieldGroup.ASSUMPTIONS, ("model", "om", "analyst"), required=False),
    FieldSpec("downtime_assumption", "Downtime Assumption", FieldGroup.ASSUMPTIONS, ("model", "om", "analyst"), required=False),
    FieldSpec("exit_assumption", "Exit Assumption", FieldGroup.ASSUMPTIONS, ("model", "om", "analyst"), required=False),
]

# Per-unit companions for every money row the exhibit tables show a "$ Per Unit"
# column for. These are read from an explicit per-unit column in the model --
# never computed by dividing by a unit count -- so a model that doesn't carry
# them leaves the column showing the placeholder rather than an invented figure.
PER_UNIT_SOURCE_KEYS = [
    "purchase_price",
    "peak_cost",
    "exit_price",
    "initial_equity",
    "peak_equity",
    "gross_debt_proceeds",
    "dd_closing_costs",
    "working_capital",
    "equity_subtotal",
    "financing_cost",
    "total_sources",
    "total_uses",
]

_BASE_LABELS = {spec.key: spec.label for spec in FIELD_REGISTRY}

FIELD_REGISTRY += [
    FieldSpec(
        f"{key}_per_unit",
        f"{_BASE_LABELS[key]} / Unit",
        FieldGroup.PER_UNIT,
        ("model", "analyst"),
        unit="$",
        required=False,
    )
    for key in PER_UNIT_SOURCE_KEYS
]

FIELD_BY_KEY = {f.key: f for f in FIELD_REGISTRY}

PLACEHOLDER = "TBD — confirm with sponsor"


@dataclass
class ResolvedField:
    key: str
    label: str
    group: str
    value: str | None
    provenance: str  # Provenance enum value
    required: bool
    unit: str = ""


@dataclass
class Deal:
    om_facts: dict = field(default_factory=dict)
    model_facts: dict = field(default_factory=dict)
    analyst_inputs: dict = field(default_factory=dict)
    om_images: list = field(default_factory=list)  # list of dicts: {path, width, height, page}

    def resolve(self, spec: FieldSpec) -> ResolvedField:
        buckets = {
            "om": self.om_facts,
            "model": self.model_facts,
            "analyst": self.analyst_inputs,
        }
        for source in spec.source_priority:
            val = buckets[source].get(spec.key)
            if val not in (None, ""):
                return ResolvedField(spec.key, spec.label, spec.group.value, val, source, spec.required, spec.unit)
        return ResolvedField(spec.key, spec.label, spec.group.value, None, Provenance.MISSING.value, spec.required, spec.unit)

    def gap_analysis(self) -> list[ResolvedField]:
        return [self.resolve(spec) for spec in FIELD_REGISTRY]

    def display_value(self, spec_key: str) -> str:
        spec = FIELD_BY_KEY[spec_key]
        rf = self.resolve(spec)
        if rf.value is None:
            return PLACEHOLDER
        return str(rf.value)
