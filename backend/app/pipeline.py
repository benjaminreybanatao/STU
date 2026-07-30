"""Orchestrates OM + model parsing into a Deal, and Deal -> pptx."""
import copy
from pathlib import Path

from .parsing import om_parser, excel_parser
from .schema import Deal
from .pptx_builder import build_deck


def build_deal(om_path: str | None, xlsx_path: str | None, analyst_inputs: dict, work_dir: str) -> Deal:
    deal = Deal(analyst_inputs=dict(analyst_inputs or {}))

    if om_path:
        deal.om_facts = om_parser.extract_facts(om_path)
        images_dir = str(Path(work_dir) / "om_images")
        deal.om_images = om_parser.extract_images(om_path, images_dir)

    if xlsx_path:
        base_facts, downside_facts = excel_parser.extract_facts(xlsx_path)
        deal.model_facts = base_facts
        deal._downside_model_facts = downside_facts  # stashed for downside deal construction
    else:
        deal._downside_model_facts = {}

    return deal


def downside_variant(deal: Deal) -> Deal | None:
    if not deal._downside_model_facts:
        return None
    variant = copy.deepcopy(deal)
    # Downside sheet may not repeat every field (e.g. address); fall back to base model facts.
    merged = dict(deal.model_facts)
    merged.update(deal._downside_model_facts)
    variant.model_facts = merged
    return variant


def generate_deck(deal: Deal, output_path: str) -> str:
    hero_images = [img["path"] for img in deal.om_images[:2]]
    downside = downside_variant(deal)
    return build_deck(deal, hero_images, output_path, downside_deal=downside)
