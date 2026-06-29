# -*- coding: utf-8 -*-
"""Template engine — loads preset template configs and manages template selection."""
import json
import os

TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), 'templates')

PRESET_MAP = {
    'pl': 'pl_config.json',
    'bs': 'bs_config.json',
    'cf': 'cf_config.json',
}

def _load_json(filename):
    path = os.path.join(TEMPLATE_DIR, filename)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Template file not found: {path}")
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def load_preset(statement_type):
    """Load a preset template config by type.

    Args:
        statement_type: 'pl' | 'bs' | 'cf'

    Returns:
        dict with keys: template_type, subtitle, ncols, col_widths,
        header_rows, data_start_row, total_rows, columns, value_col_labels,
        merge_ranges, items
    """
    filename = PRESET_MAP.get(statement_type)
    if filename is None:
        raise ValueError(f"Unknown statement type '{statement_type}'. Use 'pl', 'bs', or 'cf'.")
    return _load_json(filename)

def list_presets():
    """Return list of available preset template types."""
    return list(PRESET_MAP.keys())

def get_item_names(statement_type):
    """Return ordered list of item names from a preset template."""
    config = load_preset(statement_type)
    if statement_type == 'bs':
        left = [item['left'] for item in config['items'] if item['left']]
        right = [item['right'] for item in config['items'] if item['right']]
        return left, right
    names = [item['name'] for item in config['items'] if item['name']]
    return names, []
