# -*- coding: utf-8 -*-
"""Mapping engine — manages 科目编码 <-> 报表行项目 correspondence.

Mapping rules (3 types):
  code_match   — match source account code to a report line item
  auto_calc    — line is computed by calc engine (e.g. depreciation, tax surcharges)
  multi_source — line aggregates multiple code matches + computed items

Data stored as JSON per company, e.g. data/companies/<name>/mappings.json
"""
import json
import os


class MappingEngine:
    """Load and manage mappings for one company."""

    def __init__(self, company_dir):
        self.company_dir = company_dir
        self.mappings_path = os.path.join(company_dir, 'mappings.json')
        self._data = None

    def _load(self):
        if self._data is None:
            if os.path.exists(self.mappings_path):
                with open(self.mappings_path, 'r', encoding='utf-8') as f:
                    self._data = json.load(f)
            else:
                self._data = {'pl': [], 'bs': [], 'cf': []}
        return self._data

    def _save(self):
        os.makedirs(self.company_dir, exist_ok=True)
        with open(self.mappings_path, 'w', encoding='utf-8') as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    def get_mappings(self, statement):
        """Return list of mapping entries for a statement ('pl'|'bs'|'cf')."""
        return self._load().get(statement, [])

    def set_mappings(self, statement, mappings_list):
        """Replace all mappings for a statement type."""
        data = self._load()
        data[statement] = mappings_list
        self._save()

    def add_code_mapping(self, statement, report_line, account_codes):
        """Add or update a code_match mapping. Merges with existing codes."""
        data = self._load()
        mappings = data.setdefault(statement, [])

        codes_list = list(account_codes) if not isinstance(account_codes, list) else account_codes

        for m in mappings:
            if m['report_line'] == report_line and m['rule_type'] == 'code_match':
                existing = set(m.get('account_codes', []))
                existing.update(codes_list)
                m['account_codes'] = sorted(existing)
                self._save()
                return

        mappings.append({
            'report_line': report_line,
            'statement': statement,
            'rule_type': 'code_match',
            'account_codes': sorted(codes_list),
            'auto_params': {},
        })
        self._save()

    def set_auto_calc(self, statement, report_line, method, params=None):
        """Mark a report line as auto-calculated."""
        data = self._load()
        mappings = data.setdefault(statement, [])

        for m in mappings:
            if m['report_line'] == report_line:
                m['rule_type'] = 'auto_calc'
                m['auto_params'] = {'method': method, **(params or {})}
                m['account_codes'] = []
                self._save()
                return

        mappings.append({
            'report_line': report_line,
            'statement': statement,
            'rule_type': 'auto_calc',
            'account_codes': [],
            'auto_params': {'method': method, **(params or {})},
        })
        self._save()

    def set_multi_source(self, statement, report_line, sources):
        """Set a report line as multi-source aggregation.

        sources is a list of {type: 'code'|'auto', value: ...}
        """
        data = self._load()
        mappings = data.setdefault(statement, [])

        for m in mappings:
            if m['report_line'] == report_line:
                m['rule_type'] = 'multi_source'
                m['sources'] = sources
                m['account_codes'] = []
                self._save()
                return

        mappings.append({
            'report_line': report_line,
            'statement': statement,
            'rule_type': 'multi_source',
            'account_codes': [],
            'auto_params': {},
            'sources': sources,
        })
        self._save()

    def get_account_codes_for_line(self, statement, report_line):
        """Return list of account codes mapped to a specific report line."""
        for m in self.get_mappings(statement):
            if m['report_line'] == report_line:
                return m.get('account_codes', [])
        return []

    def get_all_known_codes(self, statement):
        """Return set of all account codes registered across all mappings."""
        codes = set()
        for m in self.get_mappings(statement):
            codes.update(m.get('account_codes', []))
        return codes

    def find_unmapped_codes(self, statement, discovered_codes):
        """Compare discovered codes against known mappings.
        Returns list of codes that have NO mapping yet.
        """
        known = self.get_all_known_codes(statement)
        return sorted(set(discovered_codes) - known)

    def to_dict(self):
        """Dump all mappings as dict for display/export."""
        return self._load()
