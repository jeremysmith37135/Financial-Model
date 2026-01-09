"""
Mapping Persistence for Multi-Division Financial Model
Handles saving/loading account mappings to JSON files and Excel hidden sheets

VERSION: 1.0.0
DATE: 2025-12-16
"""

import json
import os
from datetime import datetime
from typing import Dict, List, Optional, Any
from dataclasses import asdict


class MappingPersistence:
    """Handle loading/saving mappings to JSON and Excel"""

    @staticmethod
    def save_to_json(mappings: Dict, divisions: List[Dict], company_name: str, filepath: str) -> bool:
        """
        Save mappings to external JSON file

        Args:
            mappings: Dict of {statement_type: {account_name: AccountMapping}}
            divisions: List of division configuration dicts
            company_name: Company name for the file
            filepath: Full path to save the JSON file

        Returns:
            True if successful, False otherwise
        """
        try:
            # Convert mappings to serializable format
            serializable_mappings = {
                "pl": [],
                "bs": []
            }

            for stmt_type in ["pl", "bs"]:
                if stmt_type in mappings:
                    for name, mapping in mappings[stmt_type].items():
                        if hasattr(mapping, 'to_dict'):
                            serializable_mappings[stmt_type].append(mapping.to_dict())
                        elif isinstance(mapping, dict):
                            serializable_mappings[stmt_type].append(mapping)
                        else:
                            # Convert dataclass or object to dict
                            serializable_mappings[stmt_type].append(asdict(mapping))

            data = {
                "version": "1.0",
                "company": company_name,
                "created": datetime.now().isoformat(),
                "last_modified": datetime.now().isoformat(),
                "divisions": divisions,
                "account_mappings": serializable_mappings,
                "unmatched_accounts": {}  # Will be populated from mappings
            }

            # Extract unmatched accounts
            for stmt_type in ["pl", "bs"]:
                for mapping_data in serializable_mappings[stmt_type]:
                    if mapping_data.get("match_type") == "unmatched":
                        for div_name in mapping_data.get("division_mappings", {}).keys():
                            if div_name not in data["unmatched_accounts"]:
                                data["unmatched_accounts"][div_name] = []
                            original_name = mapping_data["division_mappings"][div_name]
                            if original_name not in data["unmatched_accounts"][div_name]:
                                data["unmatched_accounts"][div_name].append(original_name)

            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False, default=str)

            return True

        except Exception as e:
            print(f"Error saving mappings to JSON: {e}")
            return False

    @staticmethod
    def load_from_json(filepath: str) -> Optional[Dict]:
        """
        Load mappings from JSON file

        Args:
            filepath: Full path to the JSON file

        Returns:
            Dict with mappings data or None if file doesn't exist/error
        """
        try:
            if not os.path.exists(filepath):
                return None

            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)

            return data

        except Exception as e:
            print(f"Error loading mappings from JSON: {e}")
            return None

    @staticmethod
    def save_to_excel(wb, mappings: Dict, divisions: List[Dict]) -> bool:
        """
        Save mappings to hidden Excel sheet (Mapping_Config)

        Args:
            wb: xlwings Workbook object
            mappings: Dict of {statement_type: {account_name: AccountMapping}}
            divisions: List of division configuration dicts

        Returns:
            True if successful, False otherwise
        """
        try:
            sheet_name = 'Mapping_Config'

            # Create or get Mapping_Config sheet
            existing_sheets = [s.name for s in wb.sheets]
            if sheet_name in existing_sheets:
                sheet = wb.sheets[sheet_name]
                sheet.clear()
            else:
                sheet = wb.sheets.add(sheet_name, after=wb.sheets[-1])

            # Get division names for dynamic headers
            div_names = [d.get('name', d) if isinstance(d, dict) else d for d in divisions]

            # Build headers
            headers = ['Type', 'Consolidated']
            headers.extend([f'{d}_Account' for d in div_names])
            headers.extend(['Match_Type', 'Confidence', 'Approved', 'Reasoning', 'Updated'])

            # BUILD ALL DATA IN MEMORY FIRST (OPTIMIZED)
            all_rows = [headers]  # Start with headers

            for stmt_type in ['pl', 'bs']:
                if stmt_type not in mappings:
                    continue

                for name, mapping in mappings[stmt_type].items():
                    # Convert mapping to dict if needed
                    if hasattr(mapping, 'to_dict'):
                        m = mapping.to_dict()
                    elif isinstance(mapping, dict):
                        m = mapping
                    else:
                        m = asdict(mapping)

                    row_data = [
                        stmt_type.upper(),
                        m.get('consolidated_name', name)
                    ]

                    # Add division-specific account names
                    div_mappings = m.get('division_mappings', {})
                    for div_name in div_names:
                        row_data.append(div_mappings.get(div_name, ''))

                    # Add metadata
                    row_data.extend([
                        m.get('match_type', 'unknown'),
                        m.get('confidence', 1.0),
                        m.get('approved', False),
                        m.get('reasoning', ''),
                        m.get('approved_date', datetime.now().isoformat())
                    ])

                    all_rows.append(row_data)

            # WRITE ALL DATA IN ONE BULK OPERATION
            if all_rows:
                sheet.range((1, 1), (len(all_rows), len(headers))).value = all_rows

            # Format header row
            try:
                header_range = sheet.range((1, 1), (1, len(headers)))
                header_range.font.bold = True
                header_range.color = (22, 33, 62)  # Dark blue
                header_range.font.color = (255, 255, 255)  # White text
            except:
                pass

            # Auto-fit columns
            try:
                sheet.autofit()
            except:
                pass

            # Hide sheet (VeryHidden so it doesn't show in sheet tabs)
            try:
                sheet.api.Visible = 2  # xlSheetVeryHidden = 2
            except:
                # If can't hide, at least set tab color to indicate it's system sheet
                try:
                    sheet.api.Tab.Color = 0  # Black
                except:
                    pass

            return True

        except Exception as e:
            print(f"Error saving mappings to Excel: {e}")
            return False

    @staticmethod
    def load_from_excel(wb) -> Optional[Dict]:
        """
        Load mappings from hidden Excel sheet

        Args:
            wb: xlwings Workbook object

        Returns:
            Dict with mappings data or None if sheet doesn't exist
        """
        try:
            sheet_name = 'Mapping_Config'

            # Check if sheet exists
            existing_sheets = [s.name for s in wb.sheets]
            if sheet_name not in existing_sheets:
                return None

            sheet = wb.sheets[sheet_name]

            # Temporarily unhide to read
            try:
                original_visibility = sheet.api.Visible
                sheet.api.Visible = -1  # xlSheetVisible = -1
            except:
                pass

            # Read all data
            data = sheet.used_range.value
            if not data or len(data) < 2:
                return None

            headers = data[0]

            # Parse headers to find division columns
            div_columns = {}
            for i, header in enumerate(headers):
                if header and header.endswith('_Account'):
                    div_name = header.replace('_Account', '')
                    div_columns[div_name] = i

            # Find standard column indices
            type_col = headers.index('Type') if 'Type' in headers else 0
            cons_col = headers.index('Consolidated') if 'Consolidated' in headers else 1
            match_type_col = headers.index('Match_Type') if 'Match_Type' in headers else -1
            confidence_col = headers.index('Confidence') if 'Confidence' in headers else -1
            approved_col = headers.index('Approved') if 'Approved' in headers else -1
            reasoning_col = headers.index('Reasoning') if 'Reasoning' in headers else -1

            # Build mappings structure
            mappings = {
                'pl': {},
                'bs': {}
            }

            for row in data[1:]:  # Skip header row
                if not row or not row[type_col]:
                    continue

                stmt_type = str(row[type_col]).lower()
                if stmt_type not in ['pl', 'bs']:
                    continue

                cons_name = row[cons_col]
                if not cons_name:
                    continue

                # Build division mappings
                div_mappings = {}
                for div_name, col_idx in div_columns.items():
                    if col_idx < len(row) and row[col_idx]:
                        div_mappings[div_name] = row[col_idx]

                mapping = {
                    'consolidated_name': cons_name,
                    'division_mappings': div_mappings,
                    'match_type': row[match_type_col] if match_type_col >= 0 and match_type_col < len(row) else 'unknown',
                    'confidence': float(row[confidence_col]) if confidence_col >= 0 and confidence_col < len(row) and row[confidence_col] else 1.0,
                    'approved': bool(row[approved_col]) if approved_col >= 0 and approved_col < len(row) else False,
                    'reasoning': row[reasoning_col] if reasoning_col >= 0 and reasoning_col < len(row) else ''
                }

                mappings[stmt_type][cons_name] = mapping

            # Re-hide sheet
            try:
                sheet.api.Visible = original_visibility
            except:
                pass

            return {
                'account_mappings': mappings,
                'divisions': list(div_columns.keys())
            }

        except Exception as e:
            print(f"Error loading mappings from Excel: {e}")
            return None

    @staticmethod
    def get_json_filepath(excel_filepath: str) -> str:
        """
        Generate JSON filepath based on Excel file location

        Args:
            excel_filepath: Path to the Excel file

        Returns:
            Path for the corresponding JSON mappings file
        """
        base, _ = os.path.splitext(excel_filepath)
        return f"{base}_mappings.json"

    @staticmethod
    def merge_mappings(existing: Dict, new: Dict) -> Dict:
        """
        Merge new mappings into existing, preserving user approvals

        Args:
            existing: Existing mappings dictionary
            new: New mappings to merge

        Returns:
            Merged mappings dictionary
        """
        if not existing:
            return new

        merged = {
            'pl': {},
            'bs': {}
        }

        # Start with existing mappings
        for stmt_type in ['pl', 'bs']:
            if stmt_type in existing.get('account_mappings', {}):
                merged[stmt_type] = dict(existing['account_mappings'][stmt_type])

        # Add/update with new mappings
        for stmt_type in ['pl', 'bs']:
            if stmt_type not in new.get('account_mappings', {}):
                continue

            for name, mapping in new['account_mappings'][stmt_type].items():
                if name in merged[stmt_type]:
                    # Preserve user approval status from existing
                    existing_mapping = merged[stmt_type][name]
                    if existing_mapping.get('approved') and existing_mapping.get('match_type') == 'manual':
                        # Keep the existing manual/approved mapping
                        continue

                merged[stmt_type][name] = mapping

        return {'account_mappings': merged}

    @staticmethod
    def export_mappings_report(mappings: Dict, filepath: str) -> bool:
        """
        Export a human-readable mappings report

        Args:
            mappings: Mappings dictionary
            filepath: Path for the report file

        Returns:
            True if successful
        """
        try:
            lines = [
                "Account Mappings Report",
                "=" * 60,
                f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                "",
            ]

            for stmt_type in ['pl', 'bs']:
                if stmt_type not in mappings.get('account_mappings', {}):
                    continue

                lines.append(f"\n{'P&L' if stmt_type == 'pl' else 'Balance Sheet'} Mappings")
                lines.append("-" * 40)

                stmt_mappings = mappings['account_mappings'][stmt_type]

                # Group by match type
                by_type = {'exact': [], 'intelligent': [], 'manual': [], 'unmatched': []}
                for name, m in stmt_mappings.items():
                    match_type = m.get('match_type', 'unknown')
                    if match_type in by_type:
                        by_type[match_type].append((name, m))
                    else:
                        by_type['unmatched'].append((name, m))

                for match_type, items in by_type.items():
                    if not items:
                        continue

                    lines.append(f"\n  {match_type.upper()} MATCHES ({len(items)})")
                    for name, m in items:
                        lines.append(f"    {name}")
                        for div, acct in m.get('division_mappings', {}).items():
                            lines.append(f"      - {div}: {acct}")
                        if m.get('reasoning'):
                            lines.append(f"      Reason: {m['reasoning']}")
                        lines.append("")

            with open(filepath, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))

            return True

        except Exception as e:
            print(f"Error exporting mappings report: {e}")
            return False
