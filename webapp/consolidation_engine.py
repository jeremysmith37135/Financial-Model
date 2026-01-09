"""
Consolidation Engine for Multi-Division Financial Model
Handles account matching across divisions using exact matching and ChatGPT AI

VERSION: 1.0.0
DATE: 2025-12-16
"""

import os
import json
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Optional, Tuple
from datetime import datetime


@dataclass
class DivisionConfig:
    """Configuration for a single division"""
    name: str
    is_primary: bool = False
    pl_file_path: str = ""
    bs_file_path: str = ""
    pl_accounts: List[Dict] = field(default_factory=list)
    bs_accounts: List[Dict] = field(default_factory=list)


@dataclass
class AccountMapping:
    """Mapping of a consolidated account to division-specific accounts"""
    consolidated_name: str
    division_mappings: Dict[str, str] = field(default_factory=dict)  # division_name -> original_account_name
    match_type: str = "exact"  # 'exact', 'intelligent', 'manual', 'unmatched'
    confidence: float = 1.0
    approved: bool = False
    approved_date: Optional[str] = None
    reasoning: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> 'AccountMapping':
        return cls(**data)


class ConsolidationEngine:
    """Engine for matching accounts across divisions and consolidating"""

    def __init__(self, divisions: List[DivisionConfig], api_key: str = None):
        """
        Initialize the consolidation engine

        Args:
            divisions: List of DivisionConfig objects with parsed account data
            api_key: OpenAI API key (defaults to OPENAI_API_KEY environment variable)
        """
        self.divisions = divisions
        self.api_key = api_key or os.environ.get('OPENAI_API_KEY')
        self.mappings: Dict[str, AccountMapping] = {}
        self._openai_client = None

    def _get_openai_client(self):
        """Lazy initialization of OpenAI client"""
        if self._openai_client is None and self.api_key:
            try:
                import openai
                self._openai_client = openai.OpenAI(api_key=self.api_key)
            except ImportError:
                print("Warning: openai package not installed. Intelligent matching disabled.")
                return None
            except Exception as e:
                print(f"Warning: Could not initialize OpenAI client: {e}")
                return None
        return self._openai_client

    def match_accounts(self, statement_type: str = "pl") -> Dict[str, AccountMapping]:
        """
        Execute three-phase matching algorithm

        Args:
            statement_type: "pl" for P&L accounts, "bs" for Balance Sheet accounts

        Returns:
            Dictionary mapping consolidated account names to AccountMapping objects
        """
        # Get accounts from each division
        if statement_type == "pl":
            accounts_by_division = {d.name: d.pl_accounts for d in self.divisions}
        else:
            accounts_by_division = {d.name: d.bs_accounts for d in self.divisions}

        # Phase 1: Exact matches
        exact_matches = self._find_exact_matches(accounts_by_division)

        # Phase 2: Fuzzy/intelligent matches (if API key available)
        if self.api_key and self._get_openai_client():
            intelligent_matches = self._find_intelligent_matches(accounts_by_division, exact_matches)
        else:
            intelligent_matches = {}

        # Phase 3: Identify unmatched accounts
        unmatched = self._identify_unmatched(accounts_by_division, exact_matches, intelligent_matches)

        # Combine all mappings
        all_mappings = {**exact_matches, **intelligent_matches, **unmatched}
        self.mappings = all_mappings
        return all_mappings

    def _find_exact_matches(self, accounts_by_division: Dict[str, List[Dict]]) -> Dict[str, AccountMapping]:
        """
        Phase 1: Find accounts with identical names across ANY divisions

        Changed from original: Now matches accounts even if they only exist in some divisions.
        An account named "6145 Insurance" in Division A and Division B will be consolidated
        together even if Division C doesn't have it.

        Args:
            accounts_by_division: Dict mapping division name to list of account dicts

        Returns:
            Dictionary of exact match AccountMapping objects
        """
        exact_matches = {}

        # Collect ALL unique account names across all divisions (normalized)
        all_account_names = {}  # normalized_name -> {div_name: original_name}

        for div_name, accounts in accounts_by_division.items():
            for account in accounts:
                if account.get('is_header', False):
                    continue  # Skip headers

                name = account['name']
                name_normalized = self._normalize_account_name(name)

                if name_normalized not in all_account_names:
                    all_account_names[name_normalized] = {}
                all_account_names[name_normalized][div_name] = name

        # Create mappings for each unique account name
        for name_normalized, div_mappings in all_account_names.items():
            # Use the first division's name as the consolidated name (preserve original casing)
            consolidated_name = list(div_mappings.values())[0]

            # Determine match type based on how many divisions have this account
            if len(div_mappings) == len(self.divisions):
                match_type = 'exact'  # All divisions have it
            elif len(div_mappings) > 1:
                match_type = 'exact'  # Multiple divisions have it - still exact match
            else:
                match_type = 'division_specific'  # Only one division has it

            exact_matches[consolidated_name] = AccountMapping(
                consolidated_name=consolidated_name,
                division_mappings=div_mappings,
                match_type=match_type,
                confidence=1.0,
                approved=True  # Auto-approve exact matches
            )

        return exact_matches

    def _normalize_account_name(self, name: str) -> str:
        """Normalize account name for comparison"""
        return ' '.join(name.lower().strip().split())

    def _find_intelligent_matches(self,
                                   accounts_by_division: Dict[str, List[Dict]],
                                   existing: Dict[str, AccountMapping]
                                  ) -> Dict[str, AccountMapping]:
        """
        Phase 2: Use ChatGPT API for fuzzy matching of unmatched accounts

        Args:
            accounts_by_division: Dict mapping division name to list of account dicts
            existing: Already matched accounts to exclude

        Returns:
            Dictionary of intelligent match AccountMapping objects
        """
        # Get unmatched accounts by division
        unmatched_by_div = self._get_unmatched_accounts(accounts_by_division, existing)

        if not unmatched_by_div or all(len(v) == 0 for v in unmatched_by_div.values()):
            return {}

        # Build prompt for ChatGPT
        prompt = self._build_matching_prompt(unmatched_by_div)

        try:
            # Call OpenAI API
            response = self._call_openai_api(prompt)

            # Parse response into mappings
            return self._parse_matching_response(response, unmatched_by_div)
        except Exception as e:
            print(f"Warning: Intelligent matching failed: {e}")
            return {}

    def _get_unmatched_accounts(self,
                                 accounts_by_division: Dict[str, List[Dict]],
                                 existing: Dict[str, AccountMapping]) -> Dict[str, List[str]]:
        """Get accounts that haven't been matched yet"""
        unmatched = {}

        for div_name, accounts in accounts_by_division.items():
            matched_names = set()
            for mapping in existing.values():
                if div_name in mapping.division_mappings:
                    matched_names.add(mapping.division_mappings[div_name])

            unmatched[div_name] = [
                a['name'] for a in accounts
                if a['name'] not in matched_names
                and not a.get('is_header', False)  # Skip headers
            ]

        return unmatched

    def _build_matching_prompt(self, unmatched: Dict[str, List[str]]) -> str:
        """Build prompt for ChatGPT account matching"""
        prompt = """You are a financial accounting expert. Your task is to match chart of accounts
from different business divisions that represent the same economic activity but may have different names.

Analyze these accounts from different divisions and identify which ones should be consolidated together:

"""
        for div_name, accounts in unmatched.items():
            prompt += f"\n**{div_name}:**\n"
            # Limit accounts to prevent token overflow
            for acct in accounts[:50]:
                prompt += f"  - {acct}\n"

        prompt += """
Return a JSON object with suggested mappings. Only include matches where you are confident (>0.7)
that the accounts represent the same economic activity.

Focus on common variations like:
- Rent variations (Building Rent, Facility Lease, Office Rent, Rent Expense)
- Payroll variations (Salaries, Wages, Compensation, Payroll Expense)
- Marketing variations (Advertising, Marketing, Promotions, Marketing Expense)
- Utilities (Electric, Gas, Water as separate or combined)
- Insurance variations (Insurance Expense, Business Insurance, Liability Insurance)

Return format:
{
  "mappings": [
    {
      "consolidated_name": "Standard consolidated account name",
      "matches": {"Division A": "Original name A", "Division B": "Original name B"},
      "confidence": 0.85,
      "reasoning": "Brief explanation of why these match"
    }
  ]
}

Important rules:
1. Only suggest matches if confidence > 0.7
2. Use the most descriptive account name as the consolidated_name
3. Never match accounts from different categories (e.g., don't match Revenue with Expense)
4. If unsure, don't include the match - it's better to leave accounts separate
"""
        return prompt

    def _call_openai_api(self, prompt: str) -> dict:
        """Call OpenAI API for account matching"""
        client = self._get_openai_client()
        if not client:
            return {"mappings": []}

        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a financial accounting expert specializing in chart of accounts consolidation."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                response_format={"type": "json_object"}
            )

            return json.loads(response.choices[0].message.content)
        except Exception as e:
            print(f"OpenAI API error: {e}")
            return {"mappings": []}

    def _parse_matching_response(self,
                                  response: dict,
                                  unmatched: Dict[str, List[str]]) -> Dict[str, AccountMapping]:
        """Parse ChatGPT response into AccountMapping objects"""
        mappings = {}

        for match in response.get("mappings", []):
            cons_name = match.get("consolidated_name", "")
            div_matches = match.get("matches", {})
            confidence = match.get("confidence", 0.5)
            reasoning = match.get("reasoning", "")

            # Validate the match - ensure referenced accounts actually exist
            valid_mappings = {}
            for div_name, acct_name in div_matches.items():
                if div_name in unmatched and acct_name in unmatched[div_name]:
                    valid_mappings[div_name] = acct_name

            # Only create mapping if we have valid matches from at least 2 divisions
            if len(valid_mappings) >= 2 and confidence >= 0.7:
                mappings[cons_name] = AccountMapping(
                    consolidated_name=cons_name,
                    division_mappings=valid_mappings,
                    match_type='intelligent',
                    confidence=confidence,
                    approved=False,  # Require user approval for intelligent matches
                    reasoning=reasoning
                )

        return mappings

    def _identify_unmatched(self,
                            accounts_by_division: Dict[str, List[Dict]],
                            exact: Dict[str, AccountMapping],
                            intelligent: Dict[str, AccountMapping]) -> Dict[str, AccountMapping]:
        """
        Phase 3: Identify any remaining unmatched accounts

        With the new exact matching logic that matches accounts across ANY divisions
        (not requiring all divisions to have the account), this should rarely find
        anything. But we keep it as a safety net.

        Note: We NO LONGER add "(Division)" suffix since the consolidated view
        should just show the account name and sum across all divisions that have it.
        """
        all_existing = {**exact, **intelligent}
        unmatched_mappings = {}

        for div_name, accounts in accounts_by_division.items():
            # Get already matched account names for this division
            matched_names = set()
            for mapping in all_existing.values():
                if div_name in mapping.division_mappings:
                    matched_names.add(mapping.division_mappings[div_name])

            # Any remaining unmatched accounts (should be rare with new logic)
            for account in accounts:
                name = account['name']
                if name not in matched_names and not account.get('is_header', False):
                    # Use the account name WITHOUT division prefix for consolidated view
                    # The account will just show values from the divisions that have it
                    if name not in unmatched_mappings:
                        unmatched_mappings[name] = AccountMapping(
                            consolidated_name=name,
                            division_mappings={div_name: name},
                            match_type='division_specific',
                            confidence=1.0,
                            approved=True
                        )
                    else:
                        # Add this division to existing mapping
                        unmatched_mappings[name].division_mappings[div_name] = name

        return unmatched_mappings

    def consolidate_values(self,
                           mappings: Dict[str, AccountMapping],
                           months: List[Tuple[int, int, str]],
                           statement_type: str = "pl") -> List[Dict]:
        """
        Create consolidated account list with summed values

        Args:
            mappings: Account mappings from match_accounts()
            months: List of (month, year, display_name) tuples
            statement_type: "pl" or "bs"

        Returns:
            List of consolidated account dictionaries
        """
        consolidated = []

        for cons_name, mapping in mappings.items():
            # Determine account properties from source accounts
            is_total = False
            is_header = False
            indent = 0

            # Get accounts from each division
            for div in self.divisions:
                div_account_name = mapping.division_mappings.get(div.name)
                if div_account_name:
                    accounts = div.pl_accounts if statement_type == "pl" else div.bs_accounts
                    for acct in accounts:
                        if acct['name'] == div_account_name:
                            is_total = is_total or acct.get('is_total', False)
                            is_header = is_header or acct.get('is_header', False)
                            indent = max(indent, acct.get('indent', 0))
                            break

            cons_account = {
                'name': cons_name,
                'division': 'Consolidated',
                'is_total': is_total,
                'is_header': is_header,
                'indent': indent,
                'values': {},
                'source_mappings': mapping.division_mappings,
                'match_type': mapping.match_type
            }

            # Sum values across divisions for each month
            for m, y, name in months:
                total = 0
                for div in self.divisions:
                    div_account_name = mapping.division_mappings.get(div.name)
                    if div_account_name:
                        accounts = div.pl_accounts if statement_type == "pl" else div.bs_accounts
                        for acct in accounts:
                            if acct['name'] == div_account_name:
                                total += acct['values'].get((m, y), 0)
                                break
                cons_account['values'][(m, y)] = total

            consolidated.append(cons_account)

        # Sort accounts to maintain proper P&L/BS structure
        sorted_accounts = self._sort_accounts_by_section(consolidated, statement_type)
        return sorted_accounts

    def _sort_accounts_by_section(self, accounts: List[Dict], statement_type: str) -> List[Dict]:
        """Sort accounts to maintain proper financial statement structure.

        REWRITTEN: Uses strict P&L ordering to ensure proper structure:
        1. Income (Revenue) section
        2. Cost of Goods Sold section
        3. Gross Profit (summary line)
        4. Operating Expenses section
        5. Net Operating Income (summary line)
        6. Other Income section
        7. Other Expenses section
        8. Net Income (final line)

        Accounts are sorted by:
        1. Exact match to primary division order (highest priority)
        2. Account number prefix (4xxx=Income, 5xxx=COGS, 6xxx=Expenses, etc.)
        3. Keyword matching as fallback
        """
        import re

        if not accounts or not self.divisions:
            return accounts

        # Get the primary division or first division
        primary_div = None
        for div in self.divisions:
            if div.is_primary:
                primary_div = div
                break
        if not primary_div:
            primary_div = self.divisions[0]

        # Get the reference account order from primary division
        if statement_type == "pl":
            ref_accounts = primary_div.pl_accounts
        else:
            ref_accounts = primary_div.bs_accounts

        # Build comprehensive order map and identify section boundaries
        order_map = {}  # normalized_name -> (position, section)
        current_section = 1  # Start in Income section

        # P&L Section definitions:
        # 1 = Income, 2 = COGS, 3 = Expenses, 4 = Other Income, 5 = Other Expense, 6 = Net Income

        for idx, acct in enumerate(ref_accounts):
            name_lower = acct['name'].lower().strip()

            # Detect section transitions based on summary/total lines
            if statement_type == "pl":
                # After "Total for Income" or similar, move to COGS
                if current_section == 1 and ('total' in name_lower and 'income' in name_lower and 'net' not in name_lower and 'other' not in name_lower):
                    current_section = 2
                # After "Gross Profit", move to Expenses
                elif current_section <= 2 and 'gross profit' in name_lower:
                    current_section = 3
                # After "Total for Expenses" or "Net Operating Income", move to Other Income
                elif current_section == 3 and ('total' in name_lower and 'expense' in name_lower):
                    current_section = 4
                elif current_section == 3 and 'net operating income' in name_lower:
                    current_section = 4
                # After "Total for Other Income", move to Other Expense
                elif current_section == 4 and 'total' in name_lower and 'other income' in name_lower:
                    current_section = 5
                # "Net Income" is always section 6
                elif 'net income' in name_lower and 'operating' not in name_lower and 'other' not in name_lower:
                    order_map[name_lower] = (idx, 6)
                    continue

            order_map[name_lower] = (idx, current_section)

        def get_account_section(acct_name: str) -> int:
            """Determine P&L section for an account. Returns section number 1-6."""
            name = acct_name.strip()
            name_lower = name.lower()

            # Check if it's a known summary/total line - these have fixed positions
            if 'net income' in name_lower and 'operating' not in name_lower and 'other' not in name_lower:
                return 6  # Net Income is always last
            if 'gross profit' in name_lower:
                return 2  # Gross Profit ends COGS section
            if 'net operating income' in name_lower or 'operating income' in name_lower:
                return 3  # Net Operating Income ends Expenses section

            # First try: Account number prefix
            match = re.match(r'^(\d{3,5})', name)
            if match:
                acct_num = int(match.group(1))
                # Normalize to 4-digit format
                while acct_num < 1000:
                    acct_num *= 10
                while acct_num >= 10000:
                    acct_num //= 10

                if 4000 <= acct_num < 5000:
                    return 1  # Income (4xxx)
                elif 5000 <= acct_num < 6000:
                    return 2  # COGS (5xxx)
                elif 6000 <= acct_num < 7000:
                    return 3  # Expenses (6xxx)
                elif 7000 <= acct_num < 8000:
                    return 4  # Other Income (7xxx)
                elif 8000 <= acct_num < 9000:
                    return 5  # Other Expense (8xxx)

            # Second try: Keyword matching (more specific patterns first)

            # Check for "Total" lines - these stay in their section
            if 'total' in name_lower:
                if 'income' in name_lower and 'other' not in name_lower and 'net' not in name_lower:
                    return 1  # Total Income
                if 'cost' in name_lower or 'cogs' in name_lower or 'goods' in name_lower:
                    return 2  # Total COGS
                if 'expense' in name_lower and 'other' not in name_lower:
                    return 3  # Total Expenses
                if 'other income' in name_lower:
                    return 4  # Total Other Income
                if 'other expense' in name_lower:
                    return 5  # Total Other Expense

            # Other Income keywords (check before general income)
            if 'other income' in name_lower or 'miscellaneous income' in name_lower:
                return 4

            # Other Expense keywords (check before general expense)
            if 'other expense' in name_lower or 'loss on' in name_lower or 'write-off' in name_lower:
                return 5

            # COGS keywords (check before expense since some overlap)
            cogs_patterns = [
                'cost of goods', 'cost of sales', 'cost of service', 'cogs',
                'direct cost', 'direct labor', 'direct material',
                'purchases', 'freight in', 'manufacturing', 'production',
                'job cost', 'subcontract', 'contract labor'
            ]
            if any(p in name_lower for p in cogs_patterns):
                return 2

            # Income keywords (revenue)
            income_patterns = [
                'revenue', 'sales', 'income', 'fees earned', 'service fee',
                'consulting fee', 'commission', 'royalt', 'dividend received',
                'interest earned', 'rental income', 'gain on'
            ]
            # Must NOT be expense-related
            expense_related = ['expense', 'cost', 'loss']
            if any(p in name_lower for p in income_patterns):
                if not any(e in name_lower for e in expense_related):
                    return 1

            # Expense keywords (operating expenses)
            expense_patterns = [
                'expense', 'rent', 'lease', 'utilities', 'electric', 'gas bill', 'water',
                'telephone', 'phone', 'internet', 'insurance', 'depreciation', 'amortization',
                'payroll', 'salary', 'salaries', 'wages', 'compensation', 'benefits',
                'tax', 'taxes', 'license', 'permit', 'dues', 'subscription',
                'advertising', 'marketing', 'promotion', 'office supplies', 'postage',
                'shipping', 'delivery', 'travel', 'meal', 'entertainment',
                'auto', 'vehicle', 'fuel', 'mileage', 'repair', 'maintenance',
                'professional fee', 'legal', 'accounting', 'consulting', 'bank charge',
                'credit card fee', 'interest paid', 'training', 'education',
                'software', 'computer', 'equipment', 'security', 'uniform', 'tool'
            ]
            if any(p in name_lower for p in expense_patterns):
                return 3

            # If we found this account in the reference, use its section
            if name_lower in order_map:
                return order_map[name_lower][1]

            # Last resort: Search through all divisions to find position context
            for div in self.divisions:
                div_accounts = div.pl_accounts
                section = 1

                for acct in div_accounts:
                    acct_lower = acct['name'].lower().strip()

                    # Track section transitions
                    if 'total' in acct_lower and 'income' in acct_lower and 'net' not in acct_lower and 'other' not in acct_lower:
                        section = 2
                    elif 'gross profit' in acct_lower:
                        section = 3
                    elif section == 3 and ('total' in acct_lower and 'expense' in acct_lower):
                        section = 4
                    elif 'net operating' in acct_lower:
                        section = 4

                    if acct_lower == name_lower:
                        return section

            return 3  # Default to expenses if truly unknown

        def get_sort_key(acct):
            """Generate sort key ensuring proper P&L structure."""
            name = acct['name']
            name_lower = name.lower().strip()
            is_total = acct.get('is_total', False)
            is_header = acct.get('is_header', False)

            # If exact match in reference, use that position
            if name_lower in order_map:
                pos, section = order_map[name_lower]
                # Use section as primary sort, then position within section
                return (section, 0, pos, name)

            # Otherwise categorize and place after matching accounts in that section
            section = get_account_section(name)

            # Headers come first in section, totals come last in section
            if is_header:
                sub_order = 0
            elif is_total or 'total' in name_lower:
                sub_order = 2
            else:
                sub_order = 1

            # Use section, then sub_order, then alphabetical
            return (section, sub_order, 999, name)

        sorted_accounts = sorted(accounts, key=get_sort_key)
        return sorted_accounts

    def update_mapping(self,
                       consolidated_name: str,
                       division_mappings: Dict[str, str],
                       approved: bool = True) -> AccountMapping:
        """
        Update or create a manual mapping

        Args:
            consolidated_name: The consolidated account name
            division_mappings: Dict of division_name -> account_name
            approved: Whether to mark as approved

        Returns:
            The updated AccountMapping
        """
        if consolidated_name in self.mappings:
            mapping = self.mappings[consolidated_name]
            mapping.division_mappings.update(division_mappings)
            mapping.match_type = 'manual'
            mapping.approved = approved
            mapping.approved_date = datetime.now().isoformat()
        else:
            mapping = AccountMapping(
                consolidated_name=consolidated_name,
                division_mappings=division_mappings,
                match_type='manual',
                confidence=1.0,
                approved=approved,
                approved_date=datetime.now().isoformat()
            )
            self.mappings[consolidated_name] = mapping

        return mapping

    def approve_mapping(self, consolidated_name: str) -> bool:
        """Mark a mapping as approved"""
        if consolidated_name in self.mappings:
            self.mappings[consolidated_name].approved = True
            self.mappings[consolidated_name].approved_date = datetime.now().isoformat()
            return True
        return False

    def get_unapproved_mappings(self) -> List[AccountMapping]:
        """Get all mappings that need user approval"""
        return [m for m in self.mappings.values() if not m.approved]

    def get_all_mappings(self) -> Dict[str, AccountMapping]:
        """Get all mappings"""
        return self.mappings
