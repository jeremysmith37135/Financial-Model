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
        Phase 1: Find accounts with identical names across all divisions

        Args:
            accounts_by_division: Dict mapping division name to list of account dicts

        Returns:
            Dictionary of exact match AccountMapping objects
        """
        # Get primary division
        primary_div = next((d for d in self.divisions if d.is_primary), self.divisions[0])
        primary_accounts = accounts_by_division.get(primary_div.name, [])
        other_div_names = [d.name for d in self.divisions if d.name != primary_div.name]

        exact_matches = {}

        for account in primary_accounts:
            name = account['name']
            name_normalized = self._normalize_account_name(name)

            # Check if this exact name exists in all other divisions
            matches_all = True
            div_mappings = {primary_div.name: name}

            for other_name in other_div_names:
                other_accounts = accounts_by_division.get(other_name, [])
                other_account_names = [a['name'] for a in other_accounts]
                other_normalized = [self._normalize_account_name(n) for n in other_account_names]

                # Check for exact match (case-insensitive, whitespace-normalized)
                if name_normalized in other_normalized:
                    idx = other_normalized.index(name_normalized)
                    div_mappings[other_name] = other_account_names[idx]
                elif name in other_account_names:
                    div_mappings[other_name] = name
                else:
                    matches_all = False
                    break

            if matches_all and len(div_mappings) == len(self.divisions):
                exact_matches[name] = AccountMapping(
                    consolidated_name=name,
                    division_mappings=div_mappings,
                    match_type='exact',
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
        Phase 3: Create separate entries for unmatched accounts

        These accounts will appear separately in the consolidated report,
        never discarded or suppressed.
        """
        all_existing = {**exact, **intelligent}
        unmatched_mappings = {}

        for div_name, accounts in accounts_by_division.items():
            # Get already matched account names for this division
            matched_names = set()
            for mapping in all_existing.values():
                if div_name in mapping.division_mappings:
                    matched_names.add(mapping.division_mappings[div_name])

            # Create individual mappings for unmatched accounts
            for account in accounts:
                name = account['name']
                if name not in matched_names and not account.get('is_header', False):
                    # Use division-prefixed name to avoid collisions
                    cons_name = f"{name} ({div_name})"
                    unmatched_mappings[cons_name] = AccountMapping(
                        consolidated_name=cons_name,
                        division_mappings={div_name: name},
                        match_type='unmatched',
                        confidence=1.0,
                        approved=True  # Auto-approve unmatched (they're just included separately)
                    )

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

        return consolidated

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
