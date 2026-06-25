"""
categoriser.py
--------------
Rule-based file categorisation engine for Railway File Optimiser.
Matches file names against keyword rules in categories.json.
No internet required — fully offline.
"""

import json
import os
import re
from pathlib import Path


class Categoriser:
    def __init__(self, categories_path: str = None):
        if categories_path is None:
            base = Path(__file__).parent
            categories_path = base / "categories.json"
        self.categories_path = str(categories_path)
        self.rules = {}
        self.default_category = "Miscellaneous"
        self._load_rules()

    def _load_rules(self):
        """Load categorisation rules from categories.json."""
        try:
            with open(self.categories_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.rules = data.get("categories", {})
            self.default_category = data.get("default_category", "Miscellaneous")
        except FileNotFoundError:
            print(f"[Categoriser] categories.json not found at {self.categories_path}. Using defaults.")
            self.rules = {}
        except json.JSONDecodeError as e:
            print(f"[Categoriser] Invalid categories.json: {e}. Using defaults.")
            self.rules = {}

    def reload_rules(self):
        """Reload rules from disk (useful after user edits categories.json)."""
        self._load_rules()

    def _normalise(self, text: str) -> str:
        """Normalise filename for matching: lowercase, replace separators."""
        text = text.lower()
        text = re.sub(r'[\s\-_\.]+', '_', text)
        return text

    def categorise(self, filepath: str) -> dict:
        """
        Analyse a file path and return categorisation results.

        Returns:
            dict with keys:
                - category (str): Best matching category name
                - confidence (str): 'high', 'medium', 'low'
                - matched_keyword (str): The keyword that triggered the match
                - matched_by (str): 'extension', 'keyword', 'default'
                - all_matches (list): All categories that matched
        """
        path = Path(filepath)
        filename = path.name
        stem = path.stem        # filename without extension
        extension = path.suffix.lower()

        normalised_stem = self._normalise(stem)
        normalised_full = self._normalise(filename)

        matches = []

        for category, rules in self.rules.items():
            keywords = [kw.lower() for kw in rules.get("keywords", [])]
            extensions = [ext.lower() for ext in rules.get("extensions", [])]

            # --- Extension match (high confidence for specific extensions) ---
            if extension and extension in extensions:
                # Only use extension match if there's no keyword match yet
                matches.append({
                    "category": category,
                    "confidence": "medium",
                    "matched_keyword": extension,
                    "matched_by": "extension",
                    "score": 1
                })

            # --- Keyword matching ---
            for keyword in keywords:
                norm_kw = self._normalise(keyword)
                # Prefix match (highest confidence) — keyword at start of filename
                if normalised_stem.startswith(norm_kw) or normalised_full.startswith(norm_kw):
                    matches.append({
                        "category": category,
                        "confidence": "high",
                        "matched_keyword": keyword,
                        "matched_by": "prefix",
                        "score": 3
                    })
                # Whole-word match (medium-high confidence)
                elif re.search(r'(?<![a-z0-9])' + re.escape(norm_kw) + r'(?![a-z0-9])', normalised_stem):
                    matches.append({
                        "category": category,
                        "confidence": "high",
                        "matched_keyword": keyword,
                        "matched_by": "keyword",
                        "score": 2
                    })
                # Substring match (lower confidence)
                elif norm_kw in normalised_full:
                    matches.append({
                        "category": category,
                        "confidence": "medium",
                        "matched_keyword": keyword,
                        "matched_by": "substring",
                        "score": 1
                    })

        if not matches:
            return {
                "category": self.default_category,
                "confidence": "low",
                "matched_keyword": None,
                "matched_by": "default",
                "all_matches": []
            }

        # Sort by score descending, pick best
        matches.sort(key=lambda x: x["score"], reverse=True)
        best = matches[0]

        # Deduplicate all_matches by category
        seen = set()
        unique_matches = []
        for m in matches:
            if m["category"] not in seen:
                seen.add(m["category"])
                unique_matches.append(m)

        return {
            "category": best["category"],
            "confidence": best["confidence"],
            "matched_keyword": best["matched_keyword"],
            "matched_by": best["matched_by"],
            "all_matches": unique_matches
        }

    def get_all_categories(self) -> list:
        """Return list of all category names including default."""
        cats = list(self.rules.keys())
        if self.default_category not in cats:
            cats.append(self.default_category)
        return sorted(cats)

    def get_category_color(self, category: str) -> str:
        """Return the display color for a category, or a default gray."""
        if category in self.rules:
            return self.rules[category].get("color", "#95A5A6")
        return "#95A5A6"


# ---------- Quick test ----------
if __name__ == "__main__":
    c = Categoriser()
    test_files = [
        "TDR_2024_0056_NIT_Mumbai_East_Division.pdf",
        "Circular_No_45_Safety_Instructions_2024.pdf",
        "Monthly_Progress_Report_March_2024.xlsx",
        "Drawing_Bridge_No_12_Track_Layout.dwg",
        "MoM_DRM_Meeting_June_2024.docx",
        "Budget_Estimate_2024_25_Division.pdf",
        "Staff_Transfer_Order_May_2024.pdf",
        "SomethingTotallyUnknown_xyz123.pdf",
        "Accident_Inquiry_Report_Derailment_2024.pdf",
        "Work_Order_Bridge_Construction.pdf",
    ]
    print(f"{'File Name':<55} {'Category':<25} {'Confidence':<10} {'By'}")
    print("-" * 110)
    for f in test_files:
        result = c.categorise(f)
        print(f"{f:<55} {result['category']:<25} {result['confidence']:<10} {result['matched_by']} ({result['matched_keyword']})")
