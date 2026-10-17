"""
file_matcher.py
---------------
Matches uploaded Test Case files with Test Report files.

Matching logic:
  - If exactly 1 Test Case is uploaded: pair it with every Test Report (one-to-many).
  - If multiple Test Cases are uploaded: fall back to exact filename (stem) matching.
"""
from dataclasses import dataclass
from pathlib import Path
from typing import Any, List, Tuple


@dataclass
class FilePair:
    name: str        # display key for this pair
    test_case: Any   # Streamlit UploadedFile
    test_report: Any # Streamlit UploadedFile


def match_files(
    tc_files: list,
    tr_files: list,
) -> Tuple[List[FilePair], List[str], List[str]]:
    """
    Match test-case files to test-report files.

    Returns
    -------
    matched      : list of FilePair
    unmatched_tc : TC filenames with no paired report (only relevant for multi-TC mode)
    unmatched_tr : TR filenames with no paired TC   (only relevant for multi-TC mode)
    """
    if len(tc_files) == 1:
        # One-to-many: single TC paired with every TR
        tc = tc_files[0]
        matched = [
            FilePair(name=Path(tr.name).stem, test_case=tc, test_report=tr)
            for tr in tr_files
        ]
        return matched, [], []

    # Multi-TC fallback: exact stem matching
    tc_map = {Path(f.name).stem: f for f in tc_files}
    tr_map = {Path(f.name).stem: f for f in tr_files}

    matched = [
        FilePair(name=stem, test_case=tc, test_report=tr_map[stem])
        for stem, tc in tc_map.items()
        if stem in tr_map
    ]
    unmatched_tc = [tc_map[s].name for s in tc_map if s not in tr_map]
    unmatched_tr = [tr_map[s].name for s in tr_map if s not in tc_map]

    return matched, unmatched_tc, unmatched_tr