"""
utils/excel_handler.py

Module Purpose:
---------------
This module provides utilities for working with Excel-based service data.
It focuses on reading workbook sheets, searching for service IDs, cleaning extracted
codes, and generating formatted action lines for downstream automation.

Core Responsibilities:
- Parse Excel workbooks using pandas + openpyxl.
- Perform case-insensitive lookups of 'service_id' entries.
- Extract and sanitize sub-identifiers or numeric codes.
- Produce preformatted action lines for automation pipelines.

Design Notes:
-------------
- Logging is centralized using the project's custom logger.
- Emphasis is placed on robustness and traceability: every operation logs its success/failure.
- Built for extensibility — e.g., adding support for multiple service lookup fields or
  configurable formatting rules.
"""

from typing import List, Iterable
import pandas as pd
from utils.logger import setup_logger  # Project-wide logging utility

# Instantiate a module-specific logger for fine-grained logging context
logger = setup_logger("excel_handler")


class ExcelHandler:
    """Encapsulates Excel reading and service code extraction logic."""

    def __init__(self, excel_path: str):
        """
        Initialize an ExcelHandler instance and attempt to load the workbook.

        Args:
            excel_path (str): Path to the Excel file.

        Responsibilities:
            - Initialize a pandas.ExcelFile object (lazy-loading for performance).
            - Cache sheet names for future queries.
            - Log file load events and handle access errors gracefully.

        Raises:
            FileNotFoundError: Raised when the Excel file cannot be read.
        """
        self.path = excel_path  # Persist Excel file path for reference/logging

        try:
            # Using pandas.ExcelFile for deferred reading (memory-efficient with multiple sheets)
            self._excel = pd.ExcelFile(self.path, engine="openpyxl")
            logger.info("Successfully loaded Excel file: %s", self.path)
        except Exception as e:
            # Log the error stack trace and raise a descriptive exception
            logger.exception("Failed to open Excel file: %s", e)
            raise FileNotFoundError(f"Could not open Excel file: {e}")

    def sheet_names(self) -> List[str]:
        """
        Return all available worksheet names from the loaded Excel file.

        Returns:
            List[str]: List of sheet names present in the workbook.
        """
        # ExcelFile automatically stores all sheet names upon initialization
        return self._excel.sheet_names

    def find_service_codes(self, sheet_name: str, service_id: str) -> List[str]:
        """
        Search a specific sheet for entries matching a given service_id.

        Detailed Flow:
        --------------
        1. Parse the given sheet into a DataFrame.
        2. Normalize column names to lowercase.
        3. Identify the 'service_id' column dynamically (case-insensitive search).
        4. Match the target service ID, allowing for both '1056' and 'Service_1056' formats.
        5. Extract all corresponding 'sub-identifier' values.
        6. Deduplicate and return the clean list of codes.

        Args:
            sheet_name (str): Sheet to search within.
            service_id (str): Service ID to locate (numeric or prefixed).

        Returns:
            List[str]: A list of unique codes corresponding to the given service.
        """
        try:
            # Attempt to parse the target sheet; defer reading until now for memory efficiency
            df = self._excel.parse(sheet_name)
        except Exception as e:
            # Log but don't crash if one sheet is malformed or inaccessible
            logger.warning("Unable to read sheet '%s': %s", sheet_name, e)
            return []

        # Standardize all column names (critical for case-insensitive matching)
        df.columns = [str(c).lower().strip() for c in df.columns]

        # Locate the service_id column dynamically — handles variations like 'ServiceID', 'service_id', etc.
        svc_col = next((c for c in df.columns if "service" in c and "id" in c), None)
        if not svc_col:
            logger.warning("Sheet '%s' missing a 'service_id' column.", sheet_name)
            return []

        # Normalize provided service ID (ensure consistent 'service_' prefix)
        normalized_service = str(service_id).strip()
        if not normalized_service.lower().startswith("service_"):
            normalized_service = f"service_{normalized_service}"

        # Create a case-insensitive match mask for the target service
        mask = df[svc_col].astype(str).str.strip().str.lower() == normalized_service.lower()
        filtered = df.loc[mask]

        # Early exit if no matching entries were found
        if filtered.empty:
            logger.info("No matching entries for %s in sheet '%s'.", normalized_service, sheet_name)
            return []

        # Attempt to locate a sub-identifier column for code extraction
        sub_col = next((c for c in df.columns if "sub" in c and "identifier" in c), None)
        if not sub_col:
            logger.warning("No 'sub-identifier' column found in sheet '%s'.", sheet_name)
            return []

        # Extract non-null, non-empty sub-identifier values
        codes = [str(v).strip() for v in filtered[sub_col] if pd.notna(v) and str(v).strip()]

        # Deduplicate results while preserving order
        unique_codes = list(dict.fromkeys(codes))

        # Log summary of extraction results for auditing/debugging
        logger.info(
            "Extracted %d unique code(s) for %s from sheet '%s'.",
            len(unique_codes),
            normalized_service,
            sheet_name,
        )

        return unique_codes

    @staticmethod
    def clean_codes(codes: Iterable[str]) -> List[str]:
        """
        Clean and normalize raw extracted codes.

        Cleaning Rules:
        ---------------
        - Trim whitespace.
        - Remove '?' characters.
        - Retain '*' symbols (treated as valid wildcards).
        - Remove spaces and hyphens.
        - Skip empty or invalid results.

        Args:
            codes (Iterable[str]): Iterable of raw code strings.

        Returns:
            List[str]: List of sanitized, non-empty codes.
        """
        cleaned = []
        for raw in codes:
            s = str(raw).strip()
            if not s:
                continue  # Skip blank/whitespace-only entries

            # Perform character-level sanitization (retain only meaningful symbols)
            s = s.replace("?", "").replace(" ", "").replace("-", "")
            if s:
                cleaned.append(s)

        return cleaned

    @staticmethod
    def format_action_lines(codes: Iterable[str], username: str) -> List[str]:
        """
        Format cleaned service codes into standardized action strings.

        Example Output:
            { "?.?.27840001402" }  : Actions SET_DEST_LA("cellfsc"),SET_ESME_GROUP(SAG_GROUP_1, A_ADDR)

        Args:
            codes (Iterable[str]): List of cleaned codes to format.
            username (str): Username to embed within each formatted line.

        Returns:
            List[str]: List of fully formatted action strings for script generation or export.
        """
        lines: List[str] = []

        # Sanitize username input to prevent malformed output
        username = username.replace('"', "").replace("'", "")

        # Construct the formatted action line for each code
        for code in codes:
            # Maintain a consistent pattern used by the automation system
            line = (
                f'{{ "?.?.{code}" }}  : Actions '
                f'SET_DEST_LA("{username}"),SET_ESME_GROUP(SAG_GROUP_1, A_ADDR)'
            )
            lines.append(line)

        return lines
