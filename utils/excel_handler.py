from typing import List, Iterable
import pandas as pd
from utils.logger import setup_logger  # Project-wide logging utility

# Create a module-specific logger for Excel handling operations
logger = setup_logger("excel_handler")


class ExcelHandler:
    """
    Encapsulates Excel reading and service code extraction logic.

    Responsibilities:
    -----------------
    - Load Excel files efficiently using pandas + openpyxl.
    - Retrieve sheet names.
    - Extract codes for specific Service_IDs.
    - Extract all codes from a sheet for premium generation.
    - Clean and format codes into automation-ready action lines.
    """

    def __init__(self, excel_path: str):
        """
        Initialize an ExcelHandler instance and load the workbook.

        Args:
            excel_path (str): Path to the Excel file.

        Behavior:
        ---------
        - Attempts to load the workbook using pandas.ExcelFile.
        - Logs successful load or raises FileNotFoundError on failure.
        """
        self.path = excel_path
        try:
            # Lazy-load Excel file for memory efficiency
            self._excel = pd.ExcelFile(self.path, engine="openpyxl")
            logger.info("Successfully loaded Excel file: %s", self.path)
        except Exception as e:
            logger.exception("Failed to open Excel file: %s", e)
            raise FileNotFoundError(f"Could not open Excel file: {e}")

    def sheet_names(self) -> List[str]:
        """
        Retrieve all sheet names from the loaded Excel file.

        Returns:
            List[str]: Names of all worksheets in the workbook.
        """
        return self._excel.sheet_names

    def find_service_codes(self, sheet_name: str, service_id: str) -> List[str]:
        """
        Extract codes corresponding to a specific Service_ID in a sheet.

        Args:
            sheet_name (str): Sheet to search within.
            service_id (str): Target service ID (numeric or prefixed).

        Returns:
            List[str]: Unique codes for the service, or empty list if none found.

        Behavior:
        ---------
        - Parses the sheet into a DataFrame.
        - Dynamically locates 'service_id' column.
        - Matches the normalized service_id (ensures 'service_' prefix).
        - Extracts codes from the 'sub-identifier' column.
        - Deduplicates codes while preserving order.
        - Logs results for auditing.
        """
        try:
            df = self._excel.parse(sheet_name)
        except Exception as e:
            logger.warning("Unable to read sheet '%s': %s", sheet_name, e)
            return []

        # Standardize column names for case-insensitive matching
        df.columns = [str(c).lower().strip() for c in df.columns]

        # Find the column containing the Service_ID
        svc_col = next((c for c in df.columns if "service" in c and "id" in c), None)
        if not svc_col:
            logger.warning("Sheet '%s' missing a 'service_id' column.", sheet_name)
            return []

        # Normalize Service_ID (add 'service_' prefix if missing)
        normalized_service = str(service_id).strip()
        if not normalized_service.lower().startswith("service_"):
            normalized_service = f"service_{normalized_service}"

        # Create boolean mask for matching rows
        mask = df[svc_col].astype(str).str.strip().str.lower() == normalized_service.lower()
        filtered = df.loc[mask]

        if filtered.empty:
            logger.info("No matching entries for %s in sheet '%s'.", normalized_service, sheet_name)
            return []

        # Find the sub-identifier column to extract codes
        sub_col = next((c for c in df.columns if "sub" in c and "identifier" in c), None)
        if not sub_col:
            logger.warning("No 'sub-identifier' column found in sheet '%s'.", sheet_name)
            return []

        # Extract non-null, stripped values and remove duplicates
        codes = [str(v).strip() for v in filtered[sub_col] if pd.notna(v) and str(v).strip()]
        unique_codes = list(dict.fromkeys(codes))

        logger.info(
            "Extracted %d unique code(s) for %s from sheet '%s'.",
            len(unique_codes),
            normalized_service,
            sheet_name,
        )
        return unique_codes

    def extract_all_codes(self, sheet_name: str) -> List[str]:
        """
        Extract all codes from a sheet, independent of Service_ID.

        Intended Use:
        -------------
        - When script generation is skipped.
        - Generating premium short/long code files.
        - Compiling a comprehensive list of all sub-identifiers.

        Args:
            sheet_name (str): Worksheet to extract codes from.

        Returns:
            List[str]: Unique codes extracted from the sheet.

        Behavior:
        ---------
        - Loads the sheet into a DataFrame.
        - Dynamically finds the 'sub-identifier' column.
        - Extracts all non-null, non-empty values.
        - Deduplicates results while preserving order.
        - Logs the total number of extracted codes.
        """
        try:
            df = self._excel.parse(sheet_name)
        except Exception as e:
            logger.warning("Unable to read sheet '%s' for code extraction: %s", sheet_name, e)
            return []

        df.columns = [str(c).lower().strip() for c in df.columns]

        sub_col = next((c for c in df.columns if "sub" in c and "identifier" in c), None)
        if not sub_col:
            logger.warning("No 'sub-identifier' column found in sheet '%s'.", sheet_name)
            return []

        codes = [str(v).strip() for v in df[sub_col] if pd.notna(v) and str(v).strip()]
        unique_codes = list(dict.fromkeys(codes))

        logger.info("Extracted %d total code(s) from sheet '%s'.", len(unique_codes), sheet_name)
        return unique_codes

    @staticmethod
    def clean_codes(codes: Iterable[str]) -> List[str]:
        """
        Clean and normalize raw codes for script usage.

        Rules:
        ------
        - Remove '?', spaces, and hyphens.
        - Keep non-empty, meaningful codes.
        - Trim whitespace.

        Args:
            codes (Iterable[str]): Raw code strings.

        Returns:
            List[str]: Cleaned codes.
        """
        cleaned = []
        for raw in codes:
            s = str(raw).strip()
            if not s:
                continue
            s = s.replace("?", "").replace(" ", "").replace("-", "")
            if s:
                cleaned.append(s)
        return cleaned

    @staticmethod
    def format_action_lines(codes: Iterable[str], username: str) -> List[str]:
        """
        Format codes into automation-ready action strings.

        Example format:
        { "?.?.<code>" }  : Actions SET_DEST_LA("<username>"),SET_ESME_GROUP(SAG_GROUP_1, A_ADDR)

        Args:
            codes (Iterable[str]): Cleaned codes to format.
            username (str): Destination username to embed in each line.

        Returns:
            List[str]: List of formatted action lines for script generation.
        """
        lines: List[str] = []
        username = username.replace('"', "").replace("'", "")
        for code in codes:
            line = (
                f'{{ "?.?.{code}" }}  : Actions '
                f'SET_DEST_LA("{username}"),SET_ESME_GROUP(SAG_GROUP_1, A_ADDR)'
            )
            lines.append(line)
        return lines
