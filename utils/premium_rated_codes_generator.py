from pathlib import Path
from typing import Dict, List
import re

from utils.logger import setup_logger
from utils.file_writer import FileWriter

logger = setup_logger("premium_rated_codes_generator")


def generate_premium_code_files(
    script_contents: Dict[tuple, List[str]],
    writer: FileWriter,
    output_dir: str = "output"
) -> None:
    """
    Generate premium short and long code files from existing script contents
    separately for each Service_ID.
    """
    logger.info("Starting premium rated codes file generation...")

    # Organize codes per Service_ID
    service_codes: Dict[str, List[str]] = {}

    code_pattern = re.compile(r"\?\.\?\.(\d+)")
    for (sheet_name, service_id), lines in script_contents.items():
        for line in lines:
            match = code_pattern.search(line)
            if match:
                service_codes.setdefault(service_id, []).append(match.group(1))

    if not service_codes:
        logger.warning("No codes found in script contents. Premium files not generated.")
        print("No codes found. Skipping premium file generation.")
        return

    # Process each service_id individually
    for sid, codes in service_codes.items():
        # Deduplicate while preserving order
        seen = set()
        unique_codes = [c for c in codes if not (c in seen or seen.add(c))]

        # Classify by length
        short_codes = [c for c in unique_codes if len(c) <= 6]
        long_codes = [c for c in unique_codes if len(c) > 6]

        # Format output lines
        formatted_short = [f"?,?,{c}" for c in short_codes]
        formatted_long = [f"?,?,{c}" for c in long_codes]

        # Output file paths
        short_path = Path(output_dir) / f"premium_short_codes_Service_{sid}.txt"
        long_path = Path(output_dir) / f"premium_long_codes_Service_{sid}.txt"

        # Write files
        writer.write_text_file(short_path, formatted_short)
        writer.write_text_file(long_path, formatted_long)

        # Log summary
        logger.info(
            "Service %s: %d short codes, %d long codes written.",
            sid,
            len(short_codes),
            len(long_codes)
        )
        print(f"\nService {sid} Premium Codes Generated:")
        print(f"  → Short Codes : {len(short_codes)} written to {short_path}")
        print(f"  → Long Codes  : {len(long_codes)} written to {long_path}")

    logger.info("Premium code generation complete.")
