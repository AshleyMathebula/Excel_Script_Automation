"""
main.py

Enhanced Version: Excel Service ID Script Generator with Service File Generation
-------------------------------------------------------------------------------

Purpose
-------
This script provides a CLI utility to process an Excel workbook containing
Service_ID records, generate per-Service_ID scripts, and create consolidated
summary files aggregating outputs from multiple sheets.

Key Features
------------
- User can select specific sheets or all sheets in the Excel workbook.
- Accepts one or more Service_IDs with corresponding usernames.
- Displays a preview table of found records per Service_ID before generation.
- Prompts user confirmation prior to writing any files.
- Generates both per-sheet script files and combined per-service summaries.

Dependencies
------------
- utils.excel_handler.ExcelHandler: For Excel file parsing and code extraction.
- utils.file_writer.FileWriter: For safe, UTF-8 text file writing.
- utils.logger.setup_logger: For consistent logging across the project.
- utils.service_summary.generate_service_files: To produce final summary files.
"""

from pathlib import Path
from typing import List

from utils.excel_handler import ExcelHandler
from utils.file_writer import FileWriter
from utils.logger import setup_logger
from utils.service_summary import generate_service_files

# ---------------------------------------------------------------------------
# Global Logger Configuration
# ---------------------------------------------------------------------------
# Centralized logger instance to capture events throughout this script.
logger = setup_logger("main")


# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------

def display_menu(options: List[str]) -> None:
    """
    Displays a numbered menu list for user selection.

    Args:
        options (List[str]): List of options (e.g., sheet names) to display.
    """
    for i, opt in enumerate(options, start=1):
        print(f"{i}. {opt}")
    print()  # Add spacing after menu


def parse_indices(raw: str, max_index: int) -> List[int]:
    """
    Parse user input string for selecting sheet indices.

    Supports:
      - 'all' keyword to select all sheets
      - Comma-separated list of numbers like '1,3,5'

    Args:
        raw (str): User input string.
        max_index (int): Total number of available sheets.

    Returns:
        List[int]: Zero-based indices corresponding to selected sheets.

    Raises:
        ValueError: Input contains invalid characters.
        IndexError: Selected index is out of range.
    """
    raw = raw.strip().lower()
    if raw == "all":
        return list(range(max_index))

    indices: List[int] = []
    for part in [p.strip() for p in raw.split(",") if p.strip()]:
        if not part.isdigit():
            raise ValueError("Sheet selections must be numbers or 'all'.")
        idx = int(part) - 1
        if idx < 0 or idx >= max_index:
            raise IndexError("Sheet index out of range.")
        indices.append(idx)

    return indices


def input_service_ids() -> List[str]:
    """
    Prompt the user to enter one or more Service_IDs.

    Returns:
        List[str]: List of cleaned Service_ID strings.
    """
    raw = input("Enter one or more Service_IDs (comma-separated): ").strip()
    return [s.strip() for s in raw.split(",") if s.strip()]


# ---------------------------------------------------------------------------
# Main Application Logic
# ---------------------------------------------------------------------------

def main() -> None:
    """
    Entry point for the Excel Service ID Script Generator.

    Workflow:
    1. Prompt user for Excel file path.
    2. Initialize ExcelHandler and FileWriter utilities.
    3. List and allow selection of sheets.
    4. Prompt for Service_IDs and corresponding usernames.
    5. Preview the number of matching records per sheet.
    6. Ask for user confirmation before generating scripts.
    7. Generate per-sheet scripts and track script content.
    8. Generate consolidated per-service summary files.
    9. Log results and provide console feedback.

    Error Handling:
        All unexpected exceptions are logged with traceback and reported to the user.
    """
    logger.info("Excel Service ID Script Generator started")

    print("=" * 80)
    print("Excel Service ID Script Generator (Enhanced Version)")
    print("=" * 80)

    # Step 1: Get Excel file path from user or use default
    excel_path = input(
        "Enter Excel file path (press Enter to use data/MO_Connection Database.xlsx): "
    ).strip() or "data/MO_Connection Database.xlsx"

    excel_file = Path(excel_path)
    if not excel_file.exists():
        logger.error("Excel file not found: %s", excel_file)
        print(f"Excel file not found: {excel_file}")
        return

    try:
        # Step 2: Initialize Excel handler and file writer utilities
        handler = ExcelHandler(str(excel_file))
        writer = FileWriter(output_dir="output")

        # Step 3: Retrieve available sheet names
        sheets = handler.sheet_names()
        if not sheets:
            logger.error("No sheets found in workbook: %s", excel_file)
            print(" No sheets found in the workbook.")
            return

        print("\nAvailable sheets:")
        display_menu(sheets)

        # Step 4: Allow user to select one or more sheets
        while True:
            selection = input("Select sheets by number (e.g., 1,3) or 'all': ").strip()
            try:
                selected_indices = parse_indices(selection, len(sheets))
                selected_sheets = [sheets[i] for i in selected_indices]
                break
            except (ValueError, IndexError) as e:
                print(f" {e} Please try again.")

        print(f"\nSelected sheets: {', '.join(selected_sheets)}")
        logger.info("Selected sheets: %s", selected_sheets)

        # Step 5: Collect Service_IDs from user
        service_ids = input_service_ids()
        if not service_ids:
            print(" No Service_IDs provided. Exiting.")
            return

        # Step 6: Prompt for username for each Service_ID
        service_usernames = {}
        for sid in service_ids:
            username = input(f"Enter username (destination) for Service_ID {sid}: ").strip()
            while not username:
                username = input("Username cannot be empty. Enter username: ").strip()
            service_usernames[sid] = username

        # Step 7: Preview number of matching codes per sheet
        print("\n Scanning sheets for matching records...\n")
        summary_table = []  # Stores tuples (sheet, service_id, username, count)

        for sheet in selected_sheets:
            for sid, username in service_usernames.items():
                codes = handler.find_service_codes(sheet, sid)
                count = len(codes)
                summary_table.append((sheet, sid, username, count))
                print(f"  • Sheet: {sheet:<25} | Service_{sid:<6} | User: {username:<15} | Codes found: {count}")

        # Display clean summary table
        print("\n" + "=" * 80)
        print("Summary of results:")
        print("=" * 80)
        for sheet, sid, username, count in summary_table:
            print(f"Sheet: {sheet:<25} | Service_{sid:<6} | User: {username:<15} | Total: {count}")
        print("=" * 80)

        # Step 8: Confirm with user before generating scripts
        confirm = input("\nProceed to generate scripts for these records? (y/n): ").strip().lower()
        if confirm != "y":
            print("Operation cancelled by user. No scripts generated.")
            logger.info("Operation cancelled after preview.")
            return

        # Step 9: Generate per-sheet scripts and track content
        script_contents = {}  # Maps (sheet, sid) → list of script lines
        total_processed = 0

        for sheet, sid, username, count in summary_table:
            if count == 0:
                continue  # Skip empty results

            # Extract, clean, and format codes
            codes = handler.find_service_codes(sheet, sid)
            cleaned = handler.clean_codes(codes)
            action_lines = handler.format_action_lines(cleaned, username)

            # Prepare safe output filename and write to disk
            safe_sheet = sheet.replace(" ", "_")
            out_name = f"{safe_sheet}_{sid}_script.txt"
            out_path = Path("output") / out_name
            writer.write_text_file(out_path, action_lines)

            # Save script content for summary generation
            script_contents[(sheet, sid)] = action_lines
            logger.info("Wrote %d lines to %s", len(action_lines), out_path)
            print(f" Generated {len(action_lines)} lines for {sheet} → {out_name}")
            total_processed += len(action_lines)

        # Step 10: Generate consolidated per-service summary files
        print("\n Generating consolidated service summary files...\n")
        generate_service_files(summary_table, service_usernames, script_contents, output_dir="output")
        logger.info("Service summary files generated successfully.")

        # Step 11: Completion message
        if total_processed == 0:
            print("\n No matching records were processed.")
        else:
            print(f"\n Done! Processed {total_processed} total records across selected sheets.")
            logger.info("Completed. Total processed: %d", total_processed)

    except Exception as exc:
        # Capture unexpected runtime errors
        logger.exception("Unhandled exception in main: %s", exc)
        print("An unexpected error occurred. See logs/activity.log for details.")


# ---------------------------------------------------------------------------
# Script Entry Point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
