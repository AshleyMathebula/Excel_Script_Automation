from pathlib import Path
from typing import List

from utils.excel_handler import ExcelHandler
from utils.file_writer import FileWriter
from utils.logger import setup_logger
from utils.service_summary import generate_service_files
from utils.premium_rated_codes_generator import generate_premium_code_files

# ---------------------------------------------------------------------------
# Global Logger Configuration
# ---------------------------------------------------------------------------
logger = setup_logger("main")


# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------

def display_menu(options: List[str]) -> None:
    """Display numbered menu list for user selection."""
    for i, opt in enumerate(options, start=1):
        print(f"{i}. {opt}")
    print()


def parse_indices(raw: str, max_index: int) -> List[int]:
    """Parse user input for selecting sheet indices."""
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
    """Prompt user to enter one or more Service_IDs."""
    raw = input(
        "Enter one or more Service_IDs (comma-separated, leave blank to skip script generation): "
    ).strip()
    return [s.strip() for s in raw.split(",") if s.strip()]


# ---------------------------------------------------------------------------
# Main Application Logic
# ---------------------------------------------------------------------------

def main() -> None:
    global summary_table
    logger.info("Excel Service ID Script Generator started")
    print("=" * 80)
    print("Excel Service ID Script Generator (Enhanced Version)")
    print("=" * 80)

    # Step 1: Get Excel file path
    excel_path = input(
        "Enter Excel file path (press Enter to use data/MO_Connection Database.xlsx): "
    ).strip() or "data/MO_Connection Database.xlsx"

    excel_file = Path(excel_path)
    if not excel_file.exists():
        logger.error("Excel file not found: %s", excel_file)
        print(f"Excel file not found: {excel_file}")
        return

    try:
        # Step 2: Initialize utilities
        handler = ExcelHandler(str(excel_file))
        writer = FileWriter(output_dir="output")

        # Step 3: Retrieve available sheet names
        sheets = handler.sheet_names()
        if not sheets:
            logger.error("No sheets found in workbook: %s", excel_file)
            print("No sheets found in the workbook.")
            return

        print("\nAvailable sheets:")
        display_menu(sheets)

        # Step 4: Sheet selection
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

        # Step 5: Collect Service_IDs (optional)
        service_ids = input_service_ids()
        service_usernames = {}
        generate_scripts = bool(service_ids)

        if generate_scripts:
            for sid in service_ids:
                username = input(f"Enter username (destination) for Service_ID {sid}: ").strip()
                while not username:
                    username = input("Username cannot be empty. Enter username: ").strip()
                service_usernames[sid] = username

            # Step 6: Preview results
            print("\nScanning sheets for matching records...\n")
            summary_table = []

            for sheet in selected_sheets:
                for sid, username in service_usernames.items():
                    codes = handler.find_service_codes(sheet, sid)
                    count = len(codes)
                    summary_table.append((sheet, sid, username, count))
                    print(f"  • Sheet: {sheet:<25} | Service_{sid:<6} | User: {username:<15} | Codes found: {count}")

            print("\n" + "=" * 80)
            print("Summary of results:")
            print("=" * 80)
            for sheet, sid, username, count in summary_table:
                print(f"Sheet: {sheet:<25} | Service_{sid:<6} | User: {username:<15} | Total: {count}")
            print("=" * 80)

            # Step 7: Confirm before generating scripts
            confirm = input("\nProceed to generate scripts for these records? (y/n): ").strip().lower()
            if confirm != "y":
                print("Script generation skipped by user.")
                generate_scripts = False
                summary_table = []
        else:
            print("\nNo Service_IDs provided. Script generation will be skipped.")
            logger.info("User chose to skip script generation.")

        # Step 8: Generate per-sheet scripts if allowed
        script_contents = {}
        total_processed = 0

        if generate_scripts:
            for sheet, sid, username, count in summary_table:
                if count == 0:
                    continue

                codes = handler.find_service_codes(sheet, sid)
                cleaned = handler.clean_codes(codes)
                action_lines = handler.format_action_lines(cleaned, username)

                safe_sheet = sheet.replace(" ", "_")
                out_name = f"{safe_sheet}_{sid}_script.txt"
                out_path = Path("output") / out_name
                writer.write_text_file(out_path, action_lines)

                script_contents[(sheet, sid)] = action_lines
                logger.info("Wrote %d lines to %s", len(action_lines), out_path)
                print(f" Generated {len(action_lines)} lines for {sheet} → {out_name}")
                total_processed += len(action_lines)

            # Step 9: Generate consolidated summary files
            print("\nGenerating consolidated service summary files...\n")
            generate_service_files(summary_table, service_usernames, script_contents, output_dir="output")
            logger.info("Service summary files generated successfully.")

        if not generate_scripts:
            print("\nSkipping service summary generation since no scripts were created.")

        # -------------------------------------------------------------------
        # Step 10: Optional Premium Code Generation (per selected Service_IDs)
        # -------------------------------------------------------------------
        try:
            confirm_premium = input(
                "\nWould you like to generate Premium Short/Long Code files for the selected Service_IDs? (y/n): "
            ).strip().lower()

            if confirm_premium == "y" and service_ids:
                premium_script_contents = {}

                # Generate premium codes only for selected Service_IDs
                for sheet in selected_sheets:
                    for sid in service_ids:
                        codes = handler.find_service_codes(sheet, sid)
                        if not codes:
                            continue
                        cleaned = handler.clean_codes(codes)
                        username = service_usernames.get(sid, "PREMIUM")
                        action_lines = handler.format_action_lines(cleaned, username)
                        premium_script_contents[(sheet, sid)] = action_lines

                if premium_script_contents:
                    print("\nGenerating Premium Code files, please wait...\n")
                    generate_premium_code_files(premium_script_contents, writer, output_dir="output")
                else:
                    print("\nNo matching codes found for the selected Service_IDs. Premium files skipped.")
            else:
                print("\nPremium Code file generation skipped.\n")

        except Exception as e:
            logger.error(f"Error during premium code generation: {e}")
            print(f"An error occurred while generating premium files: {e}")

    except Exception as exc:
        logger.exception("Unhandled exception in main: %s", exc)
        print("An unexpected error occurred. See logs/activity.log for details.")


# ---------------------------------------------------------------------------
# Script Entry Point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
