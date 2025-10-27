"""
service_summary.py

Purpose
-------
Generates per-service summary text files that consolidate script outputs from
multiple Excel sheets. Each generated file aggregates all records and script
lines associated with a specific Service_ID and its assigned username.

Responsibilities
----------------
- Aggregate all related script lines across multiple sheets per Service_ID.
- Generate one well-structured summary text file per service.
- Ensure output directory creation and consistent UTF-8 encoding.
- Provide clear progress feedback via console logging.

Design Notes
------------
- Uses pathlib for OS-agnostic path handling.
- Each summary file name encodes both username and Service_ID for traceability.
- Output is structured for readability and downstream automation compatibility.
"""

from pathlib import Path


def generate_service_files(summary_table, service_usernames, script_contents, output_dir: str = "output") -> None:
    """
    Generate per-service summary files combining script data across multiple sheets.

    Consolidates all generated script contents for each Service_ID into a single,
    human-readable summary file. Each file contains the username header,
    service identifier, and script sections grouped by source sheet.

    Args:
        summary_table (list[tuple]):
            List of tuples in the format (sheet, service_id, username, count),
            summarizing data collected during Excel parsing and processing.
        service_usernames (dict):
            Mapping of {service_id: username}, associating each Service_ID with
            its destination username.
        script_contents (dict):
            Mapping of {(sheet, service_id): [lines]} where each value is a list
            of script lines generated for that particular sheet-service pair.
        output_dir (str, optional):
            Path to the output directory where summary files will be written.
            Defaults to "output".

    Returns:
        None

    Example:
        >>> generate_service_files(summary_table, service_usernames, script_contents)

    Raises:
        OSError: If any file cannot be written due to permission or I/O errors.
    """
    # -----------------------------------------------------------------------
    # Step 1: Prepare output directory
    # -----------------------------------------------------------------------
    # Ensure the target directory exists before attempting to write any files.
    # Using parents=True guarantees recursive creation of missing directories.
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    # -----------------------------------------------------------------------
    # Step 2: Group sheet names by Service_ID
    # -----------------------------------------------------------------------
    # Construct a dictionary structure like:
    # {
    #   "12345": ["Sheet1", "Sheet3"],
    #   "67890": ["Sheet2"]
    # }
    # Only include Service_IDs with a non-zero record count.
    service_dict = {}
    for sheet, sid, username, count in summary_table:
        if count == 0:
            continue  # Skip empty Service_IDs that produced no records
        service_dict.setdefault(sid, []).append(sheet)

    # -----------------------------------------------------------------------
    # Step 3: Generate a single summary file per Service_ID
    # -----------------------------------------------------------------------
    for sid, sheets in service_dict.items():
        # Retrieve username associated with this service
        username = service_usernames[sid]

        # Construct standardized output filename
        # Example: "john_1056_summary.txt"
        file_name = f"{username}_{sid}_summary.txt"
        file_path = output_path / file_name

        # -------------------------------------------------------------------
        # Write structured summary content to disk
        # -------------------------------------------------------------------
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                # -----------------------------------------------------------
                # Header Section
                # -----------------------------------------------------------
                # Each file starts with a clear, human-readable header block
                # that identifies the user and the corresponding Service_ID.
                f.write(f"{username.upper()}\n\n")
                f.write(f"{username:<12} service_{sid}\n\n")
                f.write("OA: \n\n")

                # -----------------------------------------------------------
                # Body Section: Aggregated script content per sheet
                # -----------------------------------------------------------
                for i, sheet in enumerate(sheets):
                    # Sanitize sheet name for readability (remove spaces)
                    safe_sheet = sheet.replace(" ", "_")

                    # Sheet section title
                    f.write(f"{safe_sheet}_script\n")

                    # Retrieve generated script lines for this sheet-service pair
                    content = script_contents.get((sheet, sid), [])
                    if content:
                        # Write script block with newline separation
                        f.write("\n".join(content) + "\n")

                    # Add separator line between sheets (except for the last one)
                    if i < len(sheets) - 1:
                        f.write("\n*****************************************\n\n")

            # Log or print feedback after successful file creation
            print(f"[INFO] Summary for Service_{sid} written to: {file_path}")

        except OSError as e:
            # Capture and report file write errors with full context
            print(f"[ERROR] Failed to write summary for Service_{sid}: {e}")
            raise
