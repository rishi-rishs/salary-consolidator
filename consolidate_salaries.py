import pandas as pd
import sys
import os
from rapidfuzz import process, fuzz
from datetime import datetime

def parse_date_from_sheet_name(sheet_name):
    """
    Attempts to parse a sheet name into a 'MMM-YY' string.
    Tries various formats. Returns the formatted string or the original name if parsing fails.
    """
    formats = [
        "%b %Y", "%B %Y", "%b-%Y", "%B-%Y", "%b_%Y", "%B_%Y",
        "%b%Y", "%m-%Y", "%Y-%m", "%d-%m-%Y"
    ]

    clean_name = sheet_name.strip()

    for fmt in formats:
        try:
            dt = datetime.strptime(clean_name, fmt)
            return dt.strftime("%b-%y")
        except ValueError:
            continue

    return clean_name

def get_standardized_name(name, current_salary, prev_month_data, all_standard_names, aliases):
    """
    Determines the standardized name by comparing to previous month's employees.

    Args:
        name: The name from the current sheet.
        current_salary: The salary of this employee in current month.
        prev_month_data: Dict of {name: salary} from the previous month.
        all_standard_names: Set of all standardized names seen so far.
        aliases: Dictionary mapping alternative names to standardized names.

    Returns:
        The standardized name to use, or None if it's a new entry.
    """
    name = str(name).strip()

    # 1. Check if it's already a known exact match
    if name in all_standard_names:
        return name

    # 2. Check if it's a known alias
    if name in aliases:
        return aliases[name]

    # 3. If no previous month data, it's the first month - all are new
    if not prev_month_data:
        return None

    # 4. Fuzzy match against previous month's employees only
    prev_month_names = list(prev_month_data.keys())
    matches = process.extract(name, prev_month_names, scorer=fuzz.WRatio, limit=2)

    if not matches:
        return None

    best_match, best_score, _ = matches[0]

    # Only prompt if score is above threshold
    if best_score < 80:
        return None

    # Get second best match if available
    second_match = None
    second_score = None
    if len(matches) > 1:
        second_match, second_score, _ = matches[1]

    print(f"\n{'='*60}")
    print(f"Current Employee: '{name}'")
    print(f"Current Salary:   {current_salary}")
    print(f"{'='*60}")
    print(f"Best match:       '{best_match}' (Score: {best_score:.1f})")
    print(f"  Previous Salary: {prev_month_data[best_match]}")

    if second_match and second_score >= 60:
        print(f"Second best:      '{second_match}' (Score: {second_score:.1f})")
        print(f"  Previous Salary: {prev_month_data[second_match]}")

    print(f"{'='*60}")

    while True:
        if second_match and second_score >= 60:
            choice = input(f"Match with: [1] {best_match} / [2] {second_match} / [n]ew / [t]ype manual: ").lower().strip()
        else:
            choice = input(f"Match with: [y]es {best_match} / [n]ew / [t]ype manual: ").lower().strip()

        if choice == 'y' or choice == 'yes' or choice == '1':
            aliases[name] = best_match
            return best_match
        elif choice == '2' and second_match and second_score >= 60:
            aliases[name] = second_match
            return second_match
        elif choice == 'n' or choice == 'new':
            return None
        elif choice == 't' or choice == 'type':
            manual_name = input("Enter the correct existing target name: ").strip()
            if manual_name in all_standard_names:
                aliases[name] = manual_name
                return manual_name
            else:
                print(f"'{manual_name}' not found in existing list. Creating as new entry? (y/n)")
                confirm = input("> ").lower().strip()
                if confirm == 'y':
                    return manual_name
                else:
                    continue
        else:
            print("Invalid choice. Please try again.")

def main():
    if len(sys.argv) < 2:
        print("Usage: python consolidate_salaries.py <path_to_excel_file>")
        sys.exit(1)

    file_path = sys.argv[1]

    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' not found.")
        sys.exit(1)

    print(f"Loading '{file_path}'... this might take a moment.")

    try:
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    # Master data structure: { "Standard Name": { "Month1": Salary, "Month2": Salary } }
    master_data = {}

    # All standardized names seen so far
    all_standard_names = set()

    # Aliases to remember user decisions: { "Ms. Deepa": "Deepa Thukkaraman" }
    aliases = {}

    # Previous month's data for comparison: { "Name": salary }
    prev_month_data = {}

    # Iterate through sheets
    for sheet_name in xls.sheet_names:
        print(f"\n{'#'*60}")
        print(f"### Processing Sheet: {sheet_name} ###")
        print(f"{'#'*60}")

        column_header = parse_date_from_sheet_name(sheet_name)

        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
        except Exception as e:
            print(f"Skipping sheet '{sheet_name}' due to read error: {e}")
            continue

        if df.shape[1] < 3:
            print(f"Skipping sheet '{sheet_name}': Expected at least 3 columns (A, B, C), found {df.shape[1]}")
            continue

        # Skip header row (first row)
        df = df.iloc[1:]

        # Skip last row if it contains "total" (case-insensitive)
        if len(df) > 0:
            last_name = str(df.iloc[-1, 1]).strip().lower() if not pd.isna(df.iloc[-1, 1]) else ""
            if "total" in last_name:
                df = df.iloc[:-1]

        # Current month's data (to become prev_month_data for next iteration)
        current_month_data = {}

        for index, row in df.iterrows():
            raw_name = row.iloc[1]  # Column B for employee name
            salary = row.iloc[2]    # Column C for salary

            if pd.isna(raw_name):
                continue

            raw_name = str(raw_name).strip()

            # Determine the standard name by comparing to previous month
            target_name = get_standardized_name(
                raw_name,
                salary,
                prev_month_data,
                all_standard_names,
                aliases
            )

            if target_name is None:
                # New employee
                target_name = raw_name
                all_standard_names.add(target_name)
                print(f"Added new employee: '{target_name}'")
            elif target_name not in all_standard_names:
                # Manually typed new name
                all_standard_names.add(target_name)
                print(f"Added new employee (manually typed): '{target_name}'")

            # Update Master Data
            if target_name not in master_data:
                master_data[target_name] = {}

            # Warn if overwriting
            if column_header in master_data[target_name]:
                print(f"Warning: Duplicate entry for '{target_name}' in '{column_header}'. Overwriting.")

            master_data[target_name][column_header] = salary

            # Track this employee for next month's comparison
            current_month_data[target_name] = salary

        # Update prev_month_data for next iteration
        prev_month_data = current_month_data.copy()

    # Create final DataFrame
    print(f"\n{'#'*60}")
    print("### Finalizing Data ###")
    print(f"{'#'*60}")

    final_df = pd.DataFrame.from_dict(master_data, orient='index')

    # Sort columns chronologically
    try:
        final_df = final_df.reindex(sorted(final_df.columns, key=lambda x: datetime.strptime(x, "%b-%y")), axis=1)
    except ValueError:
        print("Could not sort columns chronologically (mixed formats). Sorting alphabetically.")
        final_df = final_df.sort_index(axis=1)

    final_df.index.name = "Employee Name"

    output_filename = "consolidated_salaries.xlsx"
    final_df.to_excel(output_filename)

    print(f"\nSuccess! Consolidated data saved to '{output_filename}'")
    print(f"Total employees: {len(final_df)}")
    print(f"Total months: {len(final_df.columns)}")
    print(f"\nPreview:")
    print(final_df.head(10))

if __name__ == "__main__":
    main()
