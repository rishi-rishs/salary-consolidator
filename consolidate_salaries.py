import pandas as pd
import sys
import os
from rapidfuzz import process, fuzz
from datetime import datetime

def parse_date_from_sheet_name(sheet_name):
    """
    Attempts to parse a sheet name into a 'MMM-YY' string.
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

def get_employee_id(name, current_salary, prev_month_data, employee_registry, aliases):
    """
    Returns an employee ID (integer) for the given name.
    Uses fuzzy matching against previous month's data.

    employee_registry: dict mapping employee_id -> list of (name, month_index) tuples
    """
    name = str(name).strip()

    # Check if this exact name is a known alias
    if name in aliases:
        return aliases[name]

    # If no previous month data, this is a new employee
    if not prev_month_data:
        return None

    # Fuzzy match against previous month's names
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
    print(f"  Previous Salary: {prev_month_data[best_match]['salary']}")

    if second_match and second_score >= 60:
        print(f"Second best:      '{second_match}' (Score: {second_score:.1f})")
        print(f"  Previous Salary: {prev_month_data[second_match]['salary']}")

    print(f"{'='*60}")

    while True:
        if second_match and second_score >= 60:
            choice = input(f"Match with: [1] {best_match} / [2] {second_match} / [n]ew / [t]ype manual: ").lower().strip()
        else:
            choice = input(f"Match with: [y]es {best_match} / [n]ew / [t]ype manual: ").lower().strip()

        if choice == 'y' or choice == 'yes' or choice == '1':
            emp_id = prev_month_data[best_match]['emp_id']
            aliases[name] = emp_id
            return emp_id
        elif choice == '2' and second_match and second_score >= 60:
            emp_id = prev_month_data[second_match]['emp_id']
            aliases[name] = emp_id
            return emp_id
        elif choice == 'n' or choice == 'new':
            return None
        elif choice == 't' or choice == 'type':
            manual_name = input("Enter the correct existing target name: ").strip()
            if manual_name in prev_month_data:
                emp_id = prev_month_data[manual_name]['emp_id']
                aliases[name] = emp_id
                return emp_id
            else:
                print(f"'{manual_name}' not found in previous month. Treating as new employee.")
                return None
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

    # Employee registry: emp_id -> {names: [(name, month_idx), ...], salaries: {month: salary}}
    employee_registry = {}
    next_emp_id = 1

    # Aliases: name -> emp_id
    aliases = {}

    # Previous month's data: name -> {emp_id, salary}
    prev_month_data = {}

    # Track month order
    months_in_order = []

    # Iterate through sheets
    for month_idx, sheet_name in enumerate(xls.sheet_names):
        print(f"\n{'#'*60}")
        print(f"### Processing Sheet: {sheet_name} ###")
        print(f"{'#'*60}")

        column_header = parse_date_from_sheet_name(sheet_name)
        months_in_order.append(column_header)

        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
        except Exception as e:
            print(f"Skipping sheet '{sheet_name}' due to read error: {e}")
            continue

        if df.shape[1] < 3:
            print(f"Skipping sheet '{sheet_name}': Expected at least 3 columns, found {df.shape[1]}")
            continue

        # Skip header row (first row)
        df = df.iloc[1:]

        # Skip last row if it contains "total"
        if len(df) > 0:
            last_name = str(df.iloc[-1, 1]).strip().lower() if not pd.isna(df.iloc[-1, 1]) else ""
            if "total" in last_name:
                df = df.iloc[:-1]

        # Current month's data
        current_month_data = {}

        for index, row in df.iterrows():
            raw_name = row.iloc[1]  # Column B for employee name
            salary = row.iloc[2]    # Column C for salary (Gross)

            if pd.isna(raw_name):
                continue

            raw_name = str(raw_name).strip()

            # Try to match to existing employee
            emp_id = get_employee_id(
                raw_name,
                salary,
                prev_month_data,
                employee_registry,
                aliases
            )

            if emp_id is None:
                # New employee
                emp_id = next_emp_id
                next_emp_id += 1
                employee_registry[emp_id] = {
                    'names': [],
                    'salaries': {}
                }
                print(f"Added new employee #{emp_id}: '{raw_name}'")

            # Record this name variant with month index (for tracking latest name)
            employee_registry[emp_id]['names'].append((raw_name, month_idx))

            # Record salary for this month
            if column_header in employee_registry[emp_id]['salaries']:
                print(f"Warning: Duplicate entry for employee #{emp_id} in '{column_header}'. Overwriting.")
            employee_registry[emp_id]['salaries'][column_header] = salary

            # Track for next month's comparison
            current_month_data[raw_name] = {'emp_id': emp_id, 'salary': salary}
            aliases[raw_name] = emp_id

        # Update prev_month_data for next iteration
        prev_month_data = current_month_data.copy()

    # Build final output using LATEST name for each employee
    print(f"\n{'#'*60}")
    print("### Finalizing Data ###")
    print(f"{'#'*60}")

    # Create output data
    output_data = []

    for emp_id, emp_data in employee_registry.items():
        # Get the latest name (highest month_idx)
        latest_name = max(emp_data['names'], key=lambda x: x[1])[0]

        row = {'Employee Name': latest_name}
        row.update(emp_data['salaries'])
        output_data.append(row)

    final_df = pd.DataFrame(output_data)

    # Sort columns chronologically (keep Employee Name first)
    date_cols = [c for c in final_df.columns if c != 'Employee Name']
    try:
        date_cols_sorted = sorted(date_cols, key=lambda x: datetime.strptime(x, "%b-%y"))
    except ValueError:
        print("Could not sort columns chronologically. Using original order.")
        date_cols_sorted = date_cols

    final_df = final_df[['Employee Name'] + date_cols_sorted]

    # Sort rows by employee name
    final_df = final_df.sort_values('Employee Name').reset_index(drop=True)

    output_filename = "consolidated_salaries.xlsx"
    final_df.to_excel(output_filename, index=False)

    print(f"\nSuccess! Consolidated data saved to '{output_filename}'")
    print(f"Total employees: {len(final_df)}")
    print(f"Total months: {len(date_cols_sorted)}")
    print(f"\nPreview:")
    print(final_df.head(10).to_string())

if __name__ == "__main__":
    main()
