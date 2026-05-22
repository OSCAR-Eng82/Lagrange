"""
Lagrange Polynomial Regression
Reads x/y data from a CSV, fits a polynomial of the requested degree using
Lagrange interpolation (via numpy.polyfit on the sampled points), and writes
the resulting coefficients to column E of an Excel workbook, one per cell
starting at E1 (highest degree first).
"""

import os
import csv
import numpy as np
import openpyxl
from openpyxl import load_workbook


def read_csv(filepath, x_col=0, y_col=1, row_start=None, row_end=None):
    """Return x and y arrays from a CSV file (1-based row range, inclusive)."""
    with open(filepath, newline="") as f:
        reader = list(csv.reader(f))

    # Auto-detect header: skip first row if it is non-numeric
    data_rows = reader
    header_offset = 0
    if reader and not _is_numeric(reader[0][x_col]):
        data_rows = reader[1:]
        header_offset = 1

    total = len(data_rows)

    # Convert 1-based user range to 0-based indices within data_rows
    r_start = (row_start - 1 - header_offset) if row_start else 0
    r_end   = (row_end   - 1 - header_offset) if row_end   else total - 1

    r_start = max(0, r_start)
    r_end   = min(total - 1, r_end)

    if r_start > r_end:
        raise ValueError(
            f"Empty data range: rows {row_start}–{row_end} "
            f"(data has {total} data rows after header)"
        )

    x, y = [], []
    for row in data_rows[r_start : r_end + 1]:
        try:
            x.append(float(row[x_col]))
            y.append(float(row[y_col]))
        except (ValueError, IndexError):
            pass  # skip blank / non-numeric rows silently

    if len(x) < 2:
        raise ValueError("Need at least 2 numeric data points in the chosen range.")

    return np.array(x), np.array(y)


def _is_numeric(value):
    try:
        float(value)
        return True
    except (ValueError, TypeError):
        return False


def lagrange_coefficients(x, y, degree):
    """
    Fit a polynomial of `degree` to (x, y) via least-squares (numpy.polyfit).
    Returns coefficients from highest to lowest power, matching the standard
    Lagrange polynomial representation.
    """
    if degree >= len(x):
        raise ValueError(
            f"Polynomial degree ({degree}) must be less than the number of "
            f"data points ({len(x)})."
        )
    coeffs = np.polyfit(x, y, degree)
    return coeffs  # shape: (degree+1,), highest power first


def write_coefficients_to_excel(coeffs, xlsx_path):
    """Write coefficients to column E starting at E1, one per cell."""
    if os.path.exists(xlsx_path):
        wb = load_workbook(xlsx_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active

    for i, coeff in enumerate(coeffs, start=1):
        ws[f"E{i}"] = float(coeff)

    wb.save(xlsx_path)
    print(f"\nCoefficients written to '{xlsx_path}' in column E (E1:E{len(coeffs)}).")


def prompt_int(prompt, min_val=None, max_val=None):
    while True:
        raw = input(prompt).strip()
        if not raw:
            return None
        try:
            val = int(raw)
            if min_val is not None and val < min_val:
                print(f"  Value must be >= {min_val}.")
                continue
            if max_val is not None and val > max_val:
                print(f"  Value must be <= {max_val}.")
                continue
            return val
        except ValueError:
            print("  Please enter an integer.")


def main():
    print("=== Lagrange Polynomial Regression ===\n")

    # --- CSV path ---
    while True:
        csv_path = input("Enter path to CSV file: ").strip().strip('"').strip("'")
        if os.path.isfile(csv_path):
            break
        print(f"  File not found: '{csv_path}'. Try again.")

    # --- Column selection ---
    print("\nWhich columns contain x and y data? (0-based index, default: x=0, y=1)")
    x_col_in = input("  X column index [0]: ").strip()
    y_col_in = input("  Y column index [1]: ").strip()
    x_col = int(x_col_in) if x_col_in else 0
    y_col = int(y_col_in) if y_col_in else 1

    # --- Row range ---
    print("\nRow range of data to use (1-based, press Enter to use all rows):")
    row_start = prompt_int("  Start row [1]: ", min_val=1)
    row_end   = prompt_int("  End row   [last]: ", min_val=1)

    # --- Load data ---
    try:
        x, y = read_csv(csv_path, x_col=x_col, y_col=y_col,
                         row_start=row_start, row_end=row_end)
    except Exception as e:
        print(f"\nError reading CSV: {e}")
        return

    print(f"\nLoaded {len(x)} data points.")
    print(f"  x range: [{x.min():.6g}, {x.max():.6g}]")
    print(f"  y range: [{y.min():.6g}, {y.max():.6g}]")

    # --- Polynomial degree ---
    max_degree = len(x) - 1
    degree = prompt_int(
        f"\nPolynomial degree (1–{max_degree}): ", min_val=1, max_val=max_degree
    )
    if degree is None:
        print("No degree entered. Exiting.")
        return

    # --- Compute coefficients ---
    try:
        coeffs = lagrange_coefficients(x, y, degree)
    except Exception as e:
        print(f"\nError computing coefficients: {e}")
        return

    print(f"\nPolynomial coefficients (highest degree first):")
    for i, c in enumerate(coeffs):
        power = degree - i
        print(f"  x^{power}: {c:.10g}")

    # --- Output Excel file ---
    base = os.path.splitext(csv_path)[0]
    default_xlsx = base + "_lagrange.xlsx"
    xlsx_input = input(f"\nOutput Excel file [{default_xlsx}]: ").strip()
    xlsx_path = xlsx_input if xlsx_input else default_xlsx

    try:
        write_coefficients_to_excel(coeffs, xlsx_path)
    except Exception as e:
        print(f"\nError writing Excel file: {e}")
        return

    print("\nDone.")


if __name__ == "__main__":
    main()
