"""
Lagrange Polynomial Regression

Reads x/y data from a CSV, asks the user for the data range and polynomial
degree, fits the polynomial via least-squares, and writes the coefficients
to column E of an Excel workbook (E1 = highest-degree term, one per cell).
The original x/y data is preserved in columns A and B of the same workbook.
"""

import os
import csv
import numpy as np
import openpyxl
from openpyxl import Workbook


# ---------------------------------------------------------------------------
# I/O helpers
# ---------------------------------------------------------------------------

def ask_int(prompt, min_val=None, max_val=None, allow_blank=False):
    while True:
        raw = input(prompt).strip()
        if allow_blank and raw == "":
            return None
        try:
            val = int(raw)
            if min_val is not None and val < min_val:
                print(f"  Must be >= {min_val}.")
                continue
            if max_val is not None and val > max_val:
                print(f"  Must be <= {max_val}.")
                continue
            return val
        except ValueError:
            print("  Please enter a whole number.")


def ask_file(prompt):
    while True:
        path = input(prompt).strip().strip('"').strip("'")
        if os.path.isfile(path):
            return path
        print(f"  File not found: {path!r}")


# ---------------------------------------------------------------------------
# CSV reading
# ---------------------------------------------------------------------------

def load_csv(path):
    """Return (headers, rows) where rows is a list of raw string lists."""
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    if not rows:
        raise ValueError("CSV file is empty.")

    # Detect and separate header row
    if rows and not _numeric(rows[0][0]):
        return rows[0], rows[1:]
    return None, rows


def _numeric(s):
    try:
        float(s)
        return True
    except (ValueError, TypeError):
        return False


def parse_data(rows, row_start, row_end, x_col, y_col):
    """Extract float x/y arrays from the selected row slice (1-based)."""
    subset = rows[row_start - 1 : row_end]          # slice is 0-based internally
    x, y = [], []
    for r in subset:
        try:
            x.append(float(r[x_col]))
            y.append(float(r[y_col]))
        except (ValueError, IndexError):
            pass
    if len(x) < 2:
        raise ValueError("Fewer than 2 numeric points in the selected range.")
    return np.array(x), np.array(y)


# ---------------------------------------------------------------------------
# Regression
# ---------------------------------------------------------------------------

def fit_polynomial(x, y, degree):
    """Least-squares polynomial fit; returns coefficients highest-power first."""
    if degree >= len(x):
        raise ValueError(
            f"Degree ({degree}) must be less than the number of points ({len(x)})."
        )
    return np.polyfit(x, y, degree)


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

def write_excel(path, headers, all_rows, x_col, y_col, coefficients, degree):
    wb = Workbook()
    ws = wb.active
    ws.title = "Lagrange Fit"

    # ---- Column headers (A / B / E) ----
    x_hdr = headers[x_col] if headers else "x"
    y_hdr = headers[y_col] if headers else "y"
    ws["A1"] = x_hdr
    ws["B1"] = y_hdr
    ws["E1"] = f"Coeff (deg {degree}, high→low)"

    # ---- Original data in A / B ----
    for i, row in enumerate(all_rows, start=2):
        try:
            ws.cell(row=i, column=1, value=float(row[x_col]))
            ws.cell(row=i, column=2, value=float(row[y_col]))
        except (ValueError, IndexError):
            pass

    # ---- Coefficients in E (E2 onward, one per cell) ----
    for i, c in enumerate(coefficients, start=2):
        ws.cell(row=i, column=5, value=float(c))

    wb.save(path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=" * 50)
    print("   Lagrange Polynomial Regression")
    print("=" * 50)

    # 1. CSV file
    csv_path = ask_file("\nCSV file path: ")
    headers, data_rows = load_csv(csv_path)
    n_rows = len(data_rows)

    if headers:
        print(f"  Header detected: {headers}")
    print(f"  Data rows available: {n_rows}")

    # 2. Column indices
    print("\nColumn indices (0-based). Press Enter for defaults x=0, y=1.")
    x_raw = input("  X column index [0]: ").strip()
    y_raw = input("  Y column index [1]: ").strip()
    x_col = int(x_raw) if x_raw else 0
    y_col = int(y_raw) if y_raw else 1

    # 3. Row range
    print(f"\nRow range to use for fitting (1 – {n_rows}, press Enter for all):")
    row_start = ask_int(f"  Start row [1]:      ", 1, n_rows, allow_blank=True) or 1
    row_end   = ask_int(f"  End row   [{n_rows}]: ", row_start, n_rows, allow_blank=True) or n_rows

    # 4. Load selected data
    try:
        x, y = parse_data(data_rows, row_start, row_end, x_col, y_col)
    except ValueError as e:
        print(f"\nError: {e}")
        return

    n_pts = len(x)
    print(f"\n  Points loaded : {n_pts}")
    print(f"  x range       : [{x.min():.6g}, {x.max():.6g}]")
    print(f"  y range       : [{y.min():.6g}, {y.max():.6g}]")

    # 5. Polynomial degree
    max_deg = n_pts - 1
    degree = ask_int(
        f"\nPolynomial degree (1 – {max_deg}): ", min_val=1, max_val=max_deg
    )

    # 6. Fit
    try:
        coeffs = fit_polynomial(x, y, degree)
    except ValueError as e:
        print(f"\nError: {e}")
        return

    # 7. Display coefficients
    print(f"\nCoefficients for degree-{degree} polynomial (highest power first):")
    print(f"  {'Power':<8}  Coefficient")
    print(f"  {'-'*8}  {'-'*20}")
    for i, c in enumerate(coeffs):
        power = degree - i
        print(f"  x^{power:<6}  {c:.10g}")

    # 8. Output Excel
    base = os.path.splitext(csv_path)[0]
    default_out = base + "_lagrange.xlsx"
    out_raw = input(f"\nOutput Excel path [{default_out}]: ").strip()
    out_path = out_raw if out_raw else default_out

    try:
        write_excel(out_path, headers, data_rows, x_col, y_col, coeffs, degree)
    except Exception as e:
        print(f"\nError writing Excel file: {e}")
        return

    print(f"\nCoefficients written to '{out_path}'")
    print(f"  Column E, rows 2 – {len(coeffs) + 1}  (one coefficient per cell)")
    print("\nDone.")


if __name__ == "__main__":
    main()
