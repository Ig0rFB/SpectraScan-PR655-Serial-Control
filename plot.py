import argparse
import csv
from pathlib import Path

import matplotlib.pyplot as plt


def load_spectrum(csv_path: Path) -> tuple[list[float], list[float]]:
    """Load wavelength and radiance columns from a PR measurement CSV file."""
    wavelengths: list[float] = []
    radiance: list[float] = []

    with csv_path.open("r", newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)

    # Find the spectral table header written by measurement.py.
    header_index = None
    for i, row in enumerate(rows):
        if len(row) >= 2 and row[0].strip() == "Wavelength (nm)":
            header_index = i
            break

    if header_index is None:
        raise ValueError("Could not find spectral data header in CSV file.")

    for row in rows[header_index + 1 :]:
        if len(row) < 2:
            continue
        left = row[0].strip()
        right = row[1].strip()
        if not left or not right:
            continue
        try:
            wl = float(left)
            rad = float(right)
        except ValueError:
            continue
        wavelengths.append(wl)
        radiance.append(rad)

    if not wavelengths:
        raise ValueError("No numeric spectral rows found in CSV file.")

    return wavelengths, radiance


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Plot spectral radiance vs wavelength from a PR measurement CSV file."
    )
    parser.add_argument("csv_file", type=Path, help="Path to PR_Measurement_*.csv")
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Optional image output path (e.g. spectrum.png)",
    )
    args = parser.parse_args()

    if not args.csv_file.exists():
        raise FileNotFoundError(f"CSV file not found: {args.csv_file}")

    wavelengths, radiance = load_spectrum(args.csv_file)

    plt.figure(figsize=(10, 5))
    plt.plot(wavelengths, radiance, linewidth=1.8)
    plt.xlabel("Wavelength (nm)")
    plt.ylabel("Spectral Radiance")
    plt.title(f"Spectral Radiance vs Wavelength\n{args.csv_file.name}")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()

    if args.output is not None:
        plt.savefig(args.output, dpi=150)
        print(f"Saved plot to: {args.output}")
    else:
        plt.show()


if __name__ == "__main__":
    main()
