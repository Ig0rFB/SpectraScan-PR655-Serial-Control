# SpectraScan PR-655 serial control

Small Python utilities to drive a **Photo Research PR-655** (or compatible command set) over **RS-232/USB serial**, capture spectral measurements, and optionally plot saved CSV files.

## Requirements

- Python 3.10+
- [pyserial](https://pypi.org/project/pyserial/) — `pip install pyserial`
- [matplotlib](https://pypi.org/project/matplotlib/) — only needed for `plot.py` — `pip install matplotlib`

## Configuration

Edit **`SERIAL_PORT`** in `measurement.py` if the default does not match your machine:

- **Windows:** defaults to `COM4`
- **Other platforms:** defaults to `/dev/tty.usbserial`

Adjust **`BAUD_RATE`**, **`MEASUREMENT_MODE_LABEL`**, and **`FILENAME_BRIGHTNESS_LABEL`** if your setup or naming convention differs.

## Running measurements

```bash
python measurement.py
```

You will see:

1. **Main menu** — `[S]` single measurement, `[C]` ColorChecker chart, `[Q]` quit  
2. **ColorChecker** — `[A]` all 24 patches, `[S]` one patch (by number or name), `[Q]` cancel  

The script enters remote mode, triggers captures, and writes **CSV** files in the current directory (metadata plus wavelength vs spectral power, typically 380–780 nm). See `pr655_manual.pdf` for instrument details.

## Plotting a saved spectrum

```bash
python plot.py path/to/measurement.csv
python plot.py path/to/measurement.csv --output spectrum.png
```
