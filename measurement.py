import serial
import time
import csv
import re
from datetime import datetime

# --- Configuration ---
SERIAL_PORT = 'COM4' if __import__("platform").system() == "Windows" else '/dev/tty.usbserial'
BAUD_RATE = 9600
TIMEOUT = 20          # Increased timeout as D6 can be a large data packet
PRE_TRIGGER_DELAY_S = 0
MEASUREMENT_MODE_LABEL = "RGBW-Mode"
FILENAME_BRIGHTNESS_LABEL = "100pct"

COLORCHECKER_PATCHES = [
    ("Dark Skin", (115, 82, 68)),
    ("Light Skin", (194, 150, 130)),
    ("Blue Sky", (98, 122, 157)),
    ("Foliage", (87, 108, 67)),
    ("Blue Flower", (133, 128, 177)),
    ("Bluish Green", (103, 189, 170)),
    ("Orange", (214, 126, 44)),
    ("Purplish Blue", (80, 91, 166)),
    ("Moderate Red", (193, 90, 99)),
    ("Purple", (94, 60, 108)),
    ("Yellow Green", (157, 188, 64)),
    ("Orange Yellow", (224, 163, 46)),
    ("Blue", (56, 61, 150)),
    ("Green", (70, 148, 73)),
    ("Red", (175, 54, 60)),
    ("Yellow", (231, 199, 31)),
    ("Magenta", (187, 86, 149)),
    ("Cyan", (8, 133, 161)),
    ("White", (243, 243, 242)),
    ("Neutral 8", (200, 200, 200)),
    ("Neutral 6.5", (160, 160, 160)),
    ("Neutral 5", (122, 122, 121)),
    ("Neutral 3.5", (85, 85, 85)),
    ("Black", (52, 52, 52)),
]


def choose_patch_sequence() -> list[tuple[int, str, tuple[int, int, int]]]:
    """Prompt user to choose all patches or a single patch in ColorChecker mode."""
    indexed = [
        (i, name, rgb)
        for i, (name, rgb) in enumerate(COLORCHECKER_PATCHES, start=1)
    ]

    while True:
        selection = input(
            "\n--- ColorChecker Measurement ---\n"
            "Select capture scope:\n"
            "  [A] All patches (1-24)\n"
            "  [S] Single patch\n"
            "  [Q] Back to main menu\n"
            "Choice: "
        ).strip().lower()

        if selection == "a":
            return indexed

        if selection == "s":
            patch_selection = input(
                "Enter patch number (1-24) or patch name (e.g. Orange): "
            ).strip()

            if patch_selection.isdigit():
                idx = int(patch_selection)
                if 1 <= idx <= len(COLORCHECKER_PATCHES):
                    name, rgb = COLORCHECKER_PATCHES[idx - 1]
                    return [(idx, name, rgb)]
                print(f"Invalid patch number: {patch_selection}. Please try again.")
                continue

            normalized = patch_selection.casefold()
            for idx, name, rgb in indexed:
                if name.casefold() == normalized:
                    return [(idx, name, rgb)]

            print(f"Unknown patch name: {patch_selection!r}. Please try again.")
            continue

        if selection == "q":
            return []

        print("Invalid choice. Please select A, S, or Q.")


def choose_measurement_workflow() -> str:
    """Choose measurement mode from the main menu."""
    while True:
        selection = input(
            "\n=== Spectral Measurement Menu ===\n"
            "Select mode:\n"
            "  [S] Single measurement\n"
            "  [C] ColorChecker chart measurement\n"
            "  [Q] Quit\n"
            "Choice: "
        ).strip().lower()
        if selection == "s":
            return "single"
        if selection == "c":
            return "colorchecker"
        if selection == "q":
            return "quit"
        print("Invalid choice. Please select S, C, or Q.")


def read_response(ser: serial.Serial, timeout_s: float = TIMEOUT) -> str:
    """Reads a line-like device response until CR/LF or timeout."""
    deadline = time.time() + timeout_s
    buf = bytearray()

    while time.time() < deadline:
        waiting = ser.in_waiting
        if waiting > 0:
            chunk = ser.read(waiting)
            if chunk:
                buf.extend(chunk)
                if b"\r" in chunk or b"\n" in chunk:
                    break
        else:
            time.sleep(0.02)

    return buf.decode("ascii", errors="replace").strip()


def read_response_until_idle(
    ser: serial.Serial,
    overall_timeout_s: float = TIMEOUT,
    idle_timeout_s: float = 0.5,
) -> str:
    """Reads all incoming bytes until the stream is idle or timeout is reached."""
    deadline = time.time() + overall_timeout_s
    last_data_time = time.time()
    buf = bytearray()

    while time.time() < deadline:
        waiting = ser.in_waiting
        if waiting > 0:
            chunk = ser.read(waiting)
            if chunk:
                buf.extend(chunk)
                last_data_time = time.time()
        else:
            if buf and (time.time() - last_data_time) >= idle_timeout_s:
                break
            time.sleep(0.02)

    return buf.decode("ascii", errors="replace").strip()


def send_command(
    ser: serial.Serial,
    command: str,
    *,
    line_ending: str = "\r",
    wait_for_response: bool = True,
    multiline: bool = False,
    timeout_s: float = TIMEOUT,
) -> str:
    """Sends an ASCII command and optionally waits for a response."""
    cmd_string = f"{command}{line_ending}" if line_ending else command
    ser.write(cmd_string.encode("ascii"))
    ser.flush()

    if not wait_for_response:
        return ""

    # Device responses may end with CR or LF depending on serial terminal settings.
    if multiline:
        return read_response_until_idle(ser, overall_timeout_s=timeout_s)
    return read_response(ser, timeout_s=timeout_s)


def send_with_fallbacks(
    ser: serial.Serial,
    command: str,
    *,
    multiline: bool = False,
    timeout_s: float = TIMEOUT,
) -> str:
    """Tries CR then CRLF line endings for instruments with terminal differences."""
    for line_ending in ("\r", "\r\n"):
        ser.reset_input_buffer()
        reply = send_command(
            ser,
            command,
            line_ending=line_ending,
            wait_for_response=True,
            multiline=multiline,
            timeout_s=timeout_s,
        )
        if reply:
            return reply
    return ""


def is_success_status(status: str) -> bool:
    """Treat any all-zero status code (e.g., 000 or 00000) as success."""
    cleaned = status.strip()
    return cleaned != "" and all(ch == "0" for ch in cleaned)


def get_status_token(reply: str) -> str:
    """Returns the leading status token from a device reply."""
    return reply.split(",", 1)[0].strip()


def parse_d120(reply: str) -> dict:
    """Parses D120 instrument spectral configuration reply."""
    parts = [p.strip() for p in reply.split(",") if p.strip() != ""]
    if len(parts) < 6:
        raise ValueError(f"Unexpected D120 format: {reply!r}")

    status = parts[0]
    if not is_success_status(status):
        raise ValueError(f"D120 returned error code: {status}")

    return {
        "status": status,
        "points": int(float(parts[1])),
        "bandwidth_nm": float(parts[2]),
        "start_nm": float(parts[3]),
        "end_nm": float(parts[4]),
        "increment_nm": float(parts[5]),
    }


def parse_m5_response(raw: str) -> tuple[dict, list[float], list[float]]:
    """Parses M5/D5 response into metadata and wavelength/radiance arrays."""
    lines = [ln.strip() for ln in raw.replace("\r", "\n").split("\n") if ln.strip()]
    if not lines:
        raise ValueError("Empty M5 response")

    header = [p.strip() for p in lines[0].split(",") if p.strip() != ""]
    if len(header) < 5:
        raise ValueError(f"Unexpected M5 header format: {lines[0]!r}")

    status = header[0]
    if not is_success_status(status):
        raise ValueError(f"M5 returned error code: {status}")

    units_code = header[1]
    m5_meta = {
        "status": status,
        "units_code": units_code,
        "peak_wavelength_nm": float(header[2]),
        "integrated_radiometric": float(header[3]),
        "integrated_photon_radiometric": float(header[4]),
    }

    wavelengths: list[float] = []
    radiance_values: list[float] = []

    # Typical format: one WL,radiance pair per following line.
    for line in lines[1:]:
        if "," not in line:
            continue
        left, right = [x.strip() for x in line.split(",", 1)]
        try:
            wl = float(left)
            rad = float(right)
        except ValueError:
            continue
        wavelengths.append(wl)
        radiance_values.append(rad)

    # Fallback in case pairs were returned on a single comma-delimited line.
    if not wavelengths and len(header) > 6:
        tail = header[5:]
        for i in range(0, len(tail) - 1, 2):
            try:
                wl = float(tail[i])
                rad = float(tail[i + 1])
            except ValueError:
                continue
            wavelengths.append(wl)
            radiance_values.append(rad)

    if not wavelengths:
        raise ValueError("No wavelength/radiance pairs found in M5 response")

    return m5_meta, wavelengths, radiance_values


def sanitize_filename_component(value: str) -> str:
    """Converts arbitrary label text to a filesystem-friendly token."""
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return safe.strip("._") or "untitled"


def save_to_csv(wavelengths, values, metadata, filename_base="PR_Measurement"):
    """Saves the spectral data to a CSV file."""
    safe_base = sanitize_filename_component(filename_base)
    filename = f"{safe_base}.csv"

    # Avoid accidental overwrite by appending an incrementing suffix.
    counter = 2
    while True:
        try:
            with open(filename, mode='x', newline='') as file:
                writer = csv.writer(file)
                # Write metadata headers
                writer.writerow(["Metadata", "Value"])
                for key, val in metadata.items():
                    writer.writerow([key, val])

                writer.writerow([]) # Spacer

                # Write Spectral Data
                writer.writerow(["Wavelength (nm)", "Spectral Power"])
                for wl, pwr in zip(wavelengths, values):
                    writer.writerow([wl, pwr])
            break
        except FileExistsError:
            filename = f"{safe_base}_{counter}.csv"
            counter += 1
            
    print(f"Data successfully saved to: {filename}")


def enter_remote_mode(ser: serial.Serial) -> None:
    """Enter remote mode by sending PHOTO as immediate characters."""
    for ch in "PHOTO":
        ser.write(ch.encode("ascii"))
        ser.flush()
        time.sleep(0.05)
    time.sleep(0.2)


def wait_for_trigger(patch_index: int, total_patches: int, patch_name: str, rgb: tuple[int, int, int]) -> bool:
    """Wait for a trigger key. Returns False if user requested quit."""
    message = (
        f"\nPatch {patch_index}/{total_patches}: {patch_name} "
        f"(RGB {rgb[0]},{rgb[1]},{rgb[2]}). "
        "Point the meter, then press any key to measure or Q to quit..."
    )
    print(message)

    try:
        import msvcrt

        key = msvcrt.getch()
        if key in (b"q", b"Q"):
            return False
        return True
    except ImportError:
        key = input("Type Q to quit, or press Enter to measure: ")
        return key.strip().lower() != "q"


def wait_for_custom_trigger() -> bool:
    """Wait for custom measurement trigger. Returns False if user quits."""
    print("\nCustom measurement: point the meter, then press any key to measure or Q to quit...")
    try:
        import msvcrt

        key = msvcrt.getch()
        return key not in (b"q", b"Q")
    except ImportError:
        key = input("Type Q to quit, or press Enter to measure: ")
        return key.strip().lower() != "q"


def perform_spectral_measurement(ser: serial.Serial) -> tuple[dict, dict, list[float], list[float]]:
    """Take one M5 spectral measurement and parse metadata plus spectral values."""
    measure_reply = send_with_fallbacks(
        ser,
        "M5",
        multiline=True,
        timeout_s=TIMEOUT + 10,
    )
    m5_meta, wavelengths, spectral_values = parse_m5_response(measure_reply)

    d120_reply = send_with_fallbacks(ser, "D120")
    try:
        d120 = parse_d120(d120_reply)
    except ValueError:
        d120 = {
            "status": "",
            "points": len(wavelengths),
            "bandwidth_nm": 8.0,
            "start_nm": min(wavelengths),
            "end_nm": max(wavelengths),
            "increment_nm": (wavelengths[1] - wavelengths[0]) if len(wavelengths) > 1 else 0.0,
        }

    filtered = [
        (wl, rad)
        for wl, rad in zip(wavelengths, spectral_values)
        if 380.0 <= wl <= 780.0
    ]
    if not filtered:
        raise ValueError("No spectral points found in 380-780 nm range")

    wavelengths_filtered = [wl for wl, _ in filtered]
    values_filtered = [rad for _, rad in filtered]
    return m5_meta, d120, wavelengths_filtered, values_filtered


def custom_measurement(ser: serial.Serial) -> bool:
    """Run one custom measurement and save it using a user-provided filename."""
    requested_name = input("Enter filename for this custom measurement: ").strip()
    filename_base = requested_name if requested_name else "Custom_Measurement"

    if not wait_for_custom_trigger():
        print("Custom measurement canceled by user.")
        return False

    print(f"Waiting {PRE_TRIGGER_DELAY_S} seconds before trigger...")
    time.sleep(PRE_TRIGGER_DELAY_S)
    print("Taking measurement... (Please wait for shutter)")

    try:
        m5_meta, d120, wavelengths, spectral_values = perform_spectral_measurement(ser)
    except ValueError as err:
        print(f"Custom measurement error: {err}")
        return False

    meta_dict = {
        "Measurement Type": "Custom",
        "Requested Filename": filename_base,
        "Status": m5_meta["status"],
        "Units Code": m5_meta["units_code"],
        "Peak WL NM": m5_meta["peak_wavelength_nm"],
        "Integrated Radiometric": m5_meta["integrated_radiometric"],
        "Integrated Photon Radiometric": m5_meta["integrated_photon_radiometric"],
        "Instrument Start NM": d120["start_nm"],
        "Instrument End NM": d120["end_nm"],
        "Instrument Increment NM": d120["increment_nm"],
        "Instrument Bandwidth NM": d120["bandwidth_nm"],
        "Saved Range Start NM": 380.0,
        "Saved Range End NM": 780.0,
        "Timestamp": datetime.now().isoformat(),
    }

    save_to_csv(wavelengths, spectral_values, meta_dict, filename_base=filename_base)
    return True


def run_spectral_test():
    try:
        workflow = choose_measurement_workflow()
        if workflow == "quit":
            print("No measurement selected. Exiting.")
            return

        selected_patches = choose_patch_sequence() if workflow == "colorchecker" else []
        if workflow == "colorchecker" and not selected_patches:
            print("ColorChecker measurement cancelled. Exiting.")
            return

        with serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=TIMEOUT) as ser:
            print(f"Initialising connection on {SERIAL_PORT}...")

            # Clear any stale bytes before starting command flow.
            ser.reset_input_buffer()
            ser.reset_output_buffer()

            print("Entering remote mode...")
            enter_remote_mode(ser)

            if workflow == "single":
                custom_measurement(ser)
                send_command(ser, "Q", line_ending="\r", wait_for_response=False)
                print("Session complete. Instrument returned to local mode.")
                return

            print("Interactive ColorChecker capture session started.")
            print("Press any key to capture each patch, or Q to stop early.")

            saved_count = 0
            total_selected = len(selected_patches)
            for sequence_index, (idx, patch_name, rgb) in enumerate(selected_patches, start=1):
                if not wait_for_trigger(sequence_index, total_selected, patch_name, rgb):
                    print("Stop requested by user.")
                    break

                print(f"Waiting {PRE_TRIGGER_DELAY_S} seconds before trigger...")
                time.sleep(PRE_TRIGGER_DELAY_S)
                print("Taking measurement... (Please wait for shutter)")
                try:
                    m5_meta, d120, wavelengths, spectral_values = perform_spectral_measurement(ser)
                except ValueError as err:
                    print(f"Measurement error for patch '{patch_name}': {err}")
                    continue

                meta_dict = {
                    "Patch Name": patch_name,
                    "Patch RGB": f"{rgb[0]},{rgb[1]},{rgb[2]}",
                    "Status": m5_meta["status"],
                    "Units Code": m5_meta["units_code"],
                    "Peak WL NM": m5_meta["peak_wavelength_nm"],
                    "Integrated Radiometric": m5_meta["integrated_radiometric"],
                    "Integrated Photon Radiometric": m5_meta["integrated_photon_radiometric"],
                    "Instrument Start NM": d120["start_nm"],
                    "Instrument End NM": d120["end_nm"],
                    "Instrument Increment NM": d120["increment_nm"],
                    "Instrument Bandwidth NM": d120["bandwidth_nm"],
                    "Saved Range Start NM": 380.0,
                    "Saved Range End NM": 780.0,
                    "Timestamp": datetime.now().isoformat(),
                }

                patch_code = f"Patch{idx:02d}"
                filename_base = (
                    f"{MEASUREMENT_MODE_LABEL}_{FILENAME_BRIGHTNESS_LABEL}_{patch_code}_{patch_name}"
                )
                save_to_csv(wavelengths, spectral_values, meta_dict, filename_base=filename_base)
                saved_count += 1

            # Exit Remote Mode
            send_command(ser, "Q", line_ending="\r", wait_for_response=False)
            print(
                f"Session complete. Saved {saved_count} patch measurement(s). "
                "Instrument returned to local mode."
            )

    except serial.SerialException as e:
        print(f"Serial Port Error: {e}. Check your COM port number.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    run_spectral_test()