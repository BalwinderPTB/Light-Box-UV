# instrument_control.py
import sys
import time
import datetime
import os
import serial
import serial.tools.list_ports

from PyQt5.QtWidgets import (
    QApplication, QWidget, QGridLayout, QHBoxLayout, QVBoxLayout,
    QPushButton, QLabel, QInputDialog, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt, QTimer

# -----------------------
# Fake serial for simulation
# -----------------------
class FakeSerial:
    def __init__(self):
        self.is_open = True
        print("[SIM] FakeSerial: simulation mode active")

    def write(self, data: bytes):
        # data is bytes; decode safely for readable console output
        try:
            s = data.decode('utf-8', errors='replace')
        except Exception:
            s = str(data)
        print(f"[SIM] -> {s.strip()}")

    def close(self):
        self.is_open = False
        print("[SIM] FakeSerial closed")


# -----------------------
# Instrument UI & Logic
# -----------------------
class InstrumentUI(QWidget):
    def __init__(self):
        super().__init__()

        # --- serial connection state (populated by connect_serial) ---
        self.serial_conn = None
        self.simulation_mode = False
        self.connected_port = None

        # Try to connect BEFORE building the UI (blocks until success or quit)
        self.connect_serial()  # will sys.exit() if user quits

        # --- instrument state ---
        self.rows, self.cols = 8, 12
        self.instrument_on = False
        self.start_time = None           # float epoch seconds when ON pressed
        self.start_datetime = None       # datetime when ON pressed

        # Brightness matrix (current values) — bytes/ints 0..100
        self.brightness_values = [[0 for _ in range(self.cols)] for _ in range(self.rows)]

        # Time-weighted average tracking (accumulate brightness * seconds)
        # Only used during an ON session
        self.brightness_time_sum = [[0.0 for _ in range(self.cols)] for _ in range(self.rows)]
        # last time we accounted for the current brightness (epoch seconds)
        # set to None while instrument is OFF; set to start_time when turned ON
        self.last_update_time = [[None for _ in range(self.cols)] for _ in range(self.rows)]

        # After connecting, send initial zero matrix to Arduino (to set known state)
        self.send_matrix()

        # Build UI (keeps your preferred layout and sizes)
        self.init_ui()

        # Timer: updates elapsed ON time and accumulates per-second brightness-time
        self.timer = QTimer()
        self.timer.timeout.connect(self._timer_tick)
        self.timer.start(1000)  # tick every second

    # -------------------------
    # Serial connection logic
    # -------------------------
    def connect_serial(self):
        """
        Discover ports and ask user to choose. If user types 'q' -> exit.
        If user types 's' -> enter simulation mode.
        Loops until success or quit.
        """
        while True:
            ports = list(serial.tools.list_ports.comports())
            print("\nDetected serial ports:")
            for i, p in enumerate(ports):
                print(f"  {i+1}) {p.device} — {p.description}")
            print("  s) Simulation mode (no Arduino)")
            print("  q) Quit")

            choice = input("Select port number, or 's' for simulation, 'q' to quit: ").strip().lower()
            if choice == 'q':
                print("User chose to quit. Exiting.")
                sys.exit(0)
            if choice == 's':
                self.simulation_mode = True
                self.serial_conn = FakeSerial()
                return

            # If no ports found, just loop again (user can plug device and press Enter)
            if not ports:
                print("No serial ports found. Plug in device and press Enter to retry, or 's' for simulation, 'q' to quit.")
                continue

            # try to parse numeric selection
            if choice.isdigit():
                idx = int(choice) - 1
                if 0 <= idx < len(ports):
                    port_name = ports[idx].device
                    try:
                        conn = serial.Serial(port_name, 115200, timeout=1)
                        self.serial_conn = conn
                        self.simulation_mode = False
                        self.connected_port = port_name
                        print(f"Connected to {port_name}")
                        return
                    except Exception as e:
                        print(f"Failed to open {port_name}: {e}")
                        # loop and let user choose again
                        continue
                else:
                    print("Invalid port number. Try again.")
                    continue

            print("Invalid input. Enter a number, 's' or 'q'.")

    # -------------------------
    # UI Construction
    # -------------------------
    def init_ui(self):
        self.setWindowTitle("Scientific Instrument Control Panel")
        self.resize(1400, 800)

        main_layout = QHBoxLayout(self)

        # Left: 8x12 grid of buttons
        grid_layout = QGridLayout()
        self.buttons = [[None for _ in range(self.cols)] for _ in range(self.rows)]
        for r in range(self.rows):
            for c in range(self.cols):
                btn = QPushButton(self._button_label_text(r, c))
                btn.setFixedSize(110, 85)     # your preferred size
                btn.setStyleSheet(self._button_style(self.brightness_values[r][c]))
                btn.clicked.connect(lambda _, rr=r, cc=c: self._on_button_click(rr, cc))
                grid_layout.addWidget(btn, r, c)
                self.buttons[r][c] = btn

        main_layout.addLayout(grid_layout, stretch=5)

        # Right: compact control panel
        control_layout = QVBoxLayout()

        # ON/OFF button (starts OFF, red)
        self.onoff_btn = QPushButton("Turn ON")
        self.onoff_btn.setFixedHeight(50)
        self._update_onoff_style()
        self.onoff_btn.clicked.connect(self._toggle_instrument)
        control_layout.addWidget(self.onoff_btn)

        # Running time label (larger font)
        self.runtime_label = QLabel("Running time: 0s")
        self.runtime_label.setAlignment(Qt.AlignCenter)
        self.runtime_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        control_layout.addWidget(self.runtime_label)

        control_layout.addStretch()
        main_layout.addLayout(control_layout, stretch=1)

        self.setLayout(main_layout)

    # -------------------------
    # Button helpers
    # -------------------------
    def _button_label_text(self, r, c):
        cur = self.brightness_values[r][c]
        avg = self._calculate_average(r, c)  # returns 0.0 if not in session
        return f"R{r+1}C{c+1}\nNow: {cur}%\nAvg: {avg:.1f}%"

    def _button_style(self, brightness):
        if brightness == 0:
            bg = "rgb(100,100,100)"
            style = "font-style: italic;"
        else:
            if brightness <= 50:
                r = 255
                g = int((brightness / 50) * 255)
                b = 0
            else:
                r = int(255 - ((brightness - 50) / 50) * 255)
                g = 255
                b = 0
            bg = f"rgb({r},{g},{b})"
            style = "font-style: normal;"
        return f"""
            QPushButton {{
                background-color: {bg};
                color: black;
                font-size: 14px;
                font-weight: bold;
                border-radius: 6px;
                {style}
            }}
        """

    # Called when a grid button is clicked
    def _on_button_click(self, row, col):
        # Ask user for brightness (0-100)
        value, ok = QInputDialog.getInt(
            self, "Set Brightness",
            f"Enter brightness for R{row+1}C{col+1} (0-100):",
            self.brightness_values[row][col], 0, 100, 1
        )
        if not ok:
            return

        old_val = self.brightness_values[row][col]
        now = time.time()

        if self.instrument_on:
            # We are in ON session: accumulate time from last_update_time to now
            lut = self.last_update_time[row][col]
            if lut is not None:
                dt = now - lut
                if dt > 0:
                    # accumulate previous brightness contribution
                    self.brightness_time_sum[row][col] += old_val * dt
            # update last update time to now and set new brightness
            self.last_update_time[row][col] = now
            self.brightness_values[row][col] = value
            # immediately send updated matrix to Arduino
            self.send_matrix()
        else:
            # Pre-ON: only update local matrix, no averaging, no sending
            self.brightness_values[row][col] = value

        # Update button label & style for immediate feedback
        btn = self.buttons[row][col]
        btn.setText(self._button_label_text(row, col))
        btn.setStyleSheet(self._button_style(self.brightness_values[row][col]))

    # -------------------------
    # Turn ON / OFF behavior
    # -------------------------
    def _toggle_instrument(self):
        if not self.instrument_on:
            self._turn_on()
        else:
            self._turn_off()

    def _turn_on(self):
        # Start session: set timers and reset accumulators
        now = time.time()
        self.instrument_on = True
        self.start_time = now
        self.start_datetime = datetime.datetime.now()

        # Reset time-sums for this session
        self.brightness_time_sum = [[0.0 for _ in range(self.cols)] for _ in range(self.rows)]

        # Set last_update_time to now for all cells (so we start fresh)
        for r in range(self.rows):
            for c in range(self.cols):
                self.last_update_time[r][c] = now

        # Send the current matrix (which may contain pre-ON configured values)
        self.send_matrix()
        self._update_onoff_style()
        # Refresh UI labels (averages were zero at start)
        self._refresh_all_buttons()

    def _turn_off(self):
        # Stop session: we must first accumulate the last partial interval into brightness_time_sum
        if not self.instrument_on:
            return

        now = time.time()
        # accumulate final period for each LED
        for r in range(self.rows):
            for c in range(self.cols):
                lut = self.last_update_time[r][c]
                if lut is not None:
                    dt = now - lut
                    if dt > 0:
                        self.brightness_time_sum[r][c] += self.brightness_values[r][c] * dt
                    # leave last_update_time as-is (we will reset after save)
        # Now send an all-zero matrix to Arduino
        self.brightness_values = [[0 for _ in range(self.cols)] for _ in range(self.rows)]
        self.send_matrix()

        # Save session log (uses self.start_datetime and brightness_time_sum)
        try:
            self.save_session_data()
        except Exception as e:
            print("Error saving session data:", e)
            QMessageBox.critical(self, "Save Error", f"Could not save session data:\n{e}")

        # Clear session timing state
        self.instrument_on = False
        self.start_time = None
        self.start_datetime = None
        self.last_update_time = [[None for _ in range(self.cols)] for _ in range(self.rows)]
        # Update UI to show zeros
        self._update_onoff_style()
        self._refresh_all_buttons()

    def _update_onoff_style(self):
        if self.instrument_on:
            self.onoff_btn.setText("Turn OFF")
            self.onoff_btn.setStyleSheet("background-color: green; color: white; font-weight: bold; font-size: 16px;")
        else:
            self.onoff_btn.setText("Turn ON")
            self.onoff_btn.setStyleSheet("background-color: red; color: white; font-weight: bold; font-size: 16px;")

    # -------------------------
    # Periodic timer tick (1 s)
    # - accumulates (brightness * dt) into brightness_time_sum
    # - refreshes UI labels
    # -------------------------
    def _timer_tick(self):
        if not self.instrument_on or self.start_time is None:
            # show zero or static runtime
            self.runtime_label.setText("Running time: 0s")
            return

        now = time.time()
        # accumulate dt * brightness for each LED since its last_update_time
        for r in range(self.rows):
            for c in range(self.cols):
                lut = self.last_update_time[r][c]
                if lut is not None:
                    dt = now - lut
                    if dt > 0:
                        self.brightness_time_sum[r][c] += self.brightness_values[r][c] * dt
                        self.last_update_time[r][c] = now

        # update running time label
        elapsed = int(now - self.start_time)
        hrs, rem = divmod(elapsed, 3600)
        mins, secs = divmod(rem, 60)
        self.runtime_label.setText(f"Running time: {hrs:02}:{mins:02}:{secs:02}")

        # refresh labels (averages updated)
        self._refresh_all_buttons()

    # -------------------------
    # Send matrix to Arduino (column-by-column) or print if simulation
    # -------------------------
    def send_matrix(self):
        # build transmission: start char '<', then values column-by-column, comma separated, then '>'
        parts = []
        for col in range(self.cols):
            for row in range(self.rows):
                parts.append(str(self.brightness_values[row][col]))
        transmission = "<" + ",".join(parts) + ">"
        if self.simulation_mode or (self.serial_conn is None):
            print("[SIM] Would send:", transmission)
            return

        if self.serial_conn and getattr(self.serial_conn, "is_open", False):
            try:
                self.serial_conn.write(transmission.encode())
            except Exception as e:
                print("Serial write failed:", e)
                QMessageBox.critical(self, "Serial Error", f"Error sending data:\n{e}")
        else:
            print("Serial not open; cannot send matrix.")

    # -------------------------
    # Average calculation used by labels and save
    # (includes current partial interval)
    # -------------------------
    def _calculate_average(self, r, c):
        # If there was no session, return 0
        if self.start_time is None and not self.instrument_on:
            return 0.0

        # determine session end time for computing average:
        # if instrument_on, use now; if instrument just turned off but start_time still set, use now too
        now = time.time()
        total_time = now - self.start_time if self.start_time is not None else 0.0
        if total_time <= 0:
            return float(self.brightness_values[r][c])

        # weighted sum = accumulated + current partial contribution
        accumulated = self.brightness_time_sum[r][c]
        lut = self.last_update_time[r][c]
        current = self.brightness_values[r][c]
        partial = 0.0
        if lut is not None:
            partial = (now - lut) * current
        weighted = accumulated + partial
        return weighted / total_time

    # expose a similarly named method used in older code
    def calculate_average(self, row, col):
        return self._calculate_average(row, col)

    # -------------------------
    # Refresh all button texts & styles
    # -------------------------
    def _refresh_all_buttons(self):
        for r in range(self.rows):
            for c in range(self.cols):
                btn = self.buttons[r][c]
                btn.setText(self._button_label_text(r, c))
                btn.setStyleSheet(self._button_style(self.brightness_values[r][c]))

    # -------------------------
    # Save session data (user-provided format), with safe fallback
    # -------------------------
    def save_session_data(self):
        """Append session's average brightness to a user-selected file (or fallback)."""
        end_datetime = datetime.datetime.now()
        total_time = datetime.datetime.now() - self.start_datetime if self.start_datetime else datetime.timedelta(0)

        file_name = self.ask_save_file()
        # if user canceled the dialog, fall back to instrument_log.txt in cwd
        if not file_name:
            file_name = os.path.abspath("instrument_log.txt")

        with open(file_name, "a") as f:
            f.write("=" * 50 + "\n")
            f.write(f"Session Start: {self.start_datetime}\n")
            f.write(f"Session End:   {end_datetime}\n")
            f.write(f"Total Session Time: {total_time.total_seconds():.2f} seconds\n")
            f.write("Average Brightness per LED (%):\n")
            # calculate average using the saved sums (include current partial if still applicable)
            for r in range(self.rows):
                for c in range(self.cols):
                    avg = self._calculate_average(r, c)
                    f.write(f"R{r+1}C{c+1}: {avg:.2f}%\n")
            f.write("=" * 50 + "\n\n")

        print("Session data appended to " + file_name)

    def ask_save_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Select log file to save session data",
            "instrument_log.txt",
            "Text Files (*.txt);;All Files (*)",
            options=QFileDialog.DontConfirmOverwrite
        )
        return filename  # may be '' if canceled

# -------------------------
# Program entry
# -------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = InstrumentUI()
    win.show()
    sys.exit(app.exec_())
