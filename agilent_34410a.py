"""
Agilent 34410A Digital Multimeter - GPIB Communication Interface
Complete control of all measurement functions via SCPI commands.
"""

import pyvisa
import time
import csv
import os
from datetime import datetime


class Agilent34410A:
    """Driver for the Agilent 34410A 6.5-digit Digital Multimeter."""

    def __init__(self, gpib_address=22, gpib_board=0):
        """
        Initialize connection to the 34410A.

        Args:
            gpib_address: GPIB address of the instrument (default 22)
            gpib_board: GPIB board/interface number (default 0)
        """
        self.resource_string = f"GPIB{gpib_board}::{gpib_address}::INSTR"
        self.rm = pyvisa.ResourceManager()
        self.inst = None

    def connect(self):
        """Open connection to the instrument."""
        try:
            self.inst = self.rm.open_resource(self.resource_string)
            self.inst.timeout = 10000  # 10 second timeout
            self.inst.read_termination = '\n'
            self.inst.write_termination = '\n'
            idn = self.identify()
            print(f"Connected: {idn}")
            return idn
        except pyvisa.errors.VisaIOError as e:
            print(f"Connection failed: {e}")
            print(f"Tried resource: {self.resource_string}")
            print("\nAvailable resources:")
            for r in self.rm.list_resources():
                print(f"  {r}")
            raise

    def disconnect(self):
        """Close connection to the instrument."""
        if self.inst:
            self.inst.close()
            self.inst = None
            print("Disconnected.")

    def write(self, command):
        """Send a SCPI command."""
        self.inst.write(command)

    def query(self, command):
        """Send a SCPI command and return the response."""
        return self.inst.query(command).strip()

    def query_float(self, command):
        """Send a SCPI command and return the response as a float."""
        return float(self.query(command))

    # ─── Identification & System ─────────────────────────────────────

    def identify(self):
        """Query instrument identification (*IDN?)."""
        return self.query("*IDN?")

    def reset(self):
        """Reset instrument to factory defaults (*RST)."""
        self.write("*RST")
        time.sleep(1)

    def clear_status(self):
        """Clear status registers (*CLS)."""
        self.write("*CLS")

    def self_test(self):
        """Run self-test. Returns 0 if passed."""
        result = self.query("*TST?")
        return int(result)

    def get_error(self):
        """Read and clear one error from the error queue."""
        return self.query("SYST:ERR?")

    def get_all_errors(self):
        """Read all errors from the error queue."""
        errors = []
        while True:
            err = self.get_error()
            if err.startswith("+0") or err.startswith("0"):
                break
            errors.append(err)
        return errors

    def beep(self):
        """Sound the instrument beeper."""
        self.write("SYST:BEEP")

    def get_scpi_version(self):
        """Get SCPI version."""
        return self.query("SYST:VERS?")

    def set_display_text(self, text):
        """Display custom text on the front panel (max 12 chars)."""
        self.write(f'DISP:WIND:TEXT "{text[:12]}"')

    def clear_display_text(self):
        """Clear custom display text and return to normal."""
        self.write("DISP:WIND:TEXT:CLE")

    def set_display_enabled(self, enabled=True):
        """Enable or disable the front panel display."""
        state = "ON" if enabled else "OFF"
        self.write(f"DISP {state}")

    # ─── DC Voltage Measurement ─────────────────────────────────────

    def measure_dcv(self, range_val="AUTO", resolution=None):
        """
        Measure DC voltage.

        Args:
            range_val: Range in volts (0.1, 1, 10, 100, 1000) or "AUTO"
            resolution: Resolution in volts (e.g., 0.0001) or None for default
        """
        cmd = "MEAS:VOLT:DC?"
        if range_val != "AUTO":
            cmd += f" {range_val}"
            if resolution:
                cmd += f",{resolution}"
        return self.query_float(cmd)

    def configure_dcv(self, range_val="AUTO", resolution=None):
        """Configure for DC voltage measurement without triggering."""
        if range_val == "AUTO":
            self.write("CONF:VOLT:DC AUTO")
        else:
            cmd = f"CONF:VOLT:DC {range_val}"
            if resolution:
                cmd += f",{resolution}"
            self.write(cmd)

    def set_dcv_range(self, range_val):
        """Set DC voltage range (0.1, 1, 10, 100, 1000 V)."""
        self.write(f"VOLT:DC:RANG {range_val}")

    def set_dcv_range_auto(self, enabled=True):
        """Enable/disable DC voltage autorange."""
        state = "ON" if enabled else "OFF"
        self.write(f"VOLT:DC:RANG:AUTO {state}")

    def set_dcv_nplc(self, nplc):
        """Set DC voltage integration time in PLCs (0.006, 0.02, 0.06, 0.2, 1, 2, 10, 100)."""
        self.write(f"VOLT:DC:NPLC {nplc}")

    def get_dcv_nplc(self):
        """Get DC voltage integration time in PLCs."""
        return self.query_float("VOLT:DC:NPLC?")

    def set_dcv_impedance_auto(self, enabled=True):
        """Enable/disable automatic input impedance (>10G for 0.1/1/10V ranges)."""
        state = "ON" if enabled else "OFF"
        self.write(f"VOLT:DC:IMP:AUTO {state}")

    def set_dcv_null(self, enabled=True, value=None):
        """Enable null (zero) function for DC voltage."""
        state = "ON" if enabled else "OFF"
        self.write(f"VOLT:DC:NULL {state}")
        if value is not None:
            self.write(f"VOLT:DC:NULL:VAL {value}")

    # ─── AC Voltage Measurement ─────────────────────────────────────

    def measure_acv(self, range_val="AUTO", resolution=None):
        """
        Measure AC voltage (True RMS).

        Args:
            range_val: Range in volts (0.1, 1, 10, 100, 750) or "AUTO"
            resolution: Resolution in volts or None for default
        """
        cmd = "MEAS:VOLT:AC?"
        if range_val != "AUTO":
            cmd += f" {range_val}"
            if resolution:
                cmd += f",{resolution}"
        return self.query_float(cmd)

    def configure_acv(self, range_val="AUTO", resolution=None):
        """Configure for AC voltage measurement without triggering."""
        if range_val == "AUTO":
            self.write("CONF:VOLT:AC AUTO")
        else:
            cmd = f"CONF:VOLT:AC {range_val}"
            if resolution:
                cmd += f",{resolution}"
            self.write(cmd)

    def set_acv_range(self, range_val):
        """Set AC voltage range (0.1, 1, 10, 100, 750 V)."""
        self.write(f"VOLT:AC:RANG {range_val}")

    def set_acv_range_auto(self, enabled=True):
        """Enable/disable AC voltage autorange."""
        state = "ON" if enabled else "OFF"
        self.write(f"VOLT:AC:RANG:AUTO {state}")

    def set_acv_bandwidth(self, bandwidth):
        """Set AC voltage bandwidth (3, 20, 200 Hz)."""
        self.write(f"VOLT:AC:BAND {bandwidth}")

    def set_acv_null(self, enabled=True, value=None):
        """Enable null function for AC voltage."""
        state = "ON" if enabled else "OFF"
        self.write(f"VOLT:AC:NULL {state}")
        if value is not None:
            self.write(f"VOLT:AC:NULL:VAL {value}")

    # ─── DC Current Measurement ──────────────────────────────────────

    def measure_dci(self, range_val="AUTO", resolution=None):
        """
        Measure DC current.

        Args:
            range_val: Range in amps (0.0001, 0.001, 0.01, 0.1, 1, 3) or "AUTO"
            resolution: Resolution in amps or None for default
        """
        cmd = "MEAS:CURR:DC?"
        if range_val != "AUTO":
            cmd += f" {range_val}"
            if resolution:
                cmd += f",{resolution}"
        return self.query_float(cmd)

    def configure_dci(self, range_val="AUTO", resolution=None):
        """Configure for DC current measurement without triggering."""
        if range_val == "AUTO":
            self.write("CONF:CURR:DC AUTO")
        else:
            cmd = f"CONF:CURR:DC {range_val}"
            if resolution:
                cmd += f",{resolution}"
            self.write(cmd)

    def set_dci_range(self, range_val):
        """Set DC current range (0.0001, 0.001, 0.01, 0.1, 1, 3 A)."""
        self.write(f"CURR:DC:RANG {range_val}")

    def set_dci_range_auto(self, enabled=True):
        """Enable/disable DC current autorange."""
        state = "ON" if enabled else "OFF"
        self.write(f"CURR:DC:RANG:AUTO {state}")

    def set_dci_nplc(self, nplc):
        """Set DC current integration time in PLCs."""
        self.write(f"CURR:DC:NPLC {nplc}")

    def set_dci_null(self, enabled=True, value=None):
        """Enable null function for DC current."""
        state = "ON" if enabled else "OFF"
        self.write(f"CURR:DC:NULL {state}")
        if value is not None:
            self.write(f"CURR:DC:NULL:VAL {value}")

    # ─── AC Current Measurement ──────────────────────────────────────

    def measure_aci(self, range_val="AUTO", resolution=None):
        """
        Measure AC current (True RMS).

        Args:
            range_val: Range in amps (0.0001, 0.001, 0.01, 0.1, 1, 3) or "AUTO"
            resolution: Resolution in amps or None for default
        """
        cmd = "MEAS:CURR:AC?"
        if range_val != "AUTO":
            cmd += f" {range_val}"
            if resolution:
                cmd += f",{resolution}"
        return self.query_float(cmd)

    def configure_aci(self, range_val="AUTO", resolution=None):
        """Configure for AC current measurement without triggering."""
        if range_val == "AUTO":
            self.write("CONF:CURR:AC AUTO")
        else:
            cmd = f"CONF:CURR:AC {range_val}"
            if resolution:
                cmd += f",{resolution}"
            self.write(cmd)

    def set_aci_range(self, range_val):
        """Set AC current range."""
        self.write(f"CURR:AC:RANG {range_val}")

    def set_aci_range_auto(self, enabled=True):
        """Enable/disable AC current autorange."""
        state = "ON" if enabled else "OFF"
        self.write(f"CURR:AC:RANG:AUTO {state}")

    def set_aci_bandwidth(self, bandwidth):
        """Set AC current bandwidth (3, 20, 200 Hz)."""
        self.write(f"CURR:AC:BAND {bandwidth}")

    def set_aci_null(self, enabled=True, value=None):
        """Enable null function for AC current."""
        state = "ON" if enabled else "OFF"
        self.write(f"CURR:AC:NULL {state}")
        if value is not None:
            self.write(f"CURR:AC:NULL:VAL {value}")

    # ─── 2-Wire Resistance Measurement ──────────────────────────────

    def measure_resistance_2w(self, range_val="AUTO", resolution=None):
        """
        Measure 2-wire resistance.

        Args:
            range_val: Range in ohms (100, 1e3, 10e3, 100e3, 1e6, 10e6, 100e6, 1e9) or "AUTO"
            resolution: Resolution in ohms or None for default
        """
        cmd = "MEAS:RES?"
        if range_val != "AUTO":
            cmd += f" {range_val}"
            if resolution:
                cmd += f",{resolution}"
        return self.query_float(cmd)

    def configure_resistance_2w(self, range_val="AUTO", resolution=None):
        """Configure for 2-wire resistance measurement."""
        if range_val == "AUTO":
            self.write("CONF:RES AUTO")
        else:
            cmd = f"CONF:RES {range_val}"
            if resolution:
                cmd += f",{resolution}"
            self.write(cmd)

    def set_resistance_2w_range(self, range_val):
        """Set 2-wire resistance range."""
        self.write(f"RES:RANG {range_val}")

    def set_resistance_2w_range_auto(self, enabled=True):
        """Enable/disable 2-wire resistance autorange."""
        state = "ON" if enabled else "OFF"
        self.write(f"RES:RANG:AUTO {state}")

    def set_resistance_2w_nplc(self, nplc):
        """Set 2-wire resistance integration time in PLCs."""
        self.write(f"RES:NPLC {nplc}")

    def set_resistance_2w_null(self, enabled=True, value=None):
        """Enable null function for 2-wire resistance."""
        state = "ON" if enabled else "OFF"
        self.write(f"RES:NULL {state}")
        if value is not None:
            self.write(f"RES:NULL:VAL {value}")

    # ─── 4-Wire Resistance Measurement ──────────────────────────────

    def measure_resistance_4w(self, range_val="AUTO", resolution=None):
        """
        Measure 4-wire resistance.

        Args:
            range_val: Range in ohms (100, 1e3, 10e3, 100e3, 1e6, 10e6, 100e6, 1e9) or "AUTO"
            resolution: Resolution in ohms or None for default
        """
        cmd = "MEAS:FRES?"
        if range_val != "AUTO":
            cmd += f" {range_val}"
            if resolution:
                cmd += f",{resolution}"
        return self.query_float(cmd)

    def configure_resistance_4w(self, range_val="AUTO", resolution=None):
        """Configure for 4-wire resistance measurement."""
        if range_val == "AUTO":
            self.write("CONF:FRES AUTO")
        else:
            cmd = f"CONF:FRES {range_val}"
            if resolution:
                cmd += f",{resolution}"
            self.write(cmd)

    def set_resistance_4w_range(self, range_val):
        """Set 4-wire resistance range."""
        self.write(f"FRES:RANG {range_val}")

    def set_resistance_4w_range_auto(self, enabled=True):
        """Enable/disable 4-wire resistance autorange."""
        state = "ON" if enabled else "OFF"
        self.write(f"FRES:RANG:AUTO {state}")

    def set_resistance_4w_nplc(self, nplc):
        """Set 4-wire resistance integration time in PLCs."""
        self.write(f"FRES:NPLC {nplc}")

    def set_resistance_4w_null(self, enabled=True, value=None):
        """Enable null function for 4-wire resistance."""
        state = "ON" if enabled else "OFF"
        self.write(f"FRES:NULL {state}")
        if value is not None:
            self.write(f"FRES:NULL:VAL {value}")

    # ─── Frequency / Period Measurement ──────────────────────────────

    def measure_frequency(self, range_val="AUTO", resolution=None):
        """
        Measure frequency.

        Args:
            range_val: Expected voltage range or "AUTO"
            resolution: Resolution in Hz or None for default
        """
        cmd = "MEAS:FREQ?"
        if range_val != "AUTO":
            cmd += f" {range_val}"
            if resolution:
                cmd += f",{resolution}"
        return self.query_float(cmd)

    def configure_frequency(self, range_val="AUTO", resolution=None):
        """Configure for frequency measurement."""
        if range_val == "AUTO":
            self.write("CONF:FREQ")
        else:
            cmd = f"CONF:FREQ {range_val}"
            if resolution:
                cmd += f",{resolution}"
            self.write(cmd)

    def set_frequency_aperture(self, aperture):
        """Set frequency gate time/aperture (0.01, 0.1, 1 seconds)."""
        self.write(f"FREQ:APER {aperture}")

    def set_frequency_null(self, enabled=True, value=None):
        """Enable null function for frequency."""
        state = "ON" if enabled else "OFF"
        self.write(f"FREQ:NULL {state}")
        if value is not None:
            self.write(f"FREQ:NULL:VAL {value}")

    def measure_period(self, range_val="AUTO", resolution=None):
        """Measure period."""
        cmd = "MEAS:PER?"
        if range_val != "AUTO":
            cmd += f" {range_val}"
            if resolution:
                cmd += f",{resolution}"
        return self.query_float(cmd)

    def configure_period(self, range_val="AUTO", resolution=None):
        """Configure for period measurement."""
        if range_val == "AUTO":
            self.write("CONF:PER")
        else:
            cmd = f"CONF:PER {range_val}"
            if resolution:
                cmd += f",{resolution}"
            self.write(cmd)

    # ─── Continuity / Diode ──────────────────────────────────────────

    def measure_continuity(self):
        """Measure continuity (fixed 1 kOhm range, threshold 10 ohms)."""
        return self.query_float("MEAS:CONT?")

    def configure_continuity(self):
        """Configure for continuity measurement."""
        self.write("CONF:CONT")

    def measure_diode(self):
        """Measure diode forward voltage (1 mA test current, 5V range)."""
        return self.query_float("MEAS:DIOD?")

    def configure_diode(self):
        """Configure for diode test."""
        self.write("CONF:DIOD")

    # ─── Temperature Measurement ─────────────────────────────────────

    def measure_temperature(self):
        """Measure temperature (requires temperature probe)."""
        return self.query_float("MEAS:TEMP?")

    def configure_temperature(self, probe_type="RTD", sub_type="PT100"):
        """
        Configure temperature measurement.

        Args:
            probe_type: "RTD", "THER" (thermistor), or "FRTD" (4-wire RTD)
            sub_type: For RTD: "PT100". For THER: "2252", "5000", "10000"
        """
        self.write(f"CONF:TEMP {probe_type},{sub_type}")

    def set_temperature_units(self, unit="C"):
        """Set temperature units: 'C' (Celsius), 'F' (Fahrenheit), 'K' (Kelvin)."""
        self.write(f"UNIT:TEMP {unit}")

    def get_temperature_units(self):
        """Get current temperature units."""
        return self.query("UNIT:TEMP?")

    # ─── Capacitance Measurement ─────────────────────────────────────

    def measure_capacitance(self, range_val="AUTO"):
        """
        Measure capacitance.

        Args:
            range_val: Range in farads (1e-9, 10e-9, 100e-9, 1e-6, 10e-6, 100e-6) or "AUTO"
        """
        cmd = "MEAS:CAP?"
        if range_val != "AUTO":
            cmd += f" {range_val}"
        return self.query_float(cmd)

    def configure_capacitance(self, range_val="AUTO"):
        """Configure for capacitance measurement."""
        if range_val == "AUTO":
            self.write("CONF:CAP")
        else:
            self.write(f"CONF:CAP {range_val}")

    def set_capacitance_range(self, range_val):
        """Set capacitance range."""
        self.write(f"CAP:RANG {range_val}")

    def set_capacitance_range_auto(self, enabled=True):
        """Enable/disable capacitance autorange."""
        state = "ON" if enabled else "OFF"
        self.write(f"CAP:RANG:AUTO {state}")

    def set_capacitance_null(self, enabled=True, value=None):
        """Enable null function for capacitance."""
        state = "ON" if enabled else "OFF"
        self.write(f"CAP:NULL {state}")
        if value is not None:
            self.write(f"CAP:NULL:VAL {value}")

    # ─── Trigger System ─────────────────────────────────────────────

    def set_trigger_source(self, source="IMM"):
        """Set trigger source: IMM (immediate), BUS, EXT (external), INT (internal)."""
        self.write(f"TRIG:SOUR {source}")

    def get_trigger_source(self):
        """Get current trigger source."""
        return self.query("TRIG:SOUR?")

    def set_trigger_delay(self, delay):
        """Set trigger delay in seconds (0 to 3600)."""
        self.write(f"TRIG:DEL {delay}")

    def set_trigger_delay_auto(self, enabled=True):
        """Enable/disable automatic trigger delay."""
        state = "ON" if enabled else "OFF"
        self.write(f"TRIG:DEL:AUTO {state}")

    def set_trigger_count(self, count):
        """Set number of triggers to accept (1 to 1000000, or INF)."""
        self.write(f"TRIG:COUN {count}")

    def set_sample_count(self, count):
        """Set number of samples per trigger (1 to 1000000)."""
        self.write(f"SAMP:COUN {count}")

    def set_sample_timer(self, interval):
        """Set sample timer interval in seconds (min ~20 us)."""
        self.write(f"SAMP:TIM {interval}")

    def set_sample_source(self, source="IMM"):
        """Set sample source: IMM (immediate) or TIM (timer)."""
        self.write(f"SAMP:SOUR {source}")

    def trigger(self):
        """Send a software trigger (when trigger source is BUS)."""
        self.write("*TRG")

    def initiate(self):
        """Initiate a measurement (move from idle to wait-for-trigger)."""
        self.write("INIT")

    def fetch(self):
        """Fetch the last measurement result(s)."""
        return self.query("FETC?")

    def read(self):
        """Initiate and fetch a measurement in one step."""
        return self.query_float("READ?")

    # ─── Math / Statistics ───────────────────────────────────────────

    def set_math_function(self, func):
        """Set math function: NULL, DB, DBM, AVER, LIM."""
        self.write(f"CALC:FUNC {func}")

    def set_math_enabled(self, enabled=True):
        """Enable/disable math function."""
        state = "ON" if enabled else "OFF"
        self.write(f"CALC:STAT {state}")

    def set_math_db_reference(self, reference):
        """Set dB reference value."""
        self.write(f"CALC:DB:REF {reference}")

    def set_math_dbm_reference(self, impedance):
        """Set dBm reference impedance (1 to 9999 ohms)."""
        self.write(f"CALC:DBM:REF {impedance}")

    def set_math_limit_lower(self, value):
        """Set lower limit for limit test."""
        self.write(f"CALC:LIM:LOW {value}")

    def set_math_limit_upper(self, value):
        """Set upper limit for limit test."""
        self.write(f"CALC:LIM:UPP {value}")

    def get_math_average(self):
        """Get the average of readings (when AVER math is enabled)."""
        return self.query_float("CALC:AVER:AVER?")

    def get_math_count(self):
        """Get the count of readings in statistics."""
        return int(self.query_float("CALC:AVER:COUN?"))

    def get_math_max(self):
        """Get the maximum reading."""
        return self.query_float("CALC:AVER:MAX?")

    def get_math_min(self):
        """Get the minimum reading."""
        return self.query_float("CALC:AVER:MIN?")

    def get_math_ptp(self):
        """Get peak-to-peak (max - min)."""
        return self.query_float("CALC:AVER:PTP?")

    def get_math_sdev(self):
        """Get standard deviation of readings."""
        return self.query_float("CALC:AVER:SDEV?")

    def clear_math_statistics(self):
        """Clear/reset all statistics."""
        self.write("CALC:AVER:CLE")

    # ─── Data Logging / Memory ───────────────────────────────────────

    def get_data_count(self):
        """Get number of readings stored in memory."""
        return int(self.query_float("DATA:COUN?"))

    def get_data_points(self, count=None):
        """
        Retrieve readings from memory.

        Args:
            count: Number of readings to retrieve, or None for all
        """
        if count:
            data = self.query(f"DATA:REM? {count}")
        else:
            data = self.query("DATA:REM?")
        return [float(x) for x in data.split(',')]

    def clear_data(self):
        """Clear all readings from memory."""
        self.write("DATA:DEL NVMEM")

    # ─── Convenience / High-Level Methods ────────────────────────────

    def quick_measure(self, function="DCV"):
        """
        Quick single measurement with auto settings.

        Args:
            function: DCV, ACV, DCI, ACI, RES2W, RES4W, FREQ, PER, CONT, DIODE, TEMP, CAP

        Returns:
            Measurement value as float
        """
        funcs = {
            "DCV":   self.measure_dcv,
            "ACV":   self.measure_acv,
            "DCI":   self.measure_dci,
            "ACI":   self.measure_aci,
            "RES2W": self.measure_resistance_2w,
            "RES4W": self.measure_resistance_4w,
            "FREQ":  self.measure_frequency,
            "PER":   self.measure_period,
            "CONT":  self.measure_continuity,
            "DIODE": self.measure_diode,
            "TEMP":  self.measure_temperature,
            "CAP":   self.measure_capacitance,
        }
        func = funcs.get(function.upper())
        if func is None:
            raise ValueError(f"Unknown function '{function}'. Valid: {list(funcs.keys())}")
        return func()

    def data_log(self, function="DCV", interval=1.0, count=10, filename=None):
        """
        Log measurements to console and optionally to CSV file.

        Args:
            function: Measurement function (DCV, ACV, DCI, ACI, RES2W, etc.)
            interval: Time between measurements in seconds
            count: Number of measurements to take
            filename: Optional CSV filename to save data
        """
        readings = []
        print(f"\nData Logging: {function}, {count} readings, {interval}s interval")
        print("-" * 50)

        for i in range(count):
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            value = self.quick_measure(function)
            readings.append((timestamp, value))
            print(f"  [{i+1}/{count}] {timestamp}  {value:+.8E}")

            if i < count - 1:
                time.sleep(interval)

        if filename:
            filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            with open(filepath, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Timestamp", f"{function} Reading"])
                writer.writerows(readings)
            print(f"\nData saved to: {filepath}")

        return readings

    def multi_read(self, count=10):
        """
        Take multiple readings using the instrument's internal sample counter.
        Faster than data_log() for burst measurements.

        Args:
            count: Number of samples (1-1000000)

        Returns:
            List of float readings
        """
        self.set_sample_count(count)
        self.set_trigger_count(1)
        self.write("INIT")
        self.write("*WAI")
        data = self.query("FETC?")
        values = [float(x) for x in data.split(',')]
        self.set_sample_count(1)  # reset
        return values

    def get_configuration(self):
        """Get current measurement configuration summary."""
        config = {}
        config["function"] = self.query("CONF?")
        config["trigger_source"] = self.get_trigger_source()
        config["trigger_delay"] = self.query("TRIG:DEL?")
        config["sample_count"] = self.query("SAMP:COUN?")
        return config

    def preset_high_speed(self):
        """Configure for maximum measurement speed (reduced accuracy)."""
        self.write("CONF:VOLT:DC AUTO")
        self.set_dcv_nplc(0.006)
        self.write("ZERO:AUTO OFF")
        self.set_trigger_delay(0)
        self.set_display_enabled(False)

    def preset_high_accuracy(self):
        """Configure for maximum accuracy (slower)."""
        self.write("CONF:VOLT:DC AUTO")
        self.set_dcv_nplc(100)
        self.set_dcv_impedance_auto(True)
        self.write("ZERO:AUTO ON")
        self.set_trigger_delay_auto(True)
        self.set_display_enabled(True)


def list_resources():
    """List all available VISA resources."""
    rm = pyvisa.ResourceManager()
    resources = rm.list_resources()
    if resources:
        print("Available VISA resources:")
        for r in resources:
            print(f"  {r}")
    else:
        print("No VISA resources found.")
        print("Check that:")
        print("  1. NI-VISA or Keysight IO Libraries are installed")
        print("  2. GPIB adapter is connected and drivers are loaded")
        print("  3. Instrument is powered on")
    return resources


# ─── Interactive Menu ────────────────────────────────────────────────

def interactive_menu():
    """Run an interactive menu for the 34410A."""
    print("=" * 60)
    print("  Agilent 34410A Digital Multimeter - Control Interface")
    print("=" * 60)

    # First list available resources
    resources = list_resources()
    print()

    # Get GPIB address
    addr = input("Enter GPIB address [22]: ").strip()
    gpib_address = int(addr) if addr else 22

    dmm = Agilent34410A(gpib_address=gpib_address)

    try:
        dmm.connect()
    except Exception as e:
        print(f"\nFailed to connect: {e}")
        return

    print()

    while True:
        print("\n" + "=" * 60)
        print("  MEASUREMENT FUNCTIONS")
        print("=" * 60)
        print("  1.  DC Voltage          7.  Frequency")
        print("  2.  AC Voltage          8.  Period")
        print("  3.  DC Current          9.  Continuity")
        print("  4.  AC Current          10. Diode Test")
        print("  5.  2-Wire Resistance   11. Temperature")
        print("  6.  4-Wire Resistance   12. Capacitance")
        print()
        print("  DATA & SETTINGS")
        print("  13. Data Log (multiple readings to CSV)")
        print("  14. Burst Read (fast multiple readings)")
        print("  15. Show Configuration")
        print("  16. Set Integration Time (NPLC)")
        print("  17. Preset: High Speed")
        print("  18. Preset: High Accuracy")
        print("  19. Display Custom Text")
        print()
        print("  SYSTEM")
        print("  20. Identify Instrument")
        print("  21. Self Test")
        print("  22. Reset to Defaults")
        print("  23. Check Errors")
        print("  24. Beep")
        print("   0. Exit")
        print("-" * 60)

        choice = input("Select [0-24]: ").strip()

        try:
            if choice == "0":
                break
            elif choice == "1":
                val = dmm.measure_dcv()
                print(f"\n  DC Voltage: {val:+.8E} V")
            elif choice == "2":
                val = dmm.measure_acv()
                print(f"\n  AC Voltage: {val:+.8E} V")
            elif choice == "3":
                val = dmm.measure_dci()
                print(f"\n  DC Current: {val:+.8E} A")
            elif choice == "4":
                val = dmm.measure_aci()
                print(f"\n  AC Current: {val:+.8E} A")
            elif choice == "5":
                val = dmm.measure_resistance_2w()
                print(f"\n  2-Wire Resistance: {val:+.8E} Ohm")
            elif choice == "6":
                val = dmm.measure_resistance_4w()
                print(f"\n  4-Wire Resistance: {val:+.8E} Ohm")
            elif choice == "7":
                val = dmm.measure_frequency()
                print(f"\n  Frequency: {val:+.8E} Hz")
            elif choice == "8":
                val = dmm.measure_period()
                print(f"\n  Period: {val:+.8E} s")
            elif choice == "9":
                val = dmm.measure_continuity()
                print(f"\n  Continuity: {val:+.8E} Ohm")
            elif choice == "10":
                val = dmm.measure_diode()
                print(f"\n  Diode: {val:+.8E} V")
            elif choice == "11":
                val = dmm.measure_temperature()
                print(f"\n  Temperature: {val:.4f} {dmm.get_temperature_units()}")
            elif choice == "12":
                val = dmm.measure_capacitance()
                print(f"\n  Capacitance: {val:+.8E} F")
            elif choice == "13":
                func = input("  Function (DCV/ACV/DCI/ACI/RES2W/RES4W/FREQ/CAP) [DCV]: ").strip() or "DCV"
                intv = input("  Interval in seconds [1.0]: ").strip()
                interval = float(intv) if intv else 1.0
                cnt = input("  Number of readings [10]: ").strip()
                count = int(cnt) if cnt else 10
                fname = input("  CSV filename (blank=none): ").strip() or None
                dmm.data_log(func, interval, count, fname)
            elif choice == "14":
                cnt = input("  Number of readings [100]: ").strip()
                count = int(cnt) if cnt else 100
                readings = dmm.multi_read(count)
                print(f"\n  Captured {len(readings)} readings")
                print(f"  Min: {min(readings):+.8E}")
                print(f"  Max: {max(readings):+.8E}")
                avg = sum(readings) / len(readings)
                print(f"  Avg: {avg:+.8E}")
            elif choice == "15":
                config = dmm.get_configuration()
                print("\n  Current Configuration:")
                for k, v in config.items():
                    print(f"    {k}: {v}")
            elif choice == "16":
                nplc = input("  NPLC (0.006/0.02/0.06/0.2/1/2/10/100) [10]: ").strip()
                nplc = float(nplc) if nplc else 10
                dmm.set_dcv_nplc(nplc)
                print(f"  NPLC set to {nplc}")
            elif choice == "17":
                dmm.preset_high_speed()
                print("  High speed preset applied.")
            elif choice == "18":
                dmm.preset_high_accuracy()
                print("  High accuracy preset applied.")
            elif choice == "19":
                text = input("  Text (max 12 chars): ").strip()
                if text:
                    dmm.set_display_text(text)
                else:
                    dmm.clear_display_text()
                    print("  Display cleared.")
            elif choice == "20":
                print(f"\n  {dmm.identify()}")
            elif choice == "21":
                print("  Running self-test (may take ~15 seconds)...")
                result = dmm.self_test()
                print(f"  Self-test {'PASSED' if result == 0 else 'FAILED'} (code: {result})")
            elif choice == "22":
                dmm.reset()
                print("  Instrument reset to defaults.")
            elif choice == "23":
                errors = dmm.get_all_errors()
                if errors:
                    print("  Errors:")
                    for e in errors:
                        print(f"    {e}")
                else:
                    print("  No errors.")
            elif choice == "24":
                dmm.beep()
            else:
                print("  Invalid selection.")

        except Exception as e:
            print(f"\n  Error: {e}")

    dmm.disconnect()
    print("Goodbye!")


if __name__ == "__main__":
    interactive_menu()
