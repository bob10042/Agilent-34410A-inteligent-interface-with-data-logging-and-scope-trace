"""
Agilent 34410A Digital Multimeter - Full GUI Control Interface
Mirrors all front panel functions with live data streaming and CSV recording.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pyvisa
import threading
import time
import csv
import os
from datetime import datetime


class Agilent34410A:
    """Low-level SCPI driver for the Agilent 34410A."""

    def __init__(self):
        self.rm = pyvisa.ResourceManager()
        self.inst = None
        self.connected = False

    def list_resources(self):
        return list(self.rm.list_resources())

    def connect(self, resource_string):
        self.inst = self.rm.open_resource(resource_string)
        self.inst.timeout = 15000
        self.inst.read_termination = '\n'
        self.inst.write_termination = '\n'
        self.connected = True
        return self.query("*IDN?")

    def disconnect(self):
        if self.inst:
            self.inst.close()
            self.inst = None
        self.connected = False

    def write(self, cmd):
        if self.inst:
            self.inst.write(cmd)

    def query(self, cmd):
        if self.inst:
            return self.inst.query(cmd).strip()
        return ""

    def query_float(self, cmd):
        return float(self.query(cmd))

    def reset(self):
        self.write("*RST")
        time.sleep(1)

    def clear_status(self):
        self.write("*CLS")

    def self_test(self):
        return int(self.query("*TST?"))

    def get_error(self):
        return self.query("SYST:ERR?")

    def get_all_errors(self):
        errors = []
        while True:
            err = self.get_error()
            if err.startswith("+0") or err.startswith("0"):
                break
            errors.append(err)
        return errors

    def beep(self):
        self.write("SYST:BEEP")

    def get_terminals(self):
        return self.query("ROUT:TERM?")


class DMMApplication(tk.Tk):
    """Main GUI Application for the Agilent 34410A."""

    # Measurement function definitions
    FUNCTIONS = {
        "DC Voltage": {
            "meas_cmd": "MEAS:VOLT:DC?",
            "conf_cmd": "CONF:VOLT:DC",
            "sense_prefix": "VOLT:DC",
            "unit": "V DC",
            "ranges": ["AUTO", "0.1", "1", "10", "100", "1000"],
            "has_nplc": True,
            "has_impedance": True,
            "has_null": True,
            "has_bandwidth": False,
            "has_aperture": False,
        },
        "AC Voltage": {
            "meas_cmd": "MEAS:VOLT:AC?",
            "conf_cmd": "CONF:VOLT:AC",
            "sense_prefix": "VOLT:AC",
            "unit": "V AC",
            "ranges": ["AUTO", "0.1", "1", "10", "100", "750"],
            "has_nplc": False,
            "has_impedance": False,
            "has_null": True,
            "has_bandwidth": True,
            "has_aperture": False,
        },
        "DC Current": {
            "meas_cmd": "MEAS:CURR:DC?",
            "conf_cmd": "CONF:CURR:DC",
            "sense_prefix": "CURR:DC",
            "unit": "A DC",
            "ranges": ["AUTO", "0.0001", "0.001", "0.01", "0.1", "1", "3"],
            "has_nplc": True,
            "has_impedance": False,
            "has_null": True,
            "has_bandwidth": False,
            "has_aperture": False,
        },
        "AC Current": {
            "meas_cmd": "MEAS:CURR:AC?",
            "conf_cmd": "CONF:CURR:AC",
            "sense_prefix": "CURR:AC",
            "unit": "A AC",
            "ranges": ["AUTO", "0.0001", "0.001", "0.01", "0.1", "1", "3"],
            "has_nplc": False,
            "has_impedance": False,
            "has_null": True,
            "has_bandwidth": True,
            "has_aperture": False,
        },
        "2-Wire Resistance": {
            "meas_cmd": "MEAS:RES?",
            "conf_cmd": "CONF:RES",
            "sense_prefix": "RES",
            "unit": "\u03a9",
            "ranges": ["AUTO", "100", "1000", "10000", "100000", "1e6", "10e6", "100e6", "1e9"],
            "has_nplc": True,
            "has_impedance": False,
            "has_null": True,
            "has_bandwidth": False,
            "has_aperture": False,
        },
        "4-Wire Resistance": {
            "meas_cmd": "MEAS:FRES?",
            "conf_cmd": "CONF:FRES",
            "sense_prefix": "FRES",
            "unit": "\u03a9 4W",
            "ranges": ["AUTO", "100", "1000", "10000", "100000", "1e6", "10e6", "100e6", "1e9"],
            "has_nplc": True,
            "has_impedance": False,
            "has_null": True,
            "has_bandwidth": False,
            "has_aperture": False,
        },
        "Frequency": {
            "meas_cmd": "MEAS:FREQ?",
            "conf_cmd": "CONF:FREQ",
            "sense_prefix": "FREQ",
            "unit": "Hz",
            "ranges": ["AUTO", "0.1", "1", "10", "100", "750"],
            "has_nplc": False,
            "has_impedance": False,
            "has_null": True,
            "has_bandwidth": False,
            "has_aperture": True,
        },
        "Period": {
            "meas_cmd": "MEAS:PER?",
            "conf_cmd": "CONF:PER",
            "sense_prefix": "PER",
            "unit": "s",
            "ranges": ["AUTO", "0.1", "1", "10", "100", "750"],
            "has_nplc": False,
            "has_impedance": False,
            "has_null": True,
            "has_bandwidth": False,
            "has_aperture": True,
        },
        "Continuity": {
            "meas_cmd": "MEAS:CONT?",
            "conf_cmd": "CONF:CONT",
            "sense_prefix": "CONT",
            "unit": "\u03a9",
            "ranges": [],
            "has_nplc": False,
            "has_impedance": False,
            "has_null": False,
            "has_bandwidth": False,
            "has_aperture": False,
        },
        "Diode": {
            "meas_cmd": "MEAS:DIOD?",
            "conf_cmd": "CONF:DIOD",
            "sense_prefix": "DIOD",
            "unit": "V",
            "ranges": [],
            "has_nplc": False,
            "has_impedance": False,
            "has_null": False,
            "has_bandwidth": False,
            "has_aperture": False,
        },
        "Temperature": {
            "meas_cmd": "MEAS:TEMP?",
            "conf_cmd": "CONF:TEMP",
            "sense_prefix": "TEMP",
            "unit": "\u00b0C",
            "ranges": [],
            "has_nplc": True,
            "has_impedance": False,
            "has_null": True,
            "has_bandwidth": False,
            "has_aperture": False,
        },
        "Capacitance": {
            "meas_cmd": "MEAS:CAP?",
            "conf_cmd": "CONF:CAP",
            "sense_prefix": "CAP",
            "unit": "F",
            "ranges": ["AUTO", "1e-9", "10e-9", "100e-9", "1e-6", "10e-6", "100e-6"],
            "has_nplc": False,
            "has_impedance": False,
            "has_null": True,
            "has_bandwidth": False,
            "has_aperture": False,
        },
    }

    NPLC_VALUES = ["0.006", "0.02", "0.06", "0.2", "1", "2", "10", "100"]
    BANDWIDTH_VALUES = ["3", "20", "200"]
    APERTURE_VALUES = ["0.01", "0.1", "1"]

    def __init__(self):
        super().__init__()
        self.title("Agilent 34410A - Digital Multimeter Control")
        self.geometry("1280x900")
        self.minsize(1100, 800)
        self.configure(bg="#2b2b2b")

        self.dmm = Agilent34410A()
        self.current_function = "DC Voltage"
        self.streaming = False
        self.stream_thread = None
        self.recording = False
        self.csv_file = None
        self.csv_writer = None
        self.reading_count = 0
        self.readings_list = []
        self.stats = {"min": None, "max": None, "sum": 0, "count": 0}

        self._build_styles()
        self._build_ui()
        self._populate_resources()
        self.after(500, self._auto_connect)

    def _build_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Dark theme colors
        bg = "#2b2b2b"
        fg = "#e0e0e0"
        btn_bg = "#3c3c3c"
        btn_active = "#505050"
        accent = "#4a9eff"
        frame_bg = "#333333"

        self.style.configure("TFrame", background=bg)
        self.style.configure("Card.TFrame", background=frame_bg)
        self.style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", background=bg, foreground=fg, font=("Segoe UI", 12, "bold"))
        self.style.configure("Card.TLabel", background=frame_bg, foreground=fg, font=("Segoe UI", 10))
        self.style.configure("CardHeader.TLabel", background=frame_bg, foreground=accent, font=("Segoe UI", 10, "bold"))
        self.style.configure("TButton", background=btn_bg, foreground=fg, font=("Segoe UI", 10), padding=6)
        self.style.map("TButton", background=[("active", btn_active)])
        self.style.configure("Accent.TButton", background=accent, foreground="white", font=("Segoe UI", 10, "bold"))
        self.style.map("Accent.TButton", background=[("active", "#3a8eef")])
        self.style.configure("Stop.TButton", background="#cc4444", foreground="white", font=("Segoe UI", 10, "bold"))
        self.style.map("Stop.TButton", background=[("active", "#aa3333")])
        self.style.configure("Func.TButton", background="#3c5c3c", foreground="white", font=("Segoe UI", 9, "bold"), padding=4)
        self.style.map("Func.TButton", background=[("active", "#4a7a4a")])
        self.style.configure("FuncActive.TButton", background="#2e8b2e", foreground="white", font=("Segoe UI", 9, "bold"), padding=4)
        self.style.configure("TCombobox", fieldbackground=btn_bg, foreground=fg, background=btn_bg)
        self.style.configure("TCheckbutton", background=bg, foreground=fg, font=("Segoe UI", 10))
        self.style.configure("TLabelframe", background=frame_bg, foreground=accent, font=("Segoe UI", 10, "bold"))
        self.style.configure("TLabelframe.Label", background=frame_bg, foreground=accent, font=("Segoe UI", 10, "bold"))

        self.style.configure("Treeview", background=btn_bg, foreground=fg, fieldbackground=btn_bg,
                             font=("Consolas", 9), rowheight=22)
        self.style.configure("Treeview.Heading", background=frame_bg, foreground=accent, font=("Segoe UI", 9, "bold"))
        self.style.map("Treeview", background=[("selected", accent)])

    def _build_ui(self):
        # ─── Top: Connection Bar ─────────────────────────────────
        conn_frame = ttk.Frame(self)
        conn_frame.pack(fill="x", padx=10, pady=(10, 5))

        ttk.Label(conn_frame, text="Resource:").pack(side="left", padx=(0, 5))
        self.resource_var = tk.StringVar()
        self.resource_combo = ttk.Combobox(conn_frame, textvariable=self.resource_var, width=30, state="readonly")
        self.resource_combo.pack(side="left", padx=(0, 5))

        ttk.Button(conn_frame, text="Refresh", command=self._populate_resources).pack(side="left", padx=2)
        self.connect_btn = ttk.Button(conn_frame, text="Connect", style="Accent.TButton", command=self._toggle_connection)
        self.connect_btn.pack(side="left", padx=5)

        self.status_var = tk.StringVar(value="Disconnected")
        ttk.Label(conn_frame, textvariable=self.status_var, foreground="#ff6666").pack(side="left", padx=10)

        self.terminals_var = tk.StringVar(value="")
        ttk.Label(conn_frame, textvariable=self.terminals_var, foreground="#aaaaaa").pack(side="right", padx=5)

        self.idn_var = tk.StringVar(value="")
        ttk.Label(conn_frame, textvariable=self.idn_var, foreground="#888888", font=("Segoe UI", 9)).pack(side="right", padx=5)

        # ─── Display ─────────────────────────────────────────────
        display_frame = tk.Frame(self, bg="#0a0a0a", relief="sunken", bd=2)
        display_frame.pack(fill="x", padx=10, pady=5)

        self.display_value = tk.StringVar(value="---")
        self.display_unit = tk.StringVar(value="")
        self.display_func = tk.StringVar(value="")

        tk.Label(display_frame, textvariable=self.display_func, bg="#0a0a0a", fg="#55aa55",
                 font=("Consolas", 14), anchor="w").pack(fill="x", padx=20, pady=(10, 0))

        value_row = tk.Frame(display_frame, bg="#0a0a0a")
        value_row.pack(fill="x", padx=20, pady=(0, 10))
        tk.Label(value_row, textvariable=self.display_value, bg="#0a0a0a", fg="#00ff88",
                 font=("Consolas", 56, "bold"), anchor="e").pack(side="left", expand=True, fill="x")
        tk.Label(value_row, textvariable=self.display_unit, bg="#0a0a0a", fg="#44cc66",
                 font=("Consolas", 24), anchor="w").pack(side="right", padx=(10, 20))

        # ─── Function Buttons (mirrors front panel layout) ───────
        func_frame = ttk.Frame(self)
        func_frame.pack(fill="x", padx=10, pady=5)

        self.func_buttons = {}
        func_names = [
            "DC Voltage", "AC Voltage", "DC Current", "AC Current",
            "2-Wire Resistance", "4-Wire Resistance", "Frequency", "Period",
            "Continuity", "Diode", "Temperature", "Capacitance"
        ]
        short_names = [
            "DCV", "ACV", "DCI", "ACI", "\u03a92W", "\u03a94W",
            "FREQ", "PER", "CONT", "DIODE", "TEMP", "CAP"
        ]

        for i, (fname, sname) in enumerate(zip(func_names, short_names)):
            btn = ttk.Button(func_frame, text=sname, width=7,
                             style="Func.TButton" if fname != self.current_function else "FuncActive.TButton",
                             command=lambda f=fname: self._select_function(f))
            btn.grid(row=0, column=i, padx=2, pady=2, sticky="ew")
            self.func_buttons[fname] = btn
            func_frame.columnconfigure(i, weight=1)

        # ─── Main Body: Left (Config) + Right (Data) ────────────
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=10, pady=5)

        # Left panel - Configuration
        left = ttk.Frame(body, width=380)
        left.pack(side="left", fill="y", padx=(0, 5))
        left.pack_propagate(False)

        # -- Range --
        range_frame = ttk.LabelFrame(left, text=" Range ", style="TLabelframe")
        range_frame.pack(fill="x", pady=(0, 5))

        self.range_var = tk.StringVar(value="AUTO")
        ttk.Label(range_frame, text="Range:", style="Card.TLabel").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        self.range_combo = ttk.Combobox(range_frame, textvariable=self.range_var, width=15, state="readonly")
        self.range_combo.grid(row=0, column=1, padx=5, pady=3, sticky="ew")
        ttk.Button(range_frame, text="Set", command=self._set_range).grid(row=0, column=2, padx=5, pady=3)
        self.auto_range_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(range_frame, text="Auto Range", variable=self.auto_range_var,
                         command=self._toggle_auto_range).grid(row=1, column=0, columnspan=3, padx=5, pady=3, sticky="w")
        range_frame.columnconfigure(1, weight=1)

        # -- Integration / NPLC --
        nplc_frame = ttk.LabelFrame(left, text=" Integration Time ", style="TLabelframe")
        nplc_frame.pack(fill="x", pady=5)

        self.nplc_var = tk.StringVar(value="10")
        ttk.Label(nplc_frame, text="NPLC:", style="Card.TLabel").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        self.nplc_combo = ttk.Combobox(nplc_frame, textvariable=self.nplc_var, values=self.NPLC_VALUES,
                                        width=10, state="readonly")
        self.nplc_combo.grid(row=0, column=1, padx=5, pady=3, sticky="ew")
        ttk.Button(nplc_frame, text="Set", command=self._set_nplc).grid(row=0, column=2, padx=5, pady=3)

        self.autozero_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(nplc_frame, text="Auto Zero", variable=self.autozero_var,
                         command=self._toggle_autozero).grid(row=1, column=0, columnspan=3, padx=5, pady=3, sticky="w")

        # -- Bandwidth (for AC) --
        self.bw_label_text = tk.StringVar(value="Bandwidth:")
        bw_frame = ttk.LabelFrame(left, text=" AC Bandwidth / Aperture ", style="TLabelframe")
        bw_frame.pack(fill="x", pady=5)

        self.bw_var = tk.StringVar(value="20")
        ttk.Label(bw_frame, textvariable=self.bw_label_text, style="Card.TLabel").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        self.bw_combo = ttk.Combobox(bw_frame, textvariable=self.bw_var, width=10, state="readonly")
        self.bw_combo.grid(row=0, column=1, padx=5, pady=3, sticky="ew")
        ttk.Button(bw_frame, text="Set", command=self._set_bandwidth).grid(row=0, column=2, padx=5, pady=3)
        bw_frame.columnconfigure(1, weight=1)
        self.bw_frame = bw_frame

        # -- Input Impedance (for DCV) --
        imp_frame = ttk.LabelFrame(left, text=" Input Impedance ", style="TLabelframe")
        imp_frame.pack(fill="x", pady=5)
        self.impedance_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(imp_frame, text="Auto (>10 G\u03a9 on 0.1/1/10V)", variable=self.impedance_var,
                         command=self._toggle_impedance).pack(padx=5, pady=5, anchor="w")
        self.imp_frame = imp_frame

        # -- Null --
        null_frame = ttk.LabelFrame(left, text=" Null / Offset ", style="TLabelframe")
        null_frame.pack(fill="x", pady=5)
        self.null_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(null_frame, text="Enable Null", variable=self.null_var,
                         command=self._toggle_null).grid(row=0, column=0, padx=5, pady=3, sticky="w")
        self.null_value_var = tk.StringVar(value="0")
        ttk.Entry(null_frame, textvariable=self.null_value_var, width=15).grid(row=0, column=1, padx=5, pady=3)
        ttk.Button(null_frame, text="Set", command=self._set_null_value).grid(row=0, column=2, padx=5, pady=3)
        self.null_frame = null_frame

        # -- Temperature Config --
        temp_frame = ttk.LabelFrame(left, text=" Temperature Config ", style="TLabelframe")
        temp_frame.pack(fill="x", pady=5)
        self.temp_probe_var = tk.StringVar(value="RTD")
        self.temp_type_var = tk.StringVar(value="PT100")
        self.temp_unit_var = tk.StringVar(value="C")
        ttk.Label(temp_frame, text="Probe:", style="Card.TLabel").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        ttk.Combobox(temp_frame, textvariable=self.temp_probe_var,
                     values=["RTD", "FRTD", "THER"], width=8, state="readonly").grid(row=0, column=1, padx=5, pady=3)
        ttk.Label(temp_frame, text="Type:", style="Card.TLabel").grid(row=0, column=2, padx=5, pady=3, sticky="w")
        ttk.Combobox(temp_frame, textvariable=self.temp_type_var,
                     values=["PT100", "2252", "5000", "10000"], width=8, state="readonly").grid(row=0, column=3, padx=5, pady=3)
        ttk.Label(temp_frame, text="Units:", style="Card.TLabel").grid(row=1, column=0, padx=5, pady=3, sticky="w")
        ttk.Combobox(temp_frame, textvariable=self.temp_unit_var,
                     values=["C", "F", "K"], width=5, state="readonly").grid(row=1, column=1, padx=5, pady=3)
        ttk.Button(temp_frame, text="Apply", command=self._apply_temp_config).grid(row=1, column=2, columnspan=2, padx=5, pady=3)
        self.temp_frame = temp_frame

        # -- Trigger --
        trig_frame = ttk.LabelFrame(left, text=" Trigger ", style="TLabelframe")
        trig_frame.pack(fill="x", pady=5)
        self.trig_source_var = tk.StringVar(value="IMM")
        ttk.Label(trig_frame, text="Source:", style="Card.TLabel").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        ttk.Combobox(trig_frame, textvariable=self.trig_source_var,
                     values=["IMM", "BUS", "EXT", "INT"], width=8, state="readonly").grid(row=0, column=1, padx=5, pady=3)
        self.trig_delay_var = tk.StringVar(value="AUTO")
        ttk.Label(trig_frame, text="Delay:", style="Card.TLabel").grid(row=0, column=2, padx=5, pady=3, sticky="w")
        ttk.Entry(trig_frame, textvariable=self.trig_delay_var, width=8).grid(row=0, column=3, padx=5, pady=3)
        ttk.Button(trig_frame, text="Apply", command=self._apply_trigger).grid(row=0, column=4, padx=5, pady=3)
        self.trig_count_var = tk.StringVar(value="1")
        ttk.Label(trig_frame, text="Trig Count:", style="Card.TLabel").grid(row=1, column=0, padx=5, pady=3, sticky="w")
        ttk.Entry(trig_frame, textvariable=self.trig_count_var, width=8).grid(row=1, column=1, padx=5, pady=3)
        self.samp_count_var = tk.StringVar(value="1")
        ttk.Label(trig_frame, text="Samples:", style="Card.TLabel").grid(row=1, column=2, padx=5, pady=3, sticky="w")
        ttk.Entry(trig_frame, textvariable=self.samp_count_var, width=8).grid(row=1, column=3, padx=5, pady=3)

        # Right panel - Data & Controls
        right = ttk.Frame(body)
        right.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # -- Control buttons --
        ctrl_frame = ttk.Frame(right)
        ctrl_frame.pack(fill="x", pady=(0, 5))

        self.single_btn = ttk.Button(ctrl_frame, text="Single", style="Accent.TButton", command=self._single_reading)
        self.single_btn.pack(side="left", padx=2)
        self.stream_btn = ttk.Button(ctrl_frame, text="Stream", style="Accent.TButton", command=self._toggle_stream)
        self.stream_btn.pack(side="left", padx=2)

        ttk.Label(ctrl_frame, text="Interval (s):").pack(side="left", padx=(10, 2))
        self.interval_var = tk.StringVar(value="0.5")
        ttk.Entry(ctrl_frame, textvariable=self.interval_var, width=6).pack(side="left", padx=2)

        self.record_btn = ttk.Button(ctrl_frame, text="Record CSV", command=self._toggle_recording)
        self.record_btn.pack(side="left", padx=(10, 2))
        ttk.Button(ctrl_frame, text="Clear", command=self._clear_data).pack(side="left", padx=2)
        ttk.Button(ctrl_frame, text="Export", command=self._export_data).pack(side="left", padx=2)

        # -- Statistics bar --
        stats_frame = ttk.Frame(right)
        stats_frame.pack(fill="x", pady=(0, 5))

        self.stat_count_var = tk.StringVar(value="Count: 0")
        self.stat_min_var = tk.StringVar(value="Min: ---")
        self.stat_max_var = tk.StringVar(value="Max: ---")
        self.stat_avg_var = tk.StringVar(value="Avg: ---")
        self.stat_ptp_var = tk.StringVar(value="P-P: ---")

        for sv in [self.stat_count_var, self.stat_min_var, self.stat_max_var, self.stat_avg_var, self.stat_ptp_var]:
            ttk.Label(stats_frame, textvariable=sv, font=("Consolas", 9), foreground="#aaaaaa").pack(side="left", padx=8)

        # -- Data table --
        table_frame = ttk.Frame(right)
        table_frame.pack(fill="both", expand=True)

        cols = ("#", "Timestamp", "Value", "Unit")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=15)
        self.tree.heading("#", text="#", anchor="center")
        self.tree.heading("Timestamp", text="Timestamp", anchor="w")
        self.tree.heading("Value", text="Value", anchor="e")
        self.tree.heading("Unit", text="Unit", anchor="w")
        self.tree.column("#", width=50, anchor="center")
        self.tree.column("Timestamp", width=200, anchor="w")
        self.tree.column("Value", width=200, anchor="e")
        self.tree.column("Unit", width=80, anchor="w")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ─── Bottom: System Buttons ──────────────────────────────
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill="x", padx=10, pady=(5, 10))

        ttk.Button(bottom_frame, text="Reset", command=self._reset).pack(side="left", padx=2)
        ttk.Button(bottom_frame, text="Self Test", command=self._self_test).pack(side="left", padx=2)
        ttk.Button(bottom_frame, text="Errors", command=self._check_errors).pack(side="left", padx=2)
        ttk.Button(bottom_frame, text="Beep", command=self._beep).pack(side="left", padx=2)

        # Display text
        ttk.Label(bottom_frame, text="Display:").pack(side="left", padx=(20, 2))
        self.disp_text_var = tk.StringVar()
        ttk.Entry(bottom_frame, textvariable=self.disp_text_var, width=14).pack(side="left", padx=2)
        ttk.Button(bottom_frame, text="Set", command=self._set_display_text).pack(side="left", padx=2)
        ttk.Button(bottom_frame, text="Clear", command=self._clear_display_text).pack(side="left", padx=2)

        # Math functions
        ttk.Label(bottom_frame, text="Math:").pack(side="left", padx=(20, 2))
        self.math_var = tk.StringVar(value="OFF")
        ttk.Combobox(bottom_frame, textvariable=self.math_var,
                     values=["OFF", "NULL", "DB", "DBM", "AVER", "LIM"], width=6, state="readonly").pack(side="left", padx=2)
        ttk.Button(bottom_frame, text="Apply", command=self._apply_math).pack(side="left", padx=2)

        self.rec_label = ttk.Label(bottom_frame, text="", foreground="#ff4444")
        self.rec_label.pack(side="right", padx=5)

        # Update UI state
        self._update_config_visibility()

    def _populate_resources(self):
        try:
            resources = self.dmm.list_resources()
            self.resource_combo["values"] = resources
            # Default to GPIB resource if found
            for r in resources:
                if "GPIB" in r:
                    self.resource_var.set(r)
                    break
            else:
                if resources:
                    self.resource_var.set(resources[0])
        except Exception as e:
            self.status_var.set(f"Error: {e}")

    def _auto_connect(self):
        """Auto-connect to GPIB instrument if found."""
        resource = self.resource_var.get()
        if resource and "GPIB" in resource and not self.dmm.connected:
            self._toggle_connection()

    def _toggle_connection(self):
        if self.dmm.connected:
            self._stop_stream()
            self._stop_recording()
            self.dmm.disconnect()
            self.connect_btn.configure(text="Connect", style="Accent.TButton")
            self.status_var.set("Disconnected")
            self.idn_var.set("")
            self.terminals_var.set("")
            self.display_value.set("---")
            self.display_unit.set("")
            self.display_func.set("")
        else:
            resource = self.resource_var.get()
            if not resource:
                messagebox.showwarning("No Resource", "Select a VISA resource first.")
                return
            try:
                idn = self.dmm.connect(resource)
                self.connect_btn.configure(text="Disconnect", style="Stop.TButton")
                self.status_var.set("Connected")
                self.idn_var.set(idn)
                try:
                    terminals = self.dmm.get_terminals()
                    self.terminals_var.set(f"Terminals: {terminals}")
                except Exception:
                    pass
                self.display_func.set(self.current_function)
                self._read_current_config()
            except Exception as e:
                messagebox.showerror("Connection Error", str(e))

    def _read_current_config(self):
        """Read current instrument configuration after connecting."""
        try:
            func_info = self.FUNCTIONS[self.current_function]
            if func_info["has_nplc"]:
                try:
                    nplc = self.dmm.query(f"{func_info['sense_prefix']}:NPLC?")
                    self.nplc_var.set(nplc)
                except Exception:
                    pass
        except Exception:
            pass

    def _select_function(self, func_name):
        self.current_function = func_name
        func_info = self.FUNCTIONS[func_name]

        # Update button highlighting
        for fname, btn in self.func_buttons.items():
            btn.configure(style="FuncActive.TButton" if fname == func_name else "Func.TButton")

        # Update display
        self.display_func.set(func_name)
        self.display_unit.set(func_info["unit"])

        # Update range options
        ranges = func_info["ranges"]
        self.range_combo["values"] = ranges
        if ranges:
            self.range_var.set("AUTO")
        else:
            self.range_var.set("")

        # Update bandwidth/aperture
        if func_info["has_bandwidth"]:
            self.bw_combo["values"] = self.BANDWIDTH_VALUES
            self.bw_var.set("20")
            self.bw_label_text.set("Bandwidth (Hz):")
        elif func_info["has_aperture"]:
            self.bw_combo["values"] = self.APERTURE_VALUES
            self.bw_var.set("0.1")
            self.bw_label_text.set("Aperture (s):")

        # Configure the instrument
        if self.dmm.connected:
            try:
                if func_name == "Temperature":
                    probe = self.temp_probe_var.get()
                    ttype = self.temp_type_var.get()
                    self.dmm.write(f"CONF:TEMP {probe},{ttype}")
                elif func_name in ("Continuity", "Diode"):
                    self.dmm.write(func_info["conf_cmd"])
                else:
                    self.dmm.write(f"{func_info['conf_cmd']} AUTO")
            except Exception as e:
                self._show_error(f"Configure failed: {e}")

        self._update_config_visibility()

    def _update_config_visibility(self):
        func_info = self.FUNCTIONS[self.current_function]

        # Show/hide config sections based on function capabilities
        if func_info["has_bandwidth"] or func_info["has_aperture"]:
            self.bw_frame.pack(fill="x", pady=5)
        else:
            self.bw_frame.pack_forget()

        if func_info["has_impedance"]:
            self.imp_frame.pack(fill="x", pady=5)
        else:
            self.imp_frame.pack_forget()

        if func_info["has_null"]:
            self.null_frame.pack(fill="x", pady=5)
        else:
            self.null_frame.pack_forget()

        if self.current_function == "Temperature":
            self.temp_frame.pack(fill="x", pady=5)
        else:
            self.temp_frame.pack_forget()

    # ─── Measurement Actions ─────────────────────────────────────

    def _take_reading(self):
        """Take a single reading and return (value, unit, timestamp)."""
        func_info = self.FUNCTIONS[self.current_function]
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        value = self.dmm.query_float(func_info["meas_cmd"])
        unit = func_info["unit"]
        return value, unit, timestamp

    def _format_value(self, value):
        """Format a measurement value for display."""
        abs_val = abs(value)
        if abs_val == 0:
            return "0.000000"
        elif abs_val >= 1e6:
            return f"{value / 1e6:.6f} M"
        elif abs_val >= 1e3:
            return f"{value / 1e3:.6f} k"
        elif abs_val >= 1:
            return f"{value:.6f}"
        elif abs_val >= 1e-3:
            return f"{value * 1e3:.6f} m"
        elif abs_val >= 1e-6:
            return f"{value * 1e6:.6f} \u00b5"
        elif abs_val >= 1e-9:
            return f"{value * 1e9:.6f} n"
        else:
            return f"{value:.6E}"

    def _update_display(self, value, unit):
        """Update the main display with a reading."""
        formatted = self._format_value(value)
        self.display_value.set(formatted)
        self.display_unit.set(unit)

    def _add_reading_to_table(self, value, unit, timestamp):
        """Add a reading to the data table and update stats."""
        self.reading_count += 1
        self.readings_list.append((self.reading_count, timestamp, value, unit))

        self.tree.insert("", "end", values=(
            self.reading_count,
            timestamp,
            f"{value:+.8E}",
            unit
        ))
        self.tree.yview_moveto(1.0)  # Auto-scroll to bottom

        # Update statistics
        self.stats["count"] += 1
        self.stats["sum"] += value
        if self.stats["min"] is None or value < self.stats["min"]:
            self.stats["min"] = value
        if self.stats["max"] is None or value > self.stats["max"]:
            self.stats["max"] = value

        self.stat_count_var.set(f"Count: {self.stats['count']}")
        self.stat_min_var.set(f"Min: {self.stats['min']:+.6E}")
        self.stat_max_var.set(f"Max: {self.stats['max']:+.6E}")
        avg = self.stats["sum"] / self.stats["count"]
        self.stat_avg_var.set(f"Avg: {avg:+.6E}")
        ptp = self.stats["max"] - self.stats["min"]
        self.stat_ptp_var.set(f"P-P: {ptp:.6E}")

        # Write to CSV if recording
        if self.recording and self.csv_writer:
            self.csv_writer.writerow([self.reading_count, timestamp, f"{value:+.8E}", unit])
            self.csv_file.flush()

    def _single_reading(self):
        if not self.dmm.connected:
            messagebox.showwarning("Not Connected", "Connect to the instrument first.")
            return

        def do_read():
            try:
                value, unit, timestamp = self._take_reading()
                self.after(0, lambda: self._update_display(value, unit))
                self.after(0, lambda: self._add_reading_to_table(value, unit, timestamp))
            except Exception as e:
                self.after(0, lambda: self._show_error(f"Read error: {e}"))

        threading.Thread(target=do_read, daemon=True).start()

    def _toggle_stream(self):
        if self.streaming:
            self._stop_stream()
        else:
            self._start_stream()

    def _start_stream(self):
        if not self.dmm.connected:
            messagebox.showwarning("Not Connected", "Connect to the instrument first.")
            return
        self.streaming = True
        self.stream_btn.configure(text="Stop", style="Stop.TButton")
        self.single_btn.configure(state="disabled")
        self.stream_thread = threading.Thread(target=self._stream_loop, daemon=True)
        self.stream_thread.start()

    def _stop_stream(self):
        self.streaming = False
        self.stream_btn.configure(text="Stream", style="Accent.TButton")
        self.single_btn.configure(state="normal")

    def _stream_loop(self):
        while self.streaming:
            try:
                interval = float(self.interval_var.get())
            except ValueError:
                interval = 0.5

            try:
                value, unit, timestamp = self._take_reading()
                self.after(0, lambda v=value, u=unit: self._update_display(v, u))
                self.after(0, lambda v=value, u=unit, t=timestamp: self._add_reading_to_table(v, u, t))
            except Exception as e:
                self.after(0, lambda: self._show_error(f"Stream error: {e}"))
                self.after(0, self._stop_stream)
                break

            time.sleep(interval)

    # ─── Recording ───────────────────────────────────────────────

    def _toggle_recording(self):
        if self.recording:
            self._stop_recording()
        else:
            self._start_recording()

    def _start_recording(self):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"34410A_{self.current_function.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not filepath:
            return

        self.csv_file = open(filepath, 'w', newline='')
        self.csv_writer = csv.writer(self.csv_file)
        self.csv_writer.writerow(["#", "Timestamp", "Value", "Unit"])
        self.recording = True
        self.record_btn.configure(text="Stop Rec", style="Stop.TButton")
        self.rec_label.configure(text=f"REC: {os.path.basename(filepath)}")

    def _stop_recording(self):
        self.recording = False
        if self.csv_file:
            self.csv_file.close()
            self.csv_file = None
            self.csv_writer = None
        self.record_btn.configure(text="Record CSV", style="TButton")
        self.rec_label.configure(text="")

    def _clear_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.reading_count = 0
        self.readings_list = []
        self.stats = {"min": None, "max": None, "sum": 0, "count": 0}
        self.stat_count_var.set("Count: 0")
        self.stat_min_var.set("Min: ---")
        self.stat_max_var.set("Max: ---")
        self.stat_avg_var.set("Avg: ---")
        self.stat_ptp_var.set("P-P: ---")

    def _export_data(self):
        if not self.readings_list:
            messagebox.showinfo("No Data", "No readings to export.")
            return
        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"34410A_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not filepath:
            return
        with open(filepath, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["#", "Timestamp", "Value", "Unit"])
            for row in self.readings_list:
                writer.writerow([row[0], row[1], f"{row[2]:+.8E}", row[3]])
        messagebox.showinfo("Exported", f"Exported {len(self.readings_list)} readings to:\n{filepath}")

    # ─── Configuration Commands ──────────────────────────────────

    def _set_range(self):
        if not self.dmm.connected:
            return
        func_info = self.FUNCTIONS[self.current_function]
        range_val = self.range_var.get()
        try:
            if range_val == "AUTO":
                self.dmm.write(f"{func_info['sense_prefix']}:RANG:AUTO ON")
                self.auto_range_var.set(True)
            else:
                self.dmm.write(f"{func_info['sense_prefix']}:RANG:AUTO OFF")
                self.dmm.write(f"{func_info['sense_prefix']}:RANG {range_val}")
                self.auto_range_var.set(False)
        except Exception as e:
            self._show_error(f"Set range failed: {e}")

    def _toggle_auto_range(self):
        if not self.dmm.connected:
            return
        func_info = self.FUNCTIONS[self.current_function]
        state = "ON" if self.auto_range_var.get() else "OFF"
        try:
            self.dmm.write(f"{func_info['sense_prefix']}:RANG:AUTO {state}")
        except Exception as e:
            self._show_error(f"Auto range failed: {e}")

    def _set_nplc(self):
        if not self.dmm.connected:
            return
        func_info = self.FUNCTIONS[self.current_function]
        if not func_info["has_nplc"]:
            return
        try:
            self.dmm.write(f"{func_info['sense_prefix']}:NPLC {self.nplc_var.get()}")
        except Exception as e:
            self._show_error(f"Set NPLC failed: {e}")

    def _toggle_autozero(self):
        if not self.dmm.connected:
            return
        func_info = self.FUNCTIONS[self.current_function]
        state = "ON" if self.autozero_var.get() else "OFF"
        try:
            self.dmm.write(f"{func_info['sense_prefix']}:ZERO:AUTO {state}")
        except Exception as e:
            self._show_error(f"Auto zero failed: {e}")

    def _set_bandwidth(self):
        if not self.dmm.connected:
            return
        func_info = self.FUNCTIONS[self.current_function]
        try:
            if func_info["has_bandwidth"]:
                self.dmm.write(f"{func_info['sense_prefix']}:BAND {self.bw_var.get()}")
            elif func_info["has_aperture"]:
                self.dmm.write(f"{func_info['sense_prefix']}:APER {self.bw_var.get()}")
        except Exception as e:
            self._show_error(f"Set bandwidth/aperture failed: {e}")

    def _toggle_impedance(self):
        if not self.dmm.connected:
            return
        state = "ON" if self.impedance_var.get() else "OFF"
        try:
            self.dmm.write(f"VOLT:DC:IMP:AUTO {state}")
        except Exception as e:
            self._show_error(f"Set impedance failed: {e}")

    def _toggle_null(self):
        if not self.dmm.connected:
            return
        func_info = self.FUNCTIONS[self.current_function]
        state = "ON" if self.null_var.get() else "OFF"
        try:
            self.dmm.write(f"{func_info['sense_prefix']}:NULL {state}")
        except Exception as e:
            self._show_error(f"Set null failed: {e}")

    def _set_null_value(self):
        if not self.dmm.connected:
            return
        func_info = self.FUNCTIONS[self.current_function]
        try:
            self.dmm.write(f"{func_info['sense_prefix']}:NULL:VAL {self.null_value_var.get()}")
        except Exception as e:
            self._show_error(f"Set null value failed: {e}")

    def _apply_temp_config(self):
        if not self.dmm.connected:
            return
        try:
            probe = self.temp_probe_var.get()
            ttype = self.temp_type_var.get()
            unit = self.temp_unit_var.get()
            self.dmm.write(f"CONF:TEMP {probe},{ttype}")
            self.dmm.write(f"UNIT:TEMP {unit}")
            unit_symbols = {"C": "\u00b0C", "F": "\u00b0F", "K": "K"}
            self.FUNCTIONS["Temperature"]["unit"] = unit_symbols.get(unit, "\u00b0C")
            self.display_unit.set(self.FUNCTIONS["Temperature"]["unit"])
        except Exception as e:
            self._show_error(f"Temp config failed: {e}")

    def _apply_trigger(self):
        if not self.dmm.connected:
            return
        try:
            self.dmm.write(f"TRIG:SOUR {self.trig_source_var.get()}")
            delay = self.trig_delay_var.get()
            if delay.upper() == "AUTO":
                self.dmm.write("TRIG:DEL:AUTO ON")
            else:
                self.dmm.write(f"TRIG:DEL:AUTO OFF")
                self.dmm.write(f"TRIG:DEL {delay}")
            self.dmm.write(f"TRIG:COUN {self.trig_count_var.get()}")
            self.dmm.write(f"SAMP:COUN {self.samp_count_var.get()}")
        except Exception as e:
            self._show_error(f"Trigger config failed: {e}")

    def _apply_math(self):
        if not self.dmm.connected:
            return
        try:
            func = self.math_var.get()
            if func == "OFF":
                self.dmm.write("CALC:STAT OFF")
            else:
                self.dmm.write(f"CALC:FUNC {func}")
                self.dmm.write("CALC:STAT ON")
        except Exception as e:
            self._show_error(f"Math config failed: {e}")

    # ─── System Commands ─────────────────────────────────────────

    def _reset(self):
        if not self.dmm.connected:
            return
        self._stop_stream()
        self.dmm.reset()
        self.dmm.clear_status()
        self.display_value.set("RESET")
        self.after(1500, lambda: self.display_value.set("---"))

    def _self_test(self):
        if not self.dmm.connected:
            return

        def run_test():
            try:
                self.after(0, lambda: self.display_value.set("TESTING..."))
                result = self.dmm.self_test()
                status = "PASS" if result == 0 else f"FAIL ({result})"
                self.after(0, lambda: self.display_value.set(status))
                self.after(0, lambda: messagebox.showinfo("Self Test", f"Result: {status}"))
            except Exception as e:
                self.after(0, lambda: self._show_error(f"Self test failed: {e}"))

        threading.Thread(target=run_test, daemon=True).start()

    def _check_errors(self):
        if not self.dmm.connected:
            return
        errors = self.dmm.get_all_errors()
        if errors:
            messagebox.showwarning("Instrument Errors", "\n".join(errors))
        else:
            messagebox.showinfo("No Errors", "Error queue is empty.")

    def _beep(self):
        if self.dmm.connected:
            self.dmm.beep()

    def _set_display_text(self):
        if not self.dmm.connected:
            return
        text = self.disp_text_var.get()[:12]
        try:
            self.dmm.write(f'DISP:WIND:TEXT "{text}"')
        except Exception as e:
            self._show_error(f"Display text failed: {e}")

    def _clear_display_text(self):
        if not self.dmm.connected:
            return
        try:
            self.dmm.write("DISP:WIND:TEXT:CLE")
        except Exception as e:
            self._show_error(f"Clear display failed: {e}")

    def _show_error(self, msg):
        self.display_value.set("ERROR")
        messagebox.showerror("Error", msg)

    def destroy(self):
        self._stop_stream()
        self._stop_recording()
        if self.dmm.connected:
            self.dmm.disconnect()
        super().destroy()


if __name__ == "__main__":
    app = DMMApplication()
    app.mainloop()
