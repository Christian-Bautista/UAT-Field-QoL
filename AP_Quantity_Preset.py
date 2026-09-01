"""Generate a UAT field report whose section 3.2.1 table covers the number of APs chosen in the GUI."""

import copy
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

TEMPLATE = "UAT Field Template.docx"
SECTION_HEADING = "Signal Coverage Test"
ROWS_PER_AP = 5  # AP ID, Test Device 1, SSID/Band header, 2.4 GHz, 5 GHz


def find_ap_table(document):
    """Return the table under section 3.2.1 (the first table of the Signal Coverage Test section)."""
    in_section = False
    for element in document.element.body.iterchildren():
        tag = element.tag.split("}")[-1]
        if tag == "p":
            paragraph = Paragraph(element, document)
            style = paragraph.style.name
            if style == "Heading 2":
                in_section = paragraph.text.strip() == SECTION_HEADING
            elif style == "Heading 1":
                in_section = False
        elif tag == "tbl" and in_section:
            return Table(element, document)
    raise LookupError("Could not find the section 3.2.1 table in the template.")


def trim_ap_block(table):
    """Drop the template's trailing band rows so one AP block matches the full template's form."""
    for row in table._tbl.tr_lst[ROWS_PER_AP:]:
        table._tbl.remove(row)


def generate_document(ap_count, output_path, report):
    """Write a report holding `ap_count` AP blocks, calling `report(message, fraction)` as it goes."""
    report("Opening template...", 0.0)
    document = Document(TEMPLATE)

    report("Locating the section 3.2.1 table...", 0.1)
    table = find_ap_table(document)

    report("Preparing the AP block...", 0.2)
    trim_ap_block(table)
    block = [copy.deepcopy(row) for row in table._tbl.tr_lst]

    for ap_number in range(2, ap_count + 1):
        report(
            "Adding AP %d of %d..." % (ap_number, ap_count),
            0.2 + 0.7 * (ap_number - 1) / ap_count,
        )
        for row in block:
            table._tbl.append(copy.deepcopy(row))

    report("Saving document...", 0.9)
    document.save(output_path)
    report("Done.", 1.0)


class APQuantityPreset:
    """The window that collects the AP count and document name, and reports generation progress."""

    def __init__(self, root):
        self.root = root
        self.root.title("AP Quantity Preset")
        self.save_folder = os.getcwd()
        self.updates = queue.Queue()

        frame = ttk.Frame(root, padding=12)
        frame.grid(sticky="nsew")

        self.location_button = ttk.Button(frame, text="File Location", command=self.choose_folder)
        self.location_button.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        ttk.Label(frame, text="How many APs to be tested: ").grid(row=1, column=0, sticky="w")
        self.ap_count_entry = ttk.Entry(frame, width=28)
        self.ap_count_entry.grid(row=1, column=1, sticky="w")

        ttk.Label(frame, text="Name of document: ").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.name_entry = ttk.Entry(frame, width=28)
        self.name_entry.grid(row=2, column=1, sticky="w", pady=(8, 0))

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, columnspan=2, sticky="w", pady=(12, 0))

        self.generate_button = ttk.Button(actions, text="Generate Document", command=self.generate)
        self.generate_button.grid(row=0, column=0, sticky="w")

        self.progress = ttk.Progressbar(actions, length=160, maximum=100)
        self.progress.grid(row=0, column=1, padx=(12, 8))

        self.status = ttk.Label(actions, text="", width=26, anchor="w")
        self.status.grid(row=0, column=2, sticky="w")

    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=self.save_folder, title="File Location")
        if folder:
            self.save_folder = folder

    def set_inputs_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        self.ap_count_entry.configure(state=state)
        self.name_entry.configure(state=state)
        self.generate_button.configure(state=state)
        self.location_button.configure(state=state)

    def generate(self):
        try:
            ap_count = int(self.ap_count_entry.get())
        except ValueError:
            messagebox.showerror("AP Quantity Preset", "Please enter a whole number of APs.")
            return
        if ap_count < 1:
            messagebox.showerror("AP Quantity Preset", "The number of APs must be at least 1.")
            return

        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror("AP Quantity Preset", "Please enter a name for the document.")
            return
        if not name.lower().endswith(".docx"):
            name += ".docx"
        output_path = os.path.join(self.save_folder, name)

        self.set_inputs_enabled(False)
        self.progress.configure(value=0)
        self.status.configure(text="Starting...")

        worker = threading.Thread(target=self.run_generation, args=(ap_count, output_path), daemon=True)
        worker.start()
        self.root.after(100, self.drain_updates)

    def run_generation(self, ap_count, output_path):
        try:
            generate_document(ap_count, output_path, lambda message, fraction: self.updates.put(("progress", message, fraction)))
        except Exception as error:  # reported in the window rather than the console
            self.updates.put(("error", str(error), 0.0))
        else:
            self.updates.put(("finished", output_path, 1.0))

    def drain_updates(self):
        finished = False
        while True:
            try:
                kind, message, fraction = self.updates.get_nowait()
            except queue.Empty:
                break
            self.progress.configure(value=fraction * 100)
            if kind == "progress":
                self.status.configure(text=message)
            elif kind == "error":
                self.status.configure(text="Failed.")
                messagebox.showerror("AP Quantity Preset", message)
                finished = True
            else:
                self.status.configure(text="Saved %s" % os.path.basename(message))
                finished = True
        if finished:
            self.set_inputs_enabled(True)
        else:
            self.root.after(100, self.drain_updates)


def main():
    root = tk.Tk()
    APQuantityPreset(root)
    root.mainloop()


if __name__ == "__main__":
    main()
