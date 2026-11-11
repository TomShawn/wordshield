"""Bilingual Tkinter interface for non-technical users."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from .core import (Analysis, Package, SafetyError, analyze, export_redacted,
                   export_restored)
from .crypto import MappingError
from .i18n import error_text, issue_text, notice_text, part_text, status_text, tr
from .storage import OutputError


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.pool = ThreadPoolExecutor(max_workers=1)
        self.analysis: Analysis | None = None
        self.source: Path | None = None
        self.busy = False
        self.last_output: Path | None = None
        self.last_restore_report: dict | None = None
        self.restore_issues: list = []
        self.language = "en"
        self.localized_widgets: list[tuple[tk.Widget, str]] = []
        self.buttons: list[ttk.Button] = []
        self.status_key = "status_choose"
        self.status_values: dict = {}

        root.geometry("1080x800")
        root.minsize(900, 720)
        root.report_callback_exception = self.callback_error
        root.protocol("WM_DELETE_WINDOW", self.close)
        style = ttk.Style(root)
        style.configure("Title.TLabel", font=("PingFang SC", 21, "bold"))

        main = ttk.Frame(root, padding=16)
        main.pack(fill="both", expand=True)
        heading = ttk.Frame(main)
        heading.pack(fill="x")
        self.title_label = ttk.Label(heading, style="Title.TLabel")
        self.title_label.pack(side="left")
        self.language_choice = tk.StringVar(value="English")
        self.language_box = ttk.Combobox(
            heading, textvariable=self.language_choice,
            values=("中文", "English"), state="readonly", width=10,
        )
        self.language_box.pack(side="right")
        self.language_label = ttk.Label(heading)
        self.language_label.pack(side="right", padx=(0, 8))
        self.language_box.bind("<<ComboboxSelected>>", self.switch_language)

        self.disclaimer_label = ttk.Label(main, wraplength=1000)
        self.disclaimer_label.pack(anchor="w", pady=(8, 18))
        self.status = tk.StringVar()
        ttk.Label(main, textvariable=self.status, wraplength=1000).pack(
            side="bottom", anchor="w", pady=(8, 0)
        )
        self.tabs = ttk.Notebook(main)
        self.tabs.pack(fill="both", expand=True)
        self.redaction_frame = ttk.Frame(self.tabs, padding=14)
        self.restore_frame = ttk.Frame(self.tabs, padding=14)
        self.tabs.add(self.redaction_frame)
        self.tabs.add(self.restore_frame)
        self.make_redaction(self.redaction_frame)
        self.make_restore(self.restore_frame)
        self.apply_language()

    def t(self, key: str, **values) -> str:
        return tr(self.language, key, **values)

    def localize(self, widget, key: str):
        self.localized_widgets.append((widget, key))
        return widget

    def button(self, parent, key: str, command):
        widget = self.localize(ttk.Button(parent, command=command), key)
        self.buttons.append(widget)
        return widget

    def set_status(self, key: str, **values) -> None:
        self.status_key = key
        self.status_values = values
        self.status.set(self.t(key, **values))

    def switch_language(self, _event=None):
        self.language = "en" if self.language_choice.get() == "English" else "zh"
        self.apply_language()

    def apply_language(self):
        self.root.title(self.t("app_title"))
        self.title_label.configure(text=self.t("app_heading"))
        self.language_label.configure(text=self.t("language"))
        self.disclaimer_label.configure(text=self.t("disclaimer"))
        self.tabs.tab(self.redaction_frame, text=self.t("tab_redact"))
        self.tabs.tab(self.restore_frame, text=self.t("tab_restore"))
        for widget, key in self.localized_widgets:
            widget.configure(text=self.t(key))
        for column, key in (
            ("text", "column_text"), ("group", "column_group"),
            ("origin", "column_origin"), ("location", "column_location"),
        ):
            self.tree.heading(column, text=self.t(key))
        if self.source is None:
            self.source_label.set(self.t("no_file"))
        status_values = self.status_values
        if self.status_key == "restore_complete_status" and self.last_restore_report is not None:
            status_values = {
                "status": status_text(self.language, self.last_restore_report["status"])
            }
        self.status.set(self.t(self.status_key, **status_values))
        self.refresh(reset_confirmations=False)
        if self.restore_issues:
            self.render_restore_issues()
        elif self.last_restore_report is None:
            self.set_text(self.restore_report, self.t("restore_pending"))
        else:
            self.render_restore_report(self.last_restore_report)

    def make_redaction(self, frame):
        footer = ttk.Frame(frame)
        footer.pack(side="bottom", fill="x", pady=(6, 0))
        top = ttk.Frame(frame)
        top.pack(fill="x")
        self.button(top, "choose_analyze", self.choose_source).pack(side="left")
        self.source_label = tk.StringVar()
        ttk.Label(top, textvariable=self.source_label).pack(side="left", padx=12)
        self.preview_help = self.localize(ttk.Label(frame, wraplength=950), "preview_help")
        self.preview_help.pack(anchor="w", pady=(12, 6))

        columns = ("text", "group", "origin", "location")
        holder = ttk.Frame(frame)
        holder.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(
            holder, columns=columns, show="headings", selectmode="browse", height=10
        )
        for column, width in (("text", 220), ("group", 100), ("origin", 140), ("location", 390)):
            self.tree.column(column, width=width, minwidth=70)
        scroll = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda _: self.edit_group())
        self.tree.tag_configure("excluded", foreground="#777777")

        actions = ttk.Frame(footer)
        actions.pack(fill="x", pady=8)
        self.button(actions, "edit_group", self.edit_group).pack(side="left")
        self.button(actions, "rename_group", self.rename_group).pack(side="left", padx=(6, 0))
        self.button(actions, "separate_group", self.separate_group).pack(side="left", padx=6)
        self.button(actions, "split_highlight", self.split).pack(side="left")
        self.button(actions, "toggle_retain", self.toggle_exclusion).pack(side="left", padx=6)

        self.include_unmarked = tk.BooleanVar(value=True)
        self.include_unmarked_check = self.localize(
            ttk.Checkbutton(footer, variable=self.include_unmarked, command=self.refresh),
            "include_unmarked",
        )
        self.include_unmarked_check.pack(anchor="w")
        self.risk = tk.Text(footer, height=3, wrap="word", font=("PingFang SC", 11))
        self.risk.pack(fill="x", pady=8)
        self.confirmed = tk.BooleanVar(value=False)
        self.ack = tk.BooleanVar(value=False)
        self.confirmed_check = self.localize(
            ttk.Checkbutton(footer, variable=self.confirmed), "confirm_preview"
        )
        self.confirmed_check.pack(anchor="w")
        self.ack_check = self.localize(
            ttk.Checkbutton(footer, variable=self.ack), "confirm_risk"
        )
        self.ack_check.pack(anchor="w", pady=4)
        self.button(footer, "export_redacted", self.redact).pack(anchor="e", pady=(8, 0))

    def make_restore(self, frame):
        self.button(frame, "open_restored", self.open_output).pack(
            side="bottom", anchor="e", pady=(12, 0)
        )
        self.restore_intro = self.localize(ttk.Label(frame, wraplength=950), "restore_intro")
        self.restore_intro.pack(anchor="w", pady=(0, 16))
        self.final_path = tk.StringVar()
        self.map_path = tk.StringVar()
        for key, variable, callback in (
            ("final_docx", self.final_path, self.choose_final),
            ("encrypted_mapping", self.map_path, self.choose_mapping),
        ):
            row = ttk.Frame(frame)
            row.pack(fill="x", pady=6)
            self.localize(ttk.Label(row, width=18), key).pack(side="left")
            ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)
            self.button(row, "choose_file", callback).pack(side="left", padx=(8, 0))
        self.button(frame, "inspect_restore", self.restore).pack(anchor="e", pady=14)
        self.restore_report = tk.Text(frame, wrap="word", font=("PingFang SC", 12), height=17)
        self.restore_report.pack(fill="both", expand=True)

    @staticmethod
    def set_text(widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def run(self, task, done, message_key):
        if self.busy:
            return
        self.busy = True
        self.set_status(message_key)
        for button in self.buttons:
            button.configure(state="disabled")
        future = self.pool.submit(task)

        def poll():
            if not future.done():
                self.root.after(100, poll)
                return
            self.busy = False
            for button in self.buttons:
                button.configure(state="normal")
            try:
                result = future.result()
            except (SafetyError, MappingError, OutputError) as exc:
                self.set_status("processing_stopped")
                messagebox.showerror(
                    self.t("not_exported"), error_text(self.language, str(exc)), parent=self.root
                )
            except OSError:
                self.set_status("io_failed")
                messagebox.showerror(
                    self.t("io_failed_title"), self.t("io_failed_body"), parent=self.root
                )
            except Exception:
                self.set_status("processing_failed")
                messagebox.showerror(
                    self.t("processing_failed_title"), self.t("processing_failed_body"),
                    parent=self.root,
                )
            else:
                done(result)
        self.root.after(100, poll)

    def choose_source(self):
        path = filedialog.askopenfilename(
            title=self.t("choose_word"), filetypes=[(self.t("word_doc"), "*.docx")]
        )
        if not path:
            return
        self.analysis = None
        self.source = Path(path)
        self.source_label.set(self.source.name)
        self.refresh()
        self.run(lambda: analyze(Path(path).read_bytes()), self.loaded, "analyzing")

    def loaded(self, result):
        self.analysis = result
        self.refresh()
        marked = sum(item.highlighted for item in result.occurrences)
        unmarked = len(result.occurrences) - marked
        self.set_status("analysis_summary", marked=marked, unmarked=unmarked)
        if not marked:
            messagebox.showinfo(
                self.t("no_highlight_title"), self.t("no_highlight_body"), parent=self.root
            )

    def refresh(self, reset_confirmations=True):
        selection = self.tree.selection()
        scroll_position = self.tree.yview()
        if reset_confirmations:
            self.confirmed.set(False)
            self.ack.set(False)
        self.tree.delete(*self.tree.get_children())
        if self.analysis is None:
            self.set_text(self.risk, self.t("risk_pending"))
            return
        for item in self.analysis.occurrences:
            display = "「" + item.text.replace("\t", "⇥").replace("\n", "↵") + "」"
            retained = item.excluded or (not item.highlighted and not self.include_unmarked.get())
            origin = self.t("origin_highlight" if item.highlighted else "origin_unmarked")
            action = self.t("action_retain" if retained else "action_replace")
            location = self.t(
                "location", part=item.part, paragraph=item.paragraph + 1,
                character=item.start + 1,
            )
            self.tree.insert(
                "", "end", iid=item.key,
                values=(display, item.group, origin + " · " + action, location),
                tags=("excluded",) if retained else (),
            )
        for key in selection:
            if self.tree.exists(key):
                self.tree.selection_set(key)
                self.tree.focus(key)
        if scroll_position:
            self.tree.yview_moveto(scroll_position[0])
        messages = [
            self.t("must_fix" if issue.blocking else "review")
            + issue_text(self.language, issue.message) + f" ({issue.part})"
            for issue in self.analysis.issues
        ]
        messages.extend([self.t("metadata_notice"), self.t("retention_notice")])
        self.set_text(self.risk, "\n".join(messages))

    def selected(self):
        if self.busy:
            return None
        keys = self.tree.selection()
        if not keys or self.analysis is None:
            messagebox.showinfo(
                self.t("select_item_title"), self.t("select_item_body"), parent=self.root
            )
            return None
        return next(item for item in self.analysis.occurrences if item.key == keys[0])

    def edit_group(self):
        item = self.selected()
        if item is None:
            return
        value = simpledialog.askstring(
            self.t("group_title"), self.t("group_prompt"),
            initialvalue=item.group, parent=self.root,
        )
        if value and value.strip():
            item.group = value.strip()
            self.refresh()

    def separate_group(self):
        item = self.selected()
        if item is not None:
            item.group = self.analysis.new_group()
            self.refresh()

    def rename_group(self):
        item = self.selected()
        if item is None:
            return
        count = sum(member.group == item.group for member in self.analysis.occurrences)
        value = simpledialog.askstring(
            self.t("rename_group"), self.t("rename_group_prompt", count=count),
            initialvalue=item.group, parent=self.root,
        )
        if value is None:
            return
        try:
            self.analysis.rename_group(item.group, value)
        except SafetyError as exc:
            messagebox.showerror(
                self.t("rename_group_failed"), error_text(self.language, str(exc)), parent=self.root
            )
        else:
            self.refresh()

    def toggle_exclusion(self):
        item = self.selected()
        if item is not None:
            self.analysis.exclude(item.key, not item.excluded)
            self.refresh()

    def split(self):
        item = self.selected()
        if item is None:
            return
        value = simpledialog.askstring(
            self.t("split_title"), self.t("split_prompt"), parent=self.root
        )
        if value:
            try:
                offsets = [int(position.strip()) for position in value.split(",")]
                self.analysis.split(item.key, offsets)
            except (ValueError, SafetyError) as exc:
                body = error_text(self.language, str(exc)) if isinstance(exc, SafetyError) else self.t("integer_positions")
                messagebox.showerror(self.t("split_failed"), body, parent=self.root)
            else:
                self.refresh()

    def password(self, confirm=False):
        prompt = self.t("password_prompt") + (self.t("password_requirement") if confirm else "")
        first = simpledialog.askstring(
            self.t("mapping_password"), prompt, show="•", parent=self.root
        )
        if first is None:
            return None
        if confirm:
            second = simpledialog.askstring(
                self.t("confirm_password"), self.t("confirm_password_prompt"),
                show="•", parent=self.root,
            )
            if second is None:
                return None
            if first != second or len(first) < 12:
                messagebox.showerror(
                    self.t("bad_password_title"), self.t("bad_password_body"), parent=self.root
                )
                return None
        return first

    def redact(self):
        if self.analysis is None or self.source is None:
            messagebox.showinfo(
                self.t("choose_document_title"), self.t("choose_document_body"), parent=self.root
            )
            return
        if not self.confirmed.get():
            messagebox.showinfo(
                self.t("review_preview_title"), self.t("review_preview_body"), parent=self.root
            )
            return
        include = self.include_unmarked.get()
        excluded = [
            item for item in self.analysis.occurrences
            if item.excluded or (not item.highlighted and not include)
        ]
        confirmed_exclusions = False
        if excluded:
            confirmed_exclusions = messagebox.askyesno(
                self.t("confirm_retention_title"),
                self.t(
                    "confirm_retention_body", count=len(excluded),
                    highlighted=sum(item.highlighted for item in excluded),
                ), parent=self.root,
            )
            if not confirmed_exclusions:
                return
        password = self.password(confirm=True)
        if password is None:
            return
        chosen = filedialog.asksaveasfilename(
            title=self.t("save_redacted"), initialfile=self.t("redacted_filename"),
            defaultextension=".docx", filetypes=[(self.t("word_doc"), "*.docx")],
        )
        if not chosen:
            return
        output = Path(chosen)
        mapping = output.with_suffix(".wsmap")
        report = output.with_suffix(
            ".检查报告.json" if self.language == "zh" else ".inspection-report.json"
        )
        analysis, source = self.analysis, self.source
        acknowledged = self.ack.get()
        report_language = self.language

        def done(result):
            self.set_status("redaction_complete_status", count=result["replacement_count"])
            messagebox.showinfo(
                self.t("redaction_complete_title"), self.t("redaction_complete_body"),
                parent=self.root,
            )
        self.run(
            lambda: export_redacted(
                source, output, mapping, report, analysis, password,
                acknowledged=acknowledged, include_unmarked=include,
                confirmed_exclusions=confirmed_exclusions, report_language=report_language,
            ), done, "redacting",
        )

    def choose_final(self):
        value = filedialog.askopenfilename(
            title=self.t("choose_final"), filetypes=[(self.t("word_doc"), "*.docx")]
        )
        if value:
            self.final_path.set(value)

    def choose_mapping(self):
        value = filedialog.askopenfilename(
            title=self.t("choose_mapping"), filetypes=[(self.t("mapping_file"), "*.wsmap")]
        )
        if value:
            self.map_path.set(value)

    def restore(self):
        if not self.final_path.get() or not self.map_path.get():
            messagebox.showinfo(
                self.t("choose_files_title"), self.t("choose_files_body"), parent=self.root
            )
            return
        source, mapping = Path(self.final_path.get()), Path(self.map_path.get())

        def inspected(issues):
            self.restore_issues = issues
            if issues:
                self.render_restore_issues()
            if any(issue.blocking for issue in issues):
                messagebox.showerror(
                    self.t("must_fix_title"), self.t("must_fix_body"), parent=self.root
                )
                return
            acknowledged = False
            if issues:
                acknowledged = messagebox.askyesno(
                    self.t("confirm_review_title"), self.t("confirm_review_body"),
                    parent=self.root,
                )
                if not acknowledged:
                    return
            self.finish_restore(source, mapping, acknowledged)
        self.run(lambda: Package(source.read_bytes()).issues(), inspected, "inspecting_final")

    def finish_restore(self, source, mapping, acknowledged):
        password = self.password()
        if password is None:
            return
        chosen = filedialog.asksaveasfilename(
            title=self.t("save_restored"), initialfile=self.t("restored_filename"),
            defaultextension=".docx", filetypes=[(self.t("word_doc"), "*.docx")],
        )
        if not chosen:
            return
        output = Path(chosen)
        report_path = output.with_suffix(
            ".检查报告.json" if self.language == "zh" else ".inspection-report.json"
        )
        report_language = self.language

        def done(report):
            self.last_output = output
            self.last_restore_report = report
            self.restore_issues = []
            self.render_restore_report(report)
            self.set_status(
                "restore_complete_status",
                status=status_text(self.language, report["status"]),
            )
        self.run(
            lambda: export_restored(
                source, mapping, output, report_path,
                password, acknowledged=acknowledged, report_language=report_language,
            ), done, "restoring",
        )

    def render_restore_issues(self):
        self.set_text(
            self.restore_report,
            "\n".join(
                issue_text(self.language, issue.message)
                + f" ({part_text(self.language, issue.part)})"
                for issue in self.restore_issues
            ),
        )

    def render_restore_report(self, report):
        result_status = status_text(self.language, report["status"])
        placeholder_status = status_text(self.language, report["placeholder_status"])
        separator = "、" if self.language == "zh" else ", "
        lines = [
            self.t("restore_result", status=result_status),
            self.t("restored_count", count=report["restored_count"]), "",
            self.t("placeholder_check", status=placeholder_status),
            self.t(
                "structure_review",
                status=self.t("review_required" if report["document_review_required"] else "no_structure_warning"),
            ),
            self.t("unknown_tokens", value=separator.join(report["unknown_tokens"]) or self.t("none")),
            self.t("missing_tokens", value=separator.join(report["missing_tokens"]) or self.t("none")),
            self.t("count_changes"),
        ]
        lines.extend(
            self.t(
                "count_change", token=change["token"], expected=change["expected"],
                actual=change["actual"],
            ) for change in report["count_changes"]
        )
        lines.extend([
            self.t("malformed_count", count=len(report["malformed_tokens"])),
            self.t("remaining_count", count=len(report["remaining_tokens"])), "",
            self.t("package_valid"), notice_text(self.language, report["notice"]),
            self.t("report_saved"),
        ])
        self.set_text(self.restore_report, "\n".join(lines))

    def open_output(self):
        if self.last_output is None or not self.last_output.exists():
            messagebox.showinfo(
                self.t("no_output_title"), self.t("no_output_body"), parent=self.root
            )
            return
        if sys.platform == "darwin":
            subprocess.Popen(["/usr/bin/open", str(self.last_output)])
        else:
            messagebox.showinfo(self.t("no_output_title"), str(self.last_output), parent=self.root)

    def callback_error(self, *_):
        messagebox.showerror(
            self.t("ui_failed_title"), self.t("ui_failed_body"), parent=self.root
        )

    def close(self):
        if self.busy:
            messagebox.showinfo(
                self.t("busy_title"), self.t("busy_body"), parent=self.root
            )
            return
        self.pool.shutdown(wait=False, cancel_futures=True)
        self.analysis = None
        self.root.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
