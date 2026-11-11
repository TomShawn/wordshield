"""使用虚构材料验证界面正常入口，系统需提供 Tk 图形会话。"""
import time
import tkinter as tk
import json
import threading

from scripts.generate_fixtures import PASSWORD, highlighted, saved
from docx import Document
from wordshield.gui import App


def test_gui_redact_and_restore_normal_entry(tmp_path, monkeypatch):
    doc = Document()
    highlighted(doc.add_paragraph(), "林知远")
    source = tmp_path / "fictional.docx"
    source.write_bytes(saved(doc))
    redacted = tmp_path / "redacted.docx"
    restored = tmp_path / "restored.docx"
    root = tk.Tk()
    app = App(root)
    messages = []
    from wordshield import gui
    monkeypatch.setattr(gui.messagebox, "showinfo", lambda *a, **k: messages.append(a))
    monkeypatch.setattr(gui.messagebox, "showerror", lambda *a, **k: messages.append(("ERROR", *a)))
    monkeypatch.setattr(gui.filedialog, "askopenfilename", lambda **k: str(source))
    monkeypatch.setattr(gui.filedialog, "asksaveasfilename", lambda **k: str(redacted))
    monkeypatch.setattr(app, "password", lambda **k: PASSWORD)
    def settle():
        deadline = time.monotonic() + 20
        while app.busy and time.monotonic() < deadline:
            root.update()
            time.sleep(0.01)
        assert not app.busy
        assert not any(m[0] == "ERROR" for m in messages)
    try:
        assert app.language == 'en'
        assert app.language_choice.get() == 'English'
        assert app.tabs.tab(app.redaction_frame, 'text') == 'Redact Word'
        assert app.status.get() == 'Choose a DOCX file to begin.'
        app.language_choice.set('中文')
        app.language_box.event_generate('<<ComboboxSelected>>')
        root.update()
        app.choose_source()
        settle()
        assert len(app.tree.get_children()) == 1
        app.confirmed.set(True)
        app.redact()
        settle()
        assert redacted.exists() and redacted.with_suffix(".wsmap").exists()
        app.final_path.set(str(redacted))
        app.map_path.set(str(redacted.with_suffix(".wsmap")))
        monkeypatch.setattr(gui.filedialog, "asksaveasfilename", lambda **k: str(restored))
        app.restore()
        settle()
        assert restored.exists()
        assert "林知远" in Document(restored).paragraphs[0].text
        assert "通过" in app.restore_report.get("1.0", "end")
        root.update_idletasks()
        # 最小支持窗口下主操作必须可见。
        root.geometry("900x720")
        root.update()
        export_button = next(b for b in app.buttons if b.cget("text").startswith("设置密码"))
        assert export_button.winfo_rooty() + export_button.winfo_height() < root.winfo_rooty() + root.winfo_height()
        check_unmarked_retention(app, tmp_path, monkeypatch, settle)
        check_punctuation_retention(app, monkeypatch)
        check_group_actions(app, monkeypatch)
        check_language_switch(app)
        check_english_round_trip(app, tmp_path, monkeypatch, settle)
        check_restore_report_lifecycle(app, tmp_path, monkeypatch)
    finally:
        app.close()


def check_unmarked_retention(app, tmp_path, monkeypatch, settle):
    from wordshield import gui
    from tests.test_core import simple
    source = tmp_path / 'fictional_retention.docx'
    source.write_bytes(simple([('林知远', True), ('另有林知远', False)]))
    prompts = []
    monkeypatch.setattr(gui.filedialog, 'askopenfilename', lambda **k: str(source))
    for accepted in (False, True):
        output = tmp_path / f'fictional_retained_{accepted}.docx'
        def confirm(*args, **kwargs):
            prompts.append(args)
            return accepted
        monkeypatch.setattr(gui.messagebox, 'askyesno', confirm)
        monkeypatch.setattr(gui.filedialog, 'asksaveasfilename', lambda **k: str(output))
        app.choose_source()
        settle()
        app.include_unmarked.set(False)
        app.refresh()
        item = next(o for o in app.analysis.occurrences if not o.highlighted)
        assert '保留' in app.tree.item(item.key, 'values')[2]
        app.confirmed.set(True)
        app.redact()
        settle()
        assert '保留原文' in prompts[-1][0]
        assert output.exists() == accepted
        if accepted:
            assert Document(output).paragraphs[0].text == '【SECRET_001】另有林知远'
    assert len(prompts) == 2


def check_punctuation_retention(app, monkeypatch):
    from wordshield import gui
    from tests.test_core import simple
    from wordshield.core import analyze
    app.loaded(analyze(simple([('林知远、周映禾', True), ('。公开甲、公开乙', False)])))
    item = app.analysis.occurrences[0]
    app.tree.selection_set(item.key)
    monkeypatch.setattr(gui.simpledialog, 'askstring', lambda *a, **k: '3,4')
    app.split()
    punctuation = next(o for o in app.analysis.occurrences if o.highlighted and o.text == '、')
    app.tree.selection_set(punctuation.key)
    app.toggle_exclusion()
    assert next(o for o in app.analysis.occurrences if o.key == punctuation.key).excluded
    assert '保留' in app.tree.item(punctuation.key, 'values')[2]
    assert not any(not o.highlighted and o.text == '、' for o in app.analysis.occurrences)
    assert not app.confirmed.get()


def check_group_actions(app, monkeypatch):
    from wordshield import gui
    from tests.test_core import simple
    from wordshield.core import analyze
    app.language_choice.set('English')
    app.language_box.event_generate('<<ComboboxSelected>>')
    app.root.update()
    app.loaded(analyze(simple([('林知远', True), ('、林知远、林知远', False)])))
    highlighted, retained, separate = app.analysis.occurrences
    app.tree.selection_set(separate.key)
    app.separate_group()
    independent_group = separate.group
    app.tree.selection_set(retained.key)
    app.toggle_exclusion()
    retained = next(item for item in app.analysis.occurrences if item.key == retained.key)
    separate = next(item for item in app.analysis.occurrences if item.key == separate.key)
    app.tree.selection_set(highlighted.key)
    prompts = []
    def rename_prompt(title, prompt, **options):
        prompts.append((title, prompt))
        return 'Representative'
    monkeypatch.setattr(gui.simpledialog, 'askstring', rename_prompt)
    app.confirmed.set(True)
    app.rename_group()
    assert prompts[0] == ('Rename group', 'Enter a new name for all 2 items in this group.')
    assert highlighted.group == retained.group == 'Representative'
    assert separate.group == independent_group
    assert retained.excluded
    assert not app.confirmed.get()
    app.tree.selection_set(highlighted.key)
    monkeypatch.setattr(gui.simpledialog, 'askstring', lambda *args, **kwargs: 'Assigned item')
    app.edit_group()
    assert highlighted.group == 'Assigned item'
    assert retained.group == 'Representative'
    assert separate.group == independent_group


def check_language_switch(app):
    from wordshield.core import Issue
    before = [(item.key, item.text, item.group, item.excluded)
              for item in app.analysis.occurrences]
    selected = app.analysis.occurrences[0]
    app.tree.selection_set(selected.key)
    app.confirmed.set(True)
    app.ack.set(True)
    include_unmarked = app.include_unmarked.get()
    app.language_choice.set('English')
    app.language_box.event_generate('<<ComboboxSelected>>')
    app.root.update()
    assert app.language == 'en'
    assert app.root.title() == 'WordShield — Offline Word Redaction & Restoration'
    assert app.tabs.tab(app.redaction_frame, 'text') == 'Redact Word'
    assert app.tree.heading('origin', 'text') == 'Source / action'
    assert any(button.cget('text') == 'Choose DOCX and analyze' for button in app.buttons)
    assert 'Yellow highlight' in app.tree.item(app.tree.get_children()[0], 'values')[2]
    assert 'Export removes document properties' in app.risk.get('1.0', 'end')
    assert app.confirmed.get()  # Language switching must not discard user confirmation.
    assert app.ack.get()
    assert app.include_unmarked.get() == include_unmarked
    assert app.tree.selection() == (selected.key,)
    assert [(item.key, item.text, item.group, item.excluded)
            for item in app.analysis.occurrences] == before
    report = {
        'status': '通过', 'placeholder_status': '通过', 'restored_count': 2,
        'document_review_required': False, 'unknown_tokens': [], 'missing_tokens': [],
        'count_changes': [], 'malformed_tokens': [], 'remaining_tokens': [],
        'notice': '消失的占位符仅列入报告，不猜测恢复位置。结构检查通过不等同于 Microsoft Word 实际打开验证。',
    }
    app.last_restore_report = report
    app.render_restore_report(report)
    app.set_status('restore_complete_status', status='Passed')
    assert 'Restoration result: Passed' in app.restore_report.get('1.0', 'end')
    app.language_choice.set('中文')
    app.language_box.event_generate('<<ComboboxSelected>>')
    app.root.update()
    assert app.tabs.tab(app.redaction_frame, 'text') == '脱敏 Word'
    assert '结果为通过' in app.status.get()
    # A rejected structural review must remain visible when changing language.
    app.restore_issues = [Issue('TEXTBOX', '存在文本框，需要在 Word 中人工审查。', 'word/document.xml')]
    for language in ('中文', 'English'):
        app.language_choice.set(language)
        app.language_box.event_generate('<<ComboboxSelected>>')
        app.root.update()
        for size in ('900x720', '1080x800'):
            app.root.geometry(size)
            for tab in (app.redaction_frame, app.restore_frame):
                app.tabs.select(tab)
                app.root.update()
                for widget, _ in app.localized_widgets:
                    if str(widget).startswith(str(tab) + '.'):
                        assert widget.winfo_ismapped()
                    if widget.winfo_ismapped():
                        assert widget.winfo_width() >= widget.winfo_reqwidth()
                        assert widget.winfo_rootx() + widget.winfo_width() <= app.root.winfo_rootx() + app.root.winfo_width()
                        assert widget.winfo_rooty() + widget.winfo_height() <= app.root.winfo_rooty() + app.root.winfo_height()
                assert app.language_box.winfo_width() >= app.language_box.winfo_reqwidth()
        expected = 'A text box is present' if language == 'English' else '存在文本框'
        assert expected in app.restore_report.get('1.0', 'end')
    app.restore_issues = []


def check_english_round_trip(app, tmp_path, monkeypatch, settle):
    from wordshield import gui
    from tests.test_core import simple
    source = tmp_path / 'english-source.docx'
    redacted = tmp_path / 'english-redacted.docx'
    restored = tmp_path / 'english-restored.docx'
    source.write_bytes(simple([('Fictional Person', True)]))
    app.language_choice.set('English')
    app.language_box.event_generate('<<ComboboxSelected>>')
    monkeypatch.setattr(gui.filedialog, 'askopenfilename', lambda **k: str(source))
    save_options = []
    def save_redacted(**options):
        save_options.append(options)
        return str(redacted)
    monkeypatch.setattr(gui.filedialog, 'asksaveasfilename', save_redacted)
    app.choose_source()
    settle()
    app.confirmed.set(True)
    app.redact()
    settle()
    assert save_options[0]['initialfile'] == 'redacted_document.docx'
    redaction_report = json.loads(redacted.with_suffix('.inspection-report.json').read_text())
    assert redaction_report['operation'] == 'Redaction'
    app.final_path.set(str(redacted))
    app.map_path.set(str(redacted.with_suffix('.wsmap')))
    def save_restored(**options):
        save_options.append(options)
        return str(restored)
    monkeypatch.setattr(gui.filedialog, 'asksaveasfilename', save_restored)
    app.restore()
    settle()
    assert save_options[1]['initialfile'] == 'restored_document.docx'
    restoration_report = json.loads(restored.with_suffix('.inspection-report.json').read_text())
    assert restoration_report['status'] == 'Passed'
    assert Document(restored).paragraphs[0].text == 'Fictional Person'
    assert 'Restoration result: Passed' in app.restore_report.get('1.0', 'end')
    app.language_choice.set('中文')
    app.language_box.event_generate('<<ComboboxSelected>>')
    app.root.update()
    assert '回填检查结果 通过' in app.restore_report.get('1.0', 'end')


def check_restore_report_lifecycle(app, tmp_path, monkeypatch):
    """Reuse the Tk session to cover changed inputs, failures, cancellation and late results."""
    from wordshield import gui
    from wordshield.core import Issue, SafetyError
    from wordshield.crypto import MappingError

    source = tmp_path / 'english-redacted.docx'
    mapping = source.with_suffix('.wsmap')
    successful_report = app.last_restore_report
    assert successful_report is not None

    def seed_success():
        app.last_restore_report = successful_report
        app.restore_issues = []
        app.render_restore_report(successful_report)
        app.set_status('restore_complete_status', status=gui.status_text(app.language, '通过'))

    def assert_pending():
        assert app.last_restore_report is None
        assert not app.restore_issues
        assert app.restore_report.get('1.0', 'end').strip() == app.t('restore_pending')
        assert app.status_key != 'restore_complete_status'

    def wait_for_idle():
        deadline = time.monotonic() + 20
        while app.busy and time.monotonic() < deadline:
            app.root.update()
            time.sleep(0.01)
        assert not app.busy

    with monkeypatch.context() as patch:
        errors = []
        patch.setattr(gui.messagebox, 'showerror', lambda *a, **k: errors.append(a))
        patch.setattr(gui.messagebox, 'showinfo', lambda *a, **k: None)
        patch.setattr(app, 'password', lambda **k: PASSWORD)

        # Both native file-selection callbacks and direct entry edits invalidate results.
        for variable, choose, value in (
            (app.final_path, app.choose_final, str(source)),
            (app.map_path, app.choose_mapping, str(mapping)),
        ):
            seed_success()
            patch.setattr(gui.filedialog, 'askopenfilename', lambda **k: value)
            choose()
            assert_pending()
            seed_success()
            variable.set('')
            assert_pending()
            variable.set(value)
        app.restore_issues = [Issue('TEXTBOX', '存在文本框，需要在 Word 中人工审查。', 'word/document.xml')]
        app.map_path.set(str(mapping))
        assert_pending()

        # A same-input retry invalidates success before the worker starts.
        seed_success()
        retry_output = tmp_path / 'report-retry.docx'
        patch.setattr(gui.filedialog, 'asksaveasfilename', lambda **k: str(retry_output))
        app.restore()
        assert_pending()
        wait_for_idle()
        assert retry_output.exists()
        assert app.last_restore_report['status'] == '通过'
        assert not errors

        # Each background error category leaves no previous success to resurrect.
        original_package = gui.Package
        for error in (OSError('fictional I/O failure'), SafetyError('fictional safety failure'),
                      MappingError('fictional mapping failure'), RuntimeError('fictional failure')):
            def fail_package(_raw, error=error):
                raise error
            patch.setattr(gui, 'Package', fail_package)
            seed_success()
            app.restore()
            wait_for_idle()
            assert_pending()
            assert errors
            for language in ('English', '中文'):
                app.language_choice.set(language)
                app.language_box.event_generate('<<ComboboxSelected>>')
                app.root.update()
                assert_pending()
            errors.clear()
        patch.setattr(gui, 'Package', original_package)

        # A real decryption failure also clears the previous successful report.
        seed_success()
        wrong_password_output = tmp_path / 'wrong-password.docx'
        patch.setattr(app, 'password', lambda **k: 'fictional-wrong-password')
        patch.setattr(gui.filedialog, 'asksaveasfilename', lambda **k: str(wrong_password_output))
        app.restore()
        wait_for_idle()
        assert_pending()
        assert errors and not wrong_password_output.exists()
        errors.clear()
        patch.setattr(app, 'password', lambda **k: PASSWORD)

        # Missing inputs and cancelled dialogs cannot leave a successful result visible.
        app.final_path.set('')
        seed_success()
        app.restore()
        assert_pending()
        app.final_path.set(str(source))
        for cancel_password in (True, False):
            seed_success()
            patch.setattr(app, 'password', lambda **k: None if cancel_password else PASSWORD)
            patch.setattr(gui.filedialog, 'asksaveasfilename', lambda **k: '')
            app.restore()
            wait_for_idle()
            assert_pending()
        patch.setattr(app, 'password', lambda **k: PASSWORD)

        # Editing an input during inspection discards its late callback and password prompt.
        entered, resume = threading.Event(), threading.Event()
        prompts = []
        def inspect_later(raw):
            entered.set()
            assert resume.wait(10)
            return original_package(raw)
        patch.setattr(gui, 'Package', inspect_later)
        patch.setattr(app, 'password', lambda **k: prompts.append(True) or PASSWORD)
        app.restore()
        try:
            assert entered.wait(5)
            app.final_path.set(str(tmp_path / 'different-final.docx'))
        finally:
            resume.set()
        wait_for_idle()
        assert_pending()
        assert not prompts
        patch.setattr(gui, 'Package', original_package)
        app.final_path.set(str(source))

        # An export already underway may finish, but its result belongs to its captured inputs.
        original_export = gui.export_restored
        entered, resume = threading.Event(), threading.Event()
        def export_later(*args, **kwargs):
            entered.set()
            assert resume.wait(10)
            return original_export(*args, **kwargs)
        patch.setattr(gui, 'export_restored', export_later)
        late_output = tmp_path / 'late-restoration.docx'
        patch.setattr(gui.filedialog, 'asksaveasfilename', lambda **k: str(late_output))
        app.restore()
        deadline = time.monotonic() + 10
        try:
            while not entered.is_set() and time.monotonic() < deadline:
                app.root.update()
                time.sleep(0.01)
            assert entered.is_set()
            app.map_path.set(str(tmp_path / 'different-mapping.wsmap'))
        finally:
            resume.set()
        wait_for_idle()
        assert late_output.exists()
        assert app.last_output == late_output
        assert_pending()
        assert not errors

        # A subsequent successful retry displays a current report in both languages.
        patch.setattr(gui, 'export_restored', original_export)
        app.map_path.set(str(mapping))
        retry_output = tmp_path / 'current-restoration.docx'
        patch.setattr(gui.filedialog, 'asksaveasfilename', lambda **k: str(retry_output))
        app.restore()
        wait_for_idle()
        for language in ('English', '中文'):
            app.language_choice.set(language)
            app.language_box.event_generate('<<ComboboxSelected>>')
            app.root.update()
            assert app.last_restore_report['status'] == '通过'
            assert app.t('restore_result', status=gui.status_text(app.language, '通过')) in app.restore_report.get('1.0', 'end')
        assert not errors
