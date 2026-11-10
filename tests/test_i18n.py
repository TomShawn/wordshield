"""Translation coverage and localized report tests."""
import ast
import json
import re

from tests.test_core import PASSWORD, simple
from wordshield.core import analyze, export_redacted, export_restored
from wordshield.i18n import (ERROR_EN, ISSUE_EN, TEXTS, error_text,
                             localized_report)


def test_english_catalog_values_do_not_contain_chinese_ui_copy():
    chinese = re.compile(r'[\u4e00-\u9fff]')
    assert all(len(pair) == 2 and all(pair) for pair in TEXTS.values())
    assert not any(chinese.search(pair[1]) for pair in TEXTS.values())
    assert not any(chinese.search(value) for value in ISSUE_EN.values())
    assert not any(chinese.search(value) for value in ERROR_EN.values())


def test_dynamic_error_and_report_translation():
    message = ('检测到已标记文字仍残留在 DOCX 内部，已阻止导出。'
               '请处理未标记位置、隐藏内容或内部属性。涉及部件数 3')
    assert error_text('en', message).endswith('Affected parts: 3')
    report = localized_report({
        'operation': '回填', 'status': '需人工检查',
        'placeholder_status': '通过',
        'notice': '消失的占位符仅列入报告，不猜测恢复位置。结构检查通过不等同于 Microsoft Word 实际打开验证。',
        'issues': [{'message': '存在文本框，需要在 Word 中人工审查。'}],
        'malformed_tokens': [{'kind': '疑似被改写或拆开的占位符'}],
    }, 'en')
    assert report['operation'] == 'Restoration'
    assert report['status'] == 'Manual review required'
    assert report['placeholder_status'] == 'Passed'
    assert report['issues'][0]['message'].startswith('A text box')
    assert report['malformed_tokens'][0]['kind'] == 'Possibly modified or split placeholder'


def test_english_export_writes_english_report(tmp_path):
    source = tmp_path / 'source.docx'
    output = tmp_path / 'redacted.docx'
    mapping = tmp_path / 'redacted.wsmap'
    report_path = tmp_path / 'redacted.inspection-report.json'
    source.write_bytes(simple([('林知远', True)]))
    export_redacted(
        source, output, mapping, report_path, analyze(source.read_bytes()), PASSWORD,
        report_language='en',
    )
    report = json.loads(report_path.read_text())
    assert report['operation'] == 'Redaction'
    assert report['notice'].startswith('Only manually marked text')
    assert output.exists() and mapping.exists()
    restored = tmp_path / 'restored.docx'
    restored_report_path = tmp_path / 'restored.inspection-report.json'
    export_restored(output, mapping, restored, restored_report_path, PASSWORD, report_language='en')
    restored_report = json.loads(restored_report_path.read_text())
    assert restored_report['operation'] == 'Restoration'
    assert restored_report['status'] == 'Passed'


def test_english_report_translates_generated_part_labels_without_changing_input():
    original = {
        'operation': '回填',
        'issues': [{'message': '存在文本框，需要在 Word 中人工审查。', 'part': '内部部件 012345abcdef'}],
        'remaining_tokens': [{'part': '内部部件 012345abcdef', 'tokens': ['【SECRET_999】']}],
        'excluded_locations': [{'part': 'word/header1.xml', 'start': 0, 'end': 3}],
    }
    translated = localized_report(original, 'en')
    assert translated['issues'][0]['part'] == 'Internal part 012345abcdef'
    assert translated['remaining_tokens'][0]['part'] == 'Internal part 012345abcdef'
    assert translated['excluded_locations'][0]['part'] == 'word/header1.xml'
    assert original['issues'][0]['part'] == '内部部件 012345abcdef'
    assert translated['remaining_tokens'][0]['tokens'] == ['【SECRET_999】']


def test_current_validation_errors_have_specific_english_translations():
    from wordshield.paths import PROJECT_ROOT
    for filename in ('core.py', 'crypto.py', 'storage.py', 'paths.py'):
        tree = ast.parse((PROJECT_ROOT / 'wordshield' / filename).read_text())
        for node in ast.walk(tree):
            if not isinstance(node, ast.Raise) or not isinstance(node.exc, ast.Call):
                continue
            arguments = node.exc.args
            if arguments and isinstance(arguments[0], ast.Constant):
                message = arguments[0].value
                if isinstance(message, str) and re.search(r'[\u4e00-\u9fff]', message):
                    assert message in ERROR_EN, message
