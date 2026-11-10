"""项目目录只读与发布清单的回归检查。"""
from pathlib import Path
import sys
import subprocess

import pytest

from scripts.generate_fixtures import make_all
from wordshield.paths import PROJECT_ROOT, development_directory, require_external
from wordshield.storage import OutputError, write_outputs


def test_outputs_cannot_be_saved_inside_source_checkout():
    target = PROJECT_ROOT / 'forbidden-output.docx'
    with pytest.raises(OutputError, match='目录之外'):
        write_outputs({target: b'fictional'})
    assert not target.exists()


def test_symlink_to_checkout_does_not_bypass_output_guard(tmp_path):
    alias = tmp_path / 'source-alias'
    alias.symlink_to(PROJECT_ROOT, target_is_directory=True)
    with pytest.raises(OutputError, match='目录之外'):
        write_outputs({alias / 'forbidden-output.docx': b'fictional'})
    assert not (PROJECT_ROOT / 'forbidden-output.docx').exists()


def test_frozen_application_installation_is_protected(tmp_path, monkeypatch):
    installation = tmp_path / 'downloaded-app'
    executable = installation / 'WordShield.app' / 'Contents' / 'MacOS' / 'WordShield'
    monkeypatch.setattr(sys, 'frozen', True, raising=False)
    monkeypatch.setattr(sys, 'executable', str(executable))
    with pytest.raises(ValueError, match='目录之外'):
        require_external(installation / 'result.docx')
    assert require_external(tmp_path / 'documents' / 'result.docx') == tmp_path / 'documents' / 'result.docx'


def test_development_directory_is_external(monkeypatch):
    monkeypatch.delenv('WORDSHIELD_DEV_HOME', raising=False)
    assert not development_directory().is_relative_to(PROJECT_ROOT)
    monkeypatch.setenv('WORDSHIELD_DEV_HOME', str(PROJECT_ROOT / 'cache'))
    with pytest.raises(ValueError, match='目录之外'):
        development_directory()


def test_generator_refuses_to_create_fixtures_in_project():
    with pytest.raises(ValueError, match='目录之外'):
        make_all(PROJECT_ROOT / 'forbidden-fixtures')
    assert not (PROJECT_ROOT / 'forbidden-fixtures').exists()


def test_gitignore_excludes_generated_and_confidential_files(tmp_path):
    (tmp_path / '.gitignore').write_bytes((PROJECT_ROOT / '.gitignore').read_bytes())
    subprocess.run(['git', 'init', '-q', str(tmp_path)], check=True)
    paths = ['.venv/data', '__pycache__/module.pyc', '.pytest_cache/state',
             'build/output', 'dist/WordShield.app/Contents/data',
             'fixtures/fictional.docx', 'private.docx', 'private.wsmap',
             'private.检查报告.json', 'private.inspection-report.json', 'debug.log', '.env', '.DS_Store']
    result = subprocess.run(['git', '-c', 'core.quotePath=false', '-C', str(tmp_path), 'check-ignore', '--no-index', '--stdin'],
                            input='\n'.join(paths) + '\n', text=True, capture_output=True, check=True)
    assert set(result.stdout.splitlines()) == set(paths)


def test_english_finder_launcher_is_executable():
    launcher = PROJECT_ROOT / 'Launch WordShield.command'
    assert launcher.is_file()
    assert launcher.stat().st_mode & 0o111
    assert not (PROJECT_ROOT / '启动应用.command').exists()


def test_frozen_application_also_protects_source_checkout_elsewhere(tmp_path, monkeypatch):
    checkout = tmp_path / 'github' / 'wordshield'
    (checkout / 'wordshield').mkdir(parents=True)
    (checkout / 'pyproject.toml').write_text('[project]\nname="wordshield-offline"\n')
    monkeypatch.setattr(sys, 'frozen', True, raising=False)
    monkeypatch.setattr(sys, 'executable', str(tmp_path / 'installed' / 'WordShield.app' / 'Contents' / 'MacOS' / 'WordShield'))
    with pytest.raises(ValueError, match='目录之外'):
        require_external(checkout / 'output.docx')
