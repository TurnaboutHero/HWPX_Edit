"""
pytest-based pipeline automation test suite

Test data paths:
- D:\\Documents\\GitHub\\HWPX_Edit\\test_data\\hwpxlib\\testFile\\reader_writer\\ (25 .hwpx files)
- D:\\Documents\\GitHub\\HWPX_Edit\\test_data\\hwpxlib\\testFile\\error\\ (multi-section files)
- D:\\Documents\\GitHub\\HWPX_Edit\\3._(application)_2026_startup_success_package.hwpx

Run: cd D:\\Documents\\GitHub\\HWPX_Edit\\pipeline && python -m pytest tests/ -v
"""
import os
import sys
import re
import shutil
import uuid
import pytest
from pathlib import Path

# Project root paths
PROJECT_ROOT = Path(__file__).parent.parent.parent
PIPELINE_DIR = PROJECT_ROOT / "pipeline"
TEST_DATA_ROOT = PROJECT_ROOT / "test_data" / "hwpxlib" / "testFile"
READER_WRITER_DIR = TEST_DATA_ROOT / "reader_writer"
ERROR_DIR = TEST_DATA_ROOT / "error"
APPLICATION_FILE = PROJECT_ROOT / "3._(신청서)_2026년_창업성공패키지_청년창업사관학교_신청서.hwpx"

# Add module path
sys.path.insert(0, str(PIPELINE_DIR))

# Import modules
from hwpx_to_md import (
    HwpxToMarkdown,
    convert_hwpx_to_md,
    detect_namespace_version,
    NS_2011,
    NS_2024,
)
from smart_replace import (
    smart_replace,
    parse_markdown_tables,
    parse_markdown_paragraphs,
    collect_text_growth_warnings,
    collect_protected_markup_changes,
    validate_hwpx_integrity,
    detect_close_tag,
    detect_namespace_version as smart_detect_namespace,
    _normalize,
    _xml_escape,
    _strip_md_format,
    _compute_text_diffs,
)
from md_to_hwpx import _patch_hwpx


# ============================================================
# Test Fixtures
# ============================================================

@pytest.fixture
def tmp_output_dir():
    """Temporary output directory"""
    base_dir = PROJECT_ROOT / "work" / "pytest-output" / uuid.uuid4().hex
    output_dir = base_dir / "output"
    yield output_dir
    shutil.rmtree(base_dir, ignore_errors=True)


@pytest.fixture
def application_file():
    """Application file path (check existence)"""
    if not APPLICATION_FILE.exists():
        pytest.skip(f"Application file not found: {APPLICATION_FILE}")
    return APPLICATION_FILE


@pytest.fixture
def reader_writer_files():
    """All hwpx files in reader_writer directory"""
    if not READER_WRITER_DIR.exists():
        pytest.skip(f"reader_writer directory not found: {READER_WRITER_DIR}")

    files = list(READER_WRITER_DIR.glob("*.hwpx"))
    if not files:
        pytest.skip(f"No hwpx files in reader_writer directory: {READER_WRITER_DIR}")

    return files


@pytest.fixture
def error_dir_files():
    """All hwpx files in error directory"""
    if not ERROR_DIR.exists():
        pytest.skip(f"error directory not found: {ERROR_DIR}")

    files = list(ERROR_DIR.glob("*.hwpx"))
    if not files:
        pytest.skip(f"No hwpx files in error directory: {ERROR_DIR}")

    return files


# ============================================================
# hwpx_to_md.py Tests
# ============================================================

class TestHwpxToMd:
    """hwpx_to_md.py tests"""

    def test_application_regression(self, application_file, tmp_output_dir):
        """Application.hwpx -> 598 line markdown regression test"""
        tmp_output_dir.mkdir(parents=True, exist_ok=True)
        output_md = tmp_output_dir / "application.md"

        # Convert
        convert_hwpx_to_md(
            str(application_file),
            str(output_md),
            extract_images=True
        )

        # Check output file exists
        assert output_md.exists(), f"Markdown file not created: {output_md}"

        # Check line count
        with open(output_md, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            line_count = len(lines)

        # 598 lines ±5% tolerance (format changes possible)
        expected = 598
        tolerance = int(expected * 0.05)
        assert abs(line_count - expected) <= tolerance, \
            f"Line count mismatch: {line_count} (expected: {expected}±{tolerance})"

        # Check images directory
        images_dir = tmp_output_dir / "images"
        if images_dir.exists():
            image_files = list(images_dir.glob("*"))
            print(f"Images extracted: {len(image_files)}")

        # Check template_info.json
        template_info = tmp_output_dir / "template_info.json"
        assert template_info.exists(), "template_info.json not created"


    def test_reader_writer_batch(self, reader_writer_files, tmp_output_dir):
        """hwpxlib reader_writer 25 files batch no-error test"""
        tmp_output_dir.mkdir(parents=True, exist_ok=True)

        success_count = 0
        errors = []
        empty_files = []

        for hwpx_file in reader_writer_files:
            output_md = tmp_output_dir / f"{hwpx_file.stem}.md"

            try:
                # Convert (disable image extraction for speed)
                convert_hwpx_to_md(
                    str(hwpx_file),
                    str(output_md),
                    extract_images=False
                )

                # Check output file exists
                assert output_md.exists(), f"Markdown file not created: {output_md}"

                # Check if file is empty (some shape-only files may be empty)
                if output_md.stat().st_size == 0:
                    empty_files.append(hwpx_file.name)

                success_count += 1

            except Exception as e:
                errors.append((hwpx_file.name, str(e)))

        # Print results
        print(f"\nConversion success: {success_count}/{len(reader_writer_files)}")

        if empty_files:
            print(f"\nEmpty output files (shape-only documents): {len(empty_files)}")
            for filename in empty_files:
                print(f"  - {filename}")

        if errors:
            print("\nConversion failures:")
            for filename, error in errors:
                print(f"  - {filename}: {error}")

        # Check all succeeded (allow empty files for shape-only documents)
        assert success_count == len(reader_writer_files), \
            f"Some files failed: {len(errors)} failures"


    def test_namespace_detection(self):
        """detect_namespace_version() unit test"""
        # 2011 namespace
        xml_2011 = b'<?xml version="1.0"?><root xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"></root>'
        assert detect_namespace_version(xml_2011) == '2011'

        # 2024 namespace
        xml_2024 = b'<?xml version="1.0"?><root xmlns:p="http://www.owpml.org/owpml/2024/paragraph"></root>'
        assert detect_namespace_version(xml_2024) == '2024'

        # Default (2011)
        xml_unknown = b'<?xml version="1.0"?><root></root>'
        assert detect_namespace_version(xml_unknown) == '2011'


    def test_multi_section(self, error_dir_files, tmp_output_dir):
        """Multi-section file (3+ sections in error/) conversion test"""
        tmp_output_dir.mkdir(parents=True, exist_ok=True)

        # Find files with 3+ sections in error directory
        import zipfile

        multi_section_files = []
        for hwpx_file in error_dir_files:
            try:
                with zipfile.ZipFile(hwpx_file, 'r') as z:
                    section_files = [n for n in z.namelist() if re.match(r'Contents/section\d+\.xml', n)]
                    if len(section_files) >= 3:
                        multi_section_files.append((hwpx_file, len(section_files)))
            except:
                pass

        if not multi_section_files:
            pytest.skip("No files with 3+ sections")

        # Test first multi-section file
        test_file, section_count = multi_section_files[0]
        output_md = tmp_output_dir / f"{test_file.stem}_multi.md"

        # Convert
        convert_hwpx_to_md(
            str(test_file),
            str(output_md),
            extract_images=False
        )

        # Check output
        assert output_md.exists(), f"Markdown file not created: {output_md}"

        # Check --- separator (should be section_count - 1)
        with open(output_md, 'r', encoding='utf-8') as f:
            content = f.read()

        # Count --- on standalone lines
        separator_count = content.count('\n---\n')

        # Section separator should be (section_count - 1)
        expected_separators = section_count - 1
        assert separator_count == expected_separators, \
            f"Section separator mismatch: {separator_count} (expected: {expected_separators}, sections: {section_count})"


    def test_form_elements(self, reader_writer_files, tmp_output_dir):
        """Form elements (check/radio/combo/button/edit) conversion"""
        tmp_output_dir.mkdir(parents=True, exist_ok=True)

        # Target files
        target_files = {
            'SimpleButtons.hwpx': r'\[버튼:',
            'SimpleComboBox.hwpx': r'\[콤보:',
            'SimpleEdit.hwpx': r'\[입력란:',
        }

        found_files = {}
        for hwpx_file in reader_writer_files:
            if hwpx_file.name in target_files:
                found_files[hwpx_file.name] = hwpx_file

        if not found_files:
            pytest.skip("No form element test files found")

        # Convert and check pattern for each file
        for filename, pattern in target_files.items():
            if filename not in found_files:
                continue

            hwpx_file = found_files[filename]
            output_md = tmp_output_dir / f"{hwpx_file.stem}.md"

            # Convert
            convert_hwpx_to_md(
                str(hwpx_file),
                str(output_md),
                extract_images=False
            )

            # Check pattern
            with open(output_md, 'r', encoding='utf-8') as f:
                content = f.read()

            assert re.search(pattern, content), \
                f"{filename}: Pattern '{pattern}' not found"

            print(f"  ✓ {filename}: {pattern} confirmed")


    def test_dutmal(self, reader_writer_files, tmp_output_dir):
        """Dutmal (ruby) conversion - SimpleDutmal.hwpx"""
        tmp_output_dir.mkdir(parents=True, exist_ok=True)

        # Find SimpleDutmal.hwpx
        dutmal_file = None
        for hwpx_file in reader_writer_files:
            if hwpx_file.name == 'SimpleDutmal.hwpx':
                dutmal_file = hwpx_file
                break

        if not dutmal_file:
            pytest.skip("SimpleDutmal.hwpx not found")

        output_md = tmp_output_dir / "dutmal.md"

        # Convert
        convert_hwpx_to_md(
            str(dutmal_file),
            str(output_md),
            extract_images=False
        )

        # Check <ruby> tags
        with open(output_md, 'r', encoding='utf-8') as f:
            content = f.read()

        assert '<ruby>' in content and '<rt>' in content and '</rt></ruby>' in content, \
            "Dutmal <ruby> tags not found"

        print("  ✓ Dutmal conversion confirmed: <ruby> tags present")


    def test_equation(self, reader_writer_files, tmp_output_dir):
        """Equation conversion - SimpleEquation.hwpx"""
        tmp_output_dir.mkdir(parents=True, exist_ok=True)

        # Find SimpleEquation.hwpx
        equation_file = None
        for hwpx_file in reader_writer_files:
            if hwpx_file.name == 'SimpleEquation.hwpx':
                equation_file = hwpx_file
                break

        if not equation_file:
            pytest.skip("SimpleEquation.hwpx not found")

        output_md = tmp_output_dir / "equation.md"

        # Convert
        convert_hwpx_to_md(
            str(equation_file),
            str(output_md),
            extract_images=False
        )

        # Check $$ equation blocks
        with open(output_md, 'r', encoding='utf-8') as f:
            content = f.read()

        assert '$$' in content, "Equation block $$ not found"

        # Check equation block count (opening $$ and closing $$ pairs)
        equation_blocks = content.count('$$')
        assert equation_blocks >= 2 and equation_blocks % 2 == 0, \
            f"Equation block pair mismatch: {equation_blocks}"

        print(f"  ✓ Equation conversion confirmed: {equation_blocks // 2} equation blocks")


# ============================================================
# smart_replace.py Tests
# ============================================================

class TestSmartReplace:
    """smart_replace.py tests"""

    def test_parse_markdown_tables(self):
        """Markdown table parsing"""
        md_text = """
# Title

| A | B |
| --- | --- |
| 1 | 2 |
| 3 | 4 |

Normal text

> Quote (1x1 table)

| X |
| --- |
| Y |
"""
        tables = parse_markdown_tables(md_text)

        # 3 tables (2x2, quote, 1x1)
        assert len(tables) == 3, f"Table count mismatch: {len(tables)}"

        # First table (2x3 - header + 2 data rows)
        assert tables[0]['type'] == 'table'
        assert len(tables[0]['cells']) == 3  # header + 2 rows
        assert tables[0]['cells'][0] == ['A', 'B']
        assert tables[0]['cells'][1] == ['1', '2']
        assert tables[0]['cells'][2] == ['3', '4']

        # Second (quote)
        assert tables[1]['type'] == 'quote'
        assert tables[1]['cells'] == [['Quote (1x1 table)']]

        # Third table (1x2 - header + 1 data row)
        assert tables[2]['type'] == 'table'
        assert len(tables[2]['cells']) == 2
        assert tables[2]['cells'][0] == ['X']
        assert tables[2]['cells'][1] == ['Y']


    def test_parse_markdown_paragraphs_keeps_hwpx_line_mapping(self):
        """Consecutive HWPX paragraph lines must not be merged."""
        md_text = """**Title**
**Question**
**Answer**

| A |
| --- |
| B |

> Quote

**Next answer**
"""
        paragraphs = parse_markdown_paragraphs(md_text)

        assert paragraphs == [
            '**Title**',
            '**Question**',
            '**Answer**',
            '**Next answer**',
        ]


    def test_smart_replace_aborts_on_paragraph_count_mismatch(self, tmp_output_dir):
        """Structural mismatch should fail instead of writing corrupt output."""
        import zipfile

        tmp_output_dir.mkdir(parents=True, exist_ok=True)
        fake_hwpx = tmp_output_dir / "two_paragraphs.hwpx"
        edited_md = tmp_output_dir / "edited.md"
        output_hwpx = tmp_output_dir / "out.hwpx"

        section_xml = """<?xml version="1.0" encoding="UTF-8"?>
<hp:body xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p paraPrIDRef="1" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>one</hp:t></hp:run></hp:p>
  <hp:p paraPrIDRef="1" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>two</hp:t></hp:run></hp:p>
</hp:body>
"""
        with zipfile.ZipFile(fake_hwpx, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('mimetype', 'application/hwp+zip', compress_type=zipfile.ZIP_STORED)
            z.writestr('Contents/section0.xml', section_xml)
        edited_md.write_text("one two\n", encoding='utf-8')

        with pytest.raises(ValueError, match="문단 수"):
            smart_replace(str(fake_hwpx), str(edited_md), str(output_hwpx))

        assert not output_hwpx.exists()


    def test_smart_replace_line_mapping_and_strip_lineseg(self, tmp_output_dir):
        """Line-mapped paragraph replacement should update text and drop layout cache."""
        import zipfile

        tmp_output_dir.mkdir(parents=True, exist_ok=True)
        fake_hwpx = tmp_output_dir / "line_mapping.hwpx"
        edited_md = tmp_output_dir / "edited.md"
        output_hwpx = tmp_output_dir / "out.hwpx"

        section_xml = """<?xml version="1.0" encoding="UTF-8"?>
<hp:body xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p paraPrIDRef="1" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>one</hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0"/></hp:linesegarray></hp:p>
  <hp:p paraPrIDRef="1" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>two</hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0"/></hp:linesegarray></hp:p>
</hp:body>
"""
        with zipfile.ZipFile(fake_hwpx, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('mimetype', 'application/hwp+zip', compress_type=zipfile.ZIP_STORED)
            z.writestr('Contents/section0.xml', section_xml)
        edited_md.write_text("one\nsecond paragraph\n", encoding='utf-8')

        smart_replace(str(fake_hwpx), str(edited_md), str(output_hwpx))

        with zipfile.ZipFile(output_hwpx, 'r') as z:
            patched_xml = z.read('Contents/section0.xml').decode('utf-8')

        assert '<hp:t>second paragraph</hp:t>' in patched_xml
        assert '<hp:linesegarray>' not in patched_xml

        validation = validate_hwpx_integrity(str(output_hwpx), require_no_lineseg=True)
        assert validation['ok'] is True
        assert validation['section_count'] == 1
        assert validation['linesegarray_count'] == 0


    def test_collect_text_growth_warnings(self):
        """Large text growth should be reported as layout risk."""
        warnings = collect_text_growth_warnings(
            ['짧은 답'],
            ['이 답변은 기존보다 훨씬 길어져서 원래 답안 줄 안에 들어가지 않을 수 있습니다. '
             '따라서 레이아웃 주의 항목으로 잡혀야 합니다.'],
            kind='문단',
            ratio_threshold=1.2,
            delta_threshold=10,
            width_threshold=20,
        )

        assert len(warnings) == 1
        assert warnings[0]['kind'] == '문단'
        assert warnings[0]['new_width'] > warnings[0]['old_width']


    def test_collect_protected_markup_changes(self):
        """Unsupported markup edits should be detected before conversion."""
        changes = collect_protected_markup_changes(
            "# Title\n![img](images/img)\nanswer\n",
            "# Changed\n![other](images/other)\nanswer\n",
        )

        assert len(changes) == 2
        assert changes[0]['old'] == '# Title'
        assert changes[0]['new'] == '# Changed'

        path_only_changes = collect_protected_markup_changes(
            "![image1](images/image1)\nanswer\n",
            "![image1](images/image1.png)\nanswer\n",
        )
        assert path_only_changes == []


    def test_smart_replace_blocks_layout_risk_by_default(self, tmp_output_dir):
        """Large text growth requires an explicit override."""
        import zipfile

        tmp_output_dir.mkdir(parents=True, exist_ok=True)
        fake_hwpx = tmp_output_dir / "layout_risk.hwpx"
        edited_md = tmp_output_dir / "edited.md"
        output_hwpx = tmp_output_dir / "out.hwpx"

        section_xml = """<?xml version="1.0" encoding="UTF-8"?>
<hp:body xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p paraPrIDRef="1" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>짧은 답</hp:t></hp:run></hp:p>
</hp:body>
"""
        long_answer = (
            "이 답변은 기존보다 훨씬 길어져서 원래 답안 줄 안에 들어가지 않을 수 있습니다. "
            "따라서 기본 정책에서는 파일 생성을 중단해야 합니다. "
            "사용자가 명시적으로 허용하지 않는 한 긴 문장으로 인한 페이지 흐름 변경을 막아야 합니다. "
            "이 문장은 테스트 임계값을 넘기기 위해 충분히 길게 작성되었습니다."
        )
        with zipfile.ZipFile(fake_hwpx, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('mimetype', 'application/hwp+zip', compress_type=zipfile.ZIP_STORED)
            z.writestr('Contents/section0.xml', section_xml)
        edited_md.write_text(long_answer + "\n", encoding='utf-8')

        with pytest.raises(ValueError, match="레이아웃 위험"):
            smart_replace(str(fake_hwpx), str(edited_md), str(output_hwpx))

        assert not output_hwpx.exists()

        smart_replace(
            str(fake_hwpx),
            str(edited_md),
            str(output_hwpx),
            allow_layout_risk=True,
        )
        assert output_hwpx.exists()


    def test_smart_replace_blocks_protected_markup_edits(self, tmp_output_dir):
        """Heading/image/form-like lines are not silently ignored."""
        import zipfile

        tmp_output_dir.mkdir(parents=True, exist_ok=True)
        fake_hwpx = tmp_output_dir / "heading.hwpx"
        edited_md = tmp_output_dir / "edited.md"
        output_hwpx = tmp_output_dir / "out.hwpx"

        header_xml = """<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:paraPr id="1"><hh:heading type="OUTLINE" level="0"/></hh:paraPr>
  <hh:charPr id="0"><hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/></hh:charPr>
</hh:head>
"""
        section_xml = """<?xml version="1.0" encoding="UTF-8"?>
<hp:body xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p paraPrIDRef="1" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>Original title</hp:t></hp:run></hp:p>
  <hp:p paraPrIDRef="2" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>body</hp:t></hp:run></hp:p>
</hp:body>
"""
        with zipfile.ZipFile(fake_hwpx, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('mimetype', 'application/hwp+zip', compress_type=zipfile.ZIP_STORED)
            z.writestr('Contents/header.xml', header_xml)
            z.writestr('Contents/section0.xml', section_xml)
        edited_md.write_text("# Changed title\nbody\n", encoding='utf-8')

        with pytest.raises(ValueError, match="이미지, 제목"):
            smart_replace(str(fake_hwpx), str(edited_md), str(output_hwpx))

        assert not output_hwpx.exists()


    def test_detect_close_tag(self):
        """Dynamic namespace prefix detection"""
        # 2011 style
        xml_2011 = '<hp:run><hp:t>text</hp:t></hp:run>'
        assert detect_close_tag(xml_2011) == '</hp:t>'

        # 2024 style (hypothetical)
        xml_2024 = '<p:run><p:t>text</p:t></p:run>'
        assert detect_close_tag(xml_2024) == '</p:t>'

        # Default
        xml_no_t = '<hp:run>text</hp:run>'
        assert detect_close_tag(xml_no_t) == '</hp:t>'


    def test_detect_namespace_version(self):
        """2011 vs 2024 namespace detection"""
        xml_2011 = b'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"'
        assert smart_detect_namespace(xml_2011) == '2011'

        xml_2024 = b'xmlns:p="http://www.owpml.org/owpml/2024/paragraph"'
        assert smart_detect_namespace(xml_2024) == '2024'


    def test_normalize(self):
        """Text normalization"""
        # Whitespace normalization
        assert _normalize('  multiple   spaces  ') == 'multiple spaces'

        # Line break normalization
        assert _normalize('line\nbreak\r\ntest') == 'line break test'

        # Footnote marker removal
        assert _normalize('text*') == 'text'
        assert _normalize('multiple* footnote* markers*') == 'multiple footnote markers'


    def test_xml_escape(self):
        """XML text node escaping"""
        assert _xml_escape('A & B') == 'A &amp; B'
        assert _xml_escape('<tag>') == '&lt;tag&gt;'
        assert _xml_escape('A < B > C & D') == 'A &lt; B &gt; C &amp; D'

        # Double escape check
        assert _xml_escape('&amp;') == '&amp;amp;'


    def test_strip_md_format(self):
        """Markdown format removal"""
        # Bold
        assert _strip_md_format('**bold**') == 'bold'

        # Italic
        assert _strip_md_format('*italic*') == 'italic'

        # Bold + Italic
        assert _strip_md_format('***bolditalic***') == 'bolditalic'

        # Strikethrough
        assert _strip_md_format('~~strikethrough~~') == 'strikethrough'

        # Line break
        assert _strip_md_format('line<br>break') == 'line break'

        # Escaped pipe
        assert _strip_md_format('A \\| B') == 'A | B'

        # Combined
        assert _strip_md_format('**bold** *italic* ~~strike~~') == 'bold italic strike'


    def test_compute_text_diffs(self):
        """Fragment diff"""
        # Simple replace
        old = "Hello World"
        new = "Hello Python"
        diffs = _compute_text_diffs(old, new)
        # difflib may split into smaller chunks
        assert len(diffs) > 0, "No diffs found"
        # Check that some replacement occurred
        assert any('W' in old_frag or 'o' in old_frag for old_frag, _ in diffs)

        # Multiple replacements
        old = "A B C D"
        new = "A X C Y"
        diffs = _compute_text_diffs(old, new)
        assert len(diffs) == 2
        assert ('B', 'X') in diffs
        assert ('D', 'Y') in diffs

        # No changes
        old = new = "same text"
        diffs = _compute_text_diffs(old, new)
        assert len(diffs) == 0


# ============================================================
# md_to_hwpx.py Tests
# ============================================================

class TestMdToHwpx:
    """md_to_hwpx.py tests"""

    def test_patch_empty_sublist(self, tmp_output_dir):
        """Empty subList patch test (regex)"""
        import zipfile
        import io

        tmp_output_dir.mkdir(parents=True, exist_ok=True)

        # Create fake HWPX with empty subList
        fake_section_xml = """<?xml version="1.0" encoding="UTF-8"?>
<hp:body xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p paraPrIDRef="1" styleIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>text</hp:t>
      <hp:ctrl>
        <hp:footnote>
          <hp:subList></hp:subList>
        </hp:footnote>
      </hp:ctrl>
    </hp:run>
  </hp:p>
</hp:body>
"""

        fake_hwpx = tmp_output_dir / "fake.hwpx"

        # Create ZIP
        with zipfile.ZipFile(fake_hwpx, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('mimetype', 'application/hwp+zip', compress_type=zipfile.ZIP_STORED)
            z.writestr('Contents/section0.xml', fake_section_xml)

        # Apply patch
        result = _patch_hwpx(str(fake_hwpx))

        # Check patch success
        assert result is True, "Patch failed"

        # Check patched XML
        with zipfile.ZipFile(fake_hwpx, 'r') as z:
            patched_xml = z.read('Contents/section0.xml').decode('utf-8')

        # Empty subList should be gone
        assert '<hp:subList></hp:subList>' not in patched_xml, "Empty subList still exists"

        # Check empty paragraph was injected
        assert '<hp:subList>' in patched_xml, "subList disappeared"
        assert '<hp:p paraPrIDRef=' in patched_xml, "Empty paragraph injection failed"
        assert '</hp:subList>' in patched_xml, "subList closing tag disappeared"

        print("  ✓ Empty subList patch confirmed: empty paragraph injected")


# ============================================================
# Run pytest when executed directly
# ============================================================

if __name__ == '__main__':
    pytest.main([__file__, '-v'])
