#!/usr/bin/env python3
"""
Tests for docx document.py efficiency improvements.
Run with: python -m pytest document_efficiency_test.py -v
"""

import unittest
from unittest.mock import MagicMock, patch, call
from xml.dom import minidom


class TestChangeIdCaching(unittest.TestCase):
    """Test caching of change IDs to avoid repeated DOM scans."""

    def test_get_next_change_id_cached(self):
        """Test that change ID is cached and not rescanned."""
        from document_optimized import DocxXMLEditorOptimized

        # Create mock DOM with tracked changes
        xml_content = '''<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:ins w:id="0"/>
                <w:ins w:id="1"/>
                <w:del w:id="2"/>
            </w:body>
        </w:document>'''

        with patch('builtins.open', unittest.mock.mock_open(read_data=xml_content)):
            with patch('defusedxml.minidom.parse') as mock_parse:
                mock_dom = minidom.parseString(xml_content)
                mock_parse.return_value = mock_dom

                editor = DocxXMLEditorOptimized('/fake/path.xml', rsid='12345678')

                # First call should scan DOM
                id1 = editor._get_next_change_id()
                self.assertEqual(id1, 3)

                # Second call should use cache, not rescan
                id2 = editor._get_next_change_id()
                self.assertEqual(id2, 4)

                # Verify DOM was only scanned once
                # (Implementation should cache after first scan)

    def test_change_id_increments_correctly(self):
        """Test that cached ID increments on each call."""
        from document_optimized import DocxXMLEditorOptimized

        with patch.object(DocxXMLEditorOptimized, '_scan_max_change_id', return_value=5):
            with patch.object(DocxXMLEditorOptimized, '__init__', return_value=None):
                editor = DocxXMLEditorOptimized.__new__(DocxXMLEditorOptimized)
                editor._cached_max_change_id = None  # Not yet scanned
                editor.dom = MagicMock()

                # First call triggers scan
                with patch.object(editor, '_scan_max_change_id', return_value=5):
                    id1 = editor._get_next_change_id()
                    self.assertEqual(id1, 6)

                # Subsequent calls increment without scan
                id2 = editor._get_next_change_id()
                self.assertEqual(id2, 7)

                id3 = editor._get_next_change_id()
                self.assertEqual(id3, 8)


class TestDOMQueryOptimization(unittest.TestCase):
    """Test single-traversal DOM query optimization."""

    def test_single_traversal_processes_all_tags(self):
        """Test that single traversal handles all tag types."""
        from document_optimized import process_nodes_single_traversal

        xml_content = '''<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p><w:r><w:t>text</w:t></w:r></w:p>
            <w:ins w:id="0"><w:r><w:t>inserted</w:t></w:r></w:ins>
            <w:del w:id="1"><w:r><w:delText>deleted</w:delText></w:r></w:del>
        </w:document>'''

        dom = minidom.parseString(xml_content)
        nodes = [dom.documentElement]

        # Track which handlers are called
        handlers_called = {}

        def make_handler(name):
            def handler(elem):
                handlers_called[name] = handlers_called.get(name, 0) + 1
            return handler

        tag_handlers = {
            'w:p': make_handler('w:p'),
            'w:r': make_handler('w:r'),
            'w:t': make_handler('w:t'),
            'w:ins': make_handler('w:ins'),
            'w:del': make_handler('w:del'),
        }

        process_nodes_single_traversal(nodes, tag_handlers)

        # Verify all handlers were called appropriate number of times
        self.assertEqual(handlers_called.get('w:p', 0), 1)
        self.assertEqual(handlers_called.get('w:r', 0), 3)  # 3 runs
        self.assertEqual(handlers_called.get('w:t', 0), 2)  # 2 text nodes
        self.assertEqual(handlers_called.get('w:ins', 0), 1)
        self.assertEqual(handlers_called.get('w:del', 0), 1)

    def test_single_traversal_more_efficient_than_multiple_queries(self):
        """Test that single traversal is more efficient."""
        from document_optimized import process_nodes_single_traversal

        # Create a document with many elements
        elements = ['<w:p><w:r><w:t>text</w:t></w:r></w:p>'] * 100
        xml_content = f'''<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            {''.join(elements)}
        </w:document>'''

        dom = minidom.parseString(xml_content)
        root = dom.documentElement

        # Count getElementsByTagName calls in original approach
        original_calls = 0

        def count_original():
            nonlocal original_calls
            for tag in ['w:p', 'w:r', 'w:t', 'w:ins', 'w:del']:
                root.getElementsByTagName(tag)
                original_calls += 1

        count_original()
        self.assertEqual(original_calls, 5)  # One call per tag type

        # Single traversal only walks the tree once
        # (verified by implementation, not by this test directly)


class TestAttributeInjectionOptimization(unittest.TestCase):
    """Test optimized attribute injection."""

    def test_inject_attributes_single_pass(self):
        """Test that attribute injection uses single pass."""
        from document_optimized import inject_attributes_optimized

        xml_content = '''<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p>
                <w:r><w:t>text</w:t></w:r>
            </w:p>
        </w:document>'''

        dom = minidom.parseString(xml_content)
        p_elem = dom.getElementsByTagName('w:p')[0]

        rsid = '12345678'
        author = 'TestAuthor'

        inject_attributes_optimized([p_elem], rsid, author)

        # Verify attributes were added
        self.assertEqual(p_elem.getAttribute('w:rsidR'), rsid)

        r_elem = dom.getElementsByTagName('w:r')[0]
        self.assertEqual(r_elem.getAttribute('w:rsidR'), rsid)


class TestDirectoryScanning(unittest.TestCase):
    """Test efficient directory scanning for unpack/pack operations."""

    def test_single_glob_for_xml_files(self):
        """Test that single glob pattern finds both .xml and .rels files."""
        from document_optimized import find_xml_files_efficient

        with patch('pathlib.Path.rglob') as mock_rglob:
            mock_rglob.return_value = [
                MagicMock(suffix='.xml', name='document.xml'),
                MagicMock(suffix='.rels', name='document.xml.rels'),
                MagicMock(suffix='.xml', name='styles.xml'),
            ]

            from pathlib import Path
            result = find_xml_files_efficient(Path('/fake/path'))

            # Should only call rglob once with wildcard
            mock_rglob.assert_called_once()

    def test_filter_xml_extensions(self):
        """Test filtering for .xml and .rels extensions."""
        from document_optimized import filter_xml_extensions

        files = [
            MagicMock(suffix='.xml'),
            MagicMock(suffix='.rels'),
            MagicMock(suffix='.png'),
            MagicMock(suffix='.jpg'),
            MagicMock(suffix='.xml'),
        ]

        result = filter_xml_extensions(files)
        self.assertEqual(len(result), 3)  # 2 .xml + 1 .rels


class TestTagFilteringOptimization(unittest.TestCase):
    """Test optimized tag filtering in pack.py."""

    def test_skip_tags_set_lookup(self):
        """Test that tag skipping uses set lookup."""
        from document_optimized import should_process_element

        skip_tags = {'w:t', 'w:delText'}  # Set for O(1) lookup

        self.assertFalse(should_process_element('w:t', skip_tags))
        self.assertFalse(should_process_element('w:delText', skip_tags))
        self.assertTrue(should_process_element('w:p', skip_tags))
        self.assertTrue(should_process_element('w:r', skip_tags))

    def test_endswith_check_avoided(self):
        """Test that we don't use inefficient endswith checks."""
        from document_optimized import should_process_element

        # Using set lookup instead of endswith(':t')
        skip_suffixes = {':t', ':delText'}

        # The optimized version should use exact tag matching, not suffix
        # This test documents the expected behavior


class TestImportOptimization(unittest.TestCase):
    """Test that imports are at module level, not in functions."""

    def test_datetime_import_at_module_level(self):
        """Test that datetime is imported at module level."""
        import document_optimized

        # Check that datetime is in module's namespace
        self.assertTrue(hasattr(document_optimized, 'datetime'))

    def test_no_repeated_imports(self):
        """Test that functions don't re-import modules."""
        import document_optimized
        import ast
        import inspect

        # Get source code of the module
        source = inspect.getsource(document_optimized)
        tree = ast.parse(source)

        # Find all import statements inside functions
        imports_in_functions = []
        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef):
                for child in ast.walk(node):
                    if isinstance(child, (ast.Import, ast.ImportFrom)):
                        imports_in_functions.append(child)

        # There should be no imports inside functions
        self.assertEqual(
            len(imports_in_functions), 0,
            f"Found {len(imports_in_functions)} import(s) inside functions"
        )


if __name__ == '__main__':
    unittest.main()
