#!/usr/bin/env python3
"""
Optimized utilities for docx document processing.
Provides efficient DOM traversal, caching, and attribute injection.
"""

from datetime import datetime, timezone  # Module-level import (not inside functions)
from pathlib import Path
from typing import Dict, List, Callable, Optional, Set, Any, Iterator
from xml.dom.minidom import Element, Document


def process_nodes_single_traversal(
    nodes: List[Element],
    tag_handlers: Dict[str, Callable[[Element], None]]
) -> None:
    """
    Process DOM nodes in a single traversal, dispatching to tag-specific handlers.

    This is more efficient than calling getElementsByTagName() multiple times.

    Args:
        nodes: List of root nodes to traverse.
        tag_handlers: Dict mapping tag names to handler functions.
    """
    def traverse(node: Element) -> None:
        if node.nodeType != node.ELEMENT_NODE:
            return

        # Check if we have a handler for this tag
        handler = tag_handlers.get(node.tagName)
        if handler:
            handler(node)

        # Recursively process children
        for child in node.childNodes:
            traverse(child)

    for node in nodes:
        traverse(node)


def inject_attributes_optimized(
    nodes: List[Element],
    rsid: str,
    author: str,
    initials: str = "C"
) -> None:
    """
    Inject RSID, author, and date attributes using single traversal.

    Args:
        nodes: List of nodes to process.
        rsid: RSID value to inject.
        author: Author name for tracked changes.
        initials: Author initials for comments.
    """
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    def add_rsid_to_p(elem: Element) -> None:
        if not elem.hasAttribute("w:rsidR"):
            elem.setAttribute("w:rsidR", rsid)
        if not elem.hasAttribute("w:rsidRDefault"):
            elem.setAttribute("w:rsidRDefault", rsid)
        if not elem.hasAttribute("w:rsidP"):
            elem.setAttribute("w:rsidP", rsid)

    def add_rsid_to_r(elem: Element) -> None:
        # Check if inside deletion
        parent = elem.parentNode
        inside_del = False
        while parent:
            if hasattr(parent, 'tagName') and parent.tagName == "w:del":
                inside_del = True
                break
            parent = getattr(parent, 'parentNode', None)

        if inside_del:
            if not elem.hasAttribute("w:rsidDel"):
                elem.setAttribute("w:rsidDel", rsid)
        else:
            if not elem.hasAttribute("w:rsidR"):
                elem.setAttribute("w:rsidR", rsid)

    def add_xml_space_to_t(elem: Element) -> None:
        if elem.firstChild and elem.firstChild.nodeType == elem.firstChild.TEXT_NODE:
            text = elem.firstChild.data
            if text and (text[0].isspace() or text[-1].isspace()):
                if not elem.hasAttribute("xml:space"):
                    elem.setAttribute("xml:space", "preserve")

    def add_tracked_change_attrs(elem: Element) -> None:
        if not elem.hasAttribute("w:author"):
            elem.setAttribute("w:author", author)
        if not elem.hasAttribute("w:date"):
            elem.setAttribute("w:date", timestamp)

    def add_comment_attrs(elem: Element) -> None:
        if not elem.hasAttribute("w:author"):
            elem.setAttribute("w:author", author)
        if not elem.hasAttribute("w:date"):
            elem.setAttribute("w:date", timestamp)
        if not elem.hasAttribute("w:initials"):
            elem.setAttribute("w:initials", initials)

    # Map tags to handlers
    tag_handlers = {
        'w:p': add_rsid_to_p,
        'w:r': add_rsid_to_r,
        'w:t': add_xml_space_to_t,
        'w:ins': add_tracked_change_attrs,
        'w:del': add_tracked_change_attrs,
        'w:comment': add_comment_attrs,
    }

    process_nodes_single_traversal(nodes, tag_handlers)


def find_xml_files_efficient(directory: Path) -> List[Path]:
    """
    Find all XML and RELS files in a directory efficiently.

    Uses a single glob with filtering instead of multiple globs.

    Args:
        directory: Directory to search.

    Returns:
        List of Path objects for XML and RELS files.
    """
    all_files = directory.rglob("*")
    return filter_xml_extensions(list(all_files))


def filter_xml_extensions(files: List[Any]) -> List[Any]:
    """
    Filter a list of files to only XML and RELS files.

    Args:
        files: List of file objects with .suffix attribute.

    Returns:
        Filtered list containing only .xml and .rels files.
    """
    valid_extensions = {'.xml', '.rels'}
    return [f for f in files if getattr(f, 'suffix', '').lower() in valid_extensions]


def should_process_element(tag_name: str, skip_tags: Set[str]) -> bool:
    """
    Check if an element should be processed based on its tag.

    Uses set lookup for O(1) performance.

    Args:
        tag_name: The element's tag name.
        skip_tags: Set of tag names to skip.

    Returns:
        True if the element should be processed.
    """
    return tag_name not in skip_tags


class DocxXMLEditorOptimized:
    """
    Optimized XML editor with cached change ID tracking.

    Key optimizations:
    - Caches max change ID instead of rescanning DOM
    - Uses single-traversal attribute injection
    - Module-level imports
    """

    def __init__(self, xml_path: str, rsid: str, author: str = "Claude", initials: str = "C"):
        """
        Initialize the optimized editor.

        Args:
            xml_path: Path to XML file.
            rsid: RSID for tracked changes.
            author: Author name.
            initials: Author initials.
        """
        self.xml_path = xml_path
        self.rsid = rsid
        self.author = author
        self.initials = initials
        self.dom = None
        self._cached_max_change_id: Optional[int] = None

        # Load DOM
        self._load_dom()

    def _load_dom(self) -> None:
        """Load the DOM from the XML file."""
        from defusedxml import minidom
        with open(self.xml_path, 'r', encoding='utf-8') as f:
            self.dom = minidom.parse(f)

    def _scan_max_change_id(self) -> int:
        """
        Scan the DOM for the maximum change ID.

        This is only called once; subsequent calls use cached value.

        Returns:
            Maximum change ID found, or -1 if none.
        """
        max_id = -1

        for tag in ("w:ins", "w:del"):
            elements = self.dom.getElementsByTagName(tag)
            for elem in elements:
                change_id = elem.getAttribute("w:id")
                if change_id:
                    try:
                        max_id = max(max_id, int(change_id))
                    except ValueError:
                        pass

        return max_id

    def _get_next_change_id(self) -> int:
        """
        Get the next available change ID.

        Uses cached value and increments, avoiding repeated DOM scans.

        Returns:
            Next available change ID.
        """
        if self._cached_max_change_id is None:
            self._cached_max_change_id = self._scan_max_change_id()

        self._cached_max_change_id += 1
        return self._cached_max_change_id

    def inject_attributes(self, nodes: List[Element]) -> None:
        """
        Inject attributes to nodes using optimized single traversal.

        Args:
            nodes: List of nodes to process.
        """
        inject_attributes_optimized(nodes, self.rsid, self.author, self.initials)


class ChangeIdCache:
    """
    Standalone cache for tracking change IDs across a document.

    Use this when you need to track IDs without the full editor.
    """

    def __init__(self, initial_max_id: int = -1):
        """
        Initialize the cache.

        Args:
            initial_max_id: Starting maximum ID (scan DOM first if needed).
        """
        self._max_id = initial_max_id

    @classmethod
    def from_dom(cls, dom: Document) -> 'ChangeIdCache':
        """
        Create a cache by scanning a DOM document.

        Args:
            dom: The DOM document to scan.

        Returns:
            ChangeIdCache initialized with max ID from document.
        """
        max_id = -1

        for tag in ("w:ins", "w:del"):
            elements = dom.getElementsByTagName(tag)
            for elem in elements:
                change_id = elem.getAttribute("w:id")
                if change_id:
                    try:
                        max_id = max(max_id, int(change_id))
                    except ValueError:
                        pass

        return cls(initial_max_id=max_id)

    def get_next_id(self) -> int:
        """
        Get the next available change ID.

        Returns:
            Next change ID (auto-increments internal counter).
        """
        self._max_id += 1
        return self._max_id

    @property
    def current_max(self) -> int:
        """Get the current maximum ID without incrementing."""
        return self._max_id
