"""
FIGURE CAPTION + ALT TEXT EXTRACTOR  (accessibility-tagged PDFs)
================================================================

For every <Figure> tag in the PDF structure tree:

  1. Alt Text   -> "Alternate Text for Images" of the <Figure> tag
                   (Acrobat: Object Properties > Tag; stored as /Alt).
                   If the <Figure> itself has none, the alt text is looked
                   up in its attribute objects and then in its child tags.
  2. Caption    -> the tag immediately following the <Figure>.
                   It is a caption only if it is a <P> or <Caption> tag and
                   its text starts with  Figure 1 / Figure 10 / Fig. 1 /
                   Fig. 10 / Figure 2.1 ...
                   If no such SIBLING is found, a <P>/<Caption> NESTED
                   INSIDE the <Figure> itself (e.g.
                   <Figure><Caption><Span>Figure 1 ...) is also checked
                   as a fallback.
  3. Figure No. -> that leading label ("Figure 1", "Fig. 10", "Figure 2.1").
                   (FIGURE_NO_MODE = "sequence" gives Figure 1, Figure 2, ...)
                   Small-caps labels that are stored as "figure 4.1" in the
                   PDF are written as "Figure 4.1".
                   If a <Figure> has no numbered caption (e.g. an unnumbered
                   screenshot followed only by "Source: ..."), the cell gets
                   the marker "[NO FIGURE NO.]" (orange).

Excel output (ONLY these six columns):

    Date | Chapter | Figure No. | Figure Page Number | Figure Caption | Figure Alt Text

    Figure Page Number = page where the <Figure> is (see PAGE_NUMBER_SOURCE).

Install:
    pip install pikepdf openpyxl pdfminer.six

Run:
    python figure_caption_alt_extractor.py
    python figure_caption_alt_extractor.py "D:\\some\\other\\folder"
"""

import os
import re
import sys
import copy
import math
import logging

from datetime import datetime
from collections import defaultdict

import pikepdf

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# pdfminer.six is used only to read the TEXT of the caption
try:

    from pdfminer.converter import PDFPageAggregator
    from pdfminer.layout import LTChar, LTContainer
    from pdfminer.pdfdocument import PDFDocument
    from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
    from pdfminer.pdfpage import PDFPage
    from pdfminer.pdfparser import PDFParser
    from pdfminer.pdftypes import resolve1
    from pdfminer.psparser import PSLiteral

    PDFMINER_AVAILABLE = True

    logging.getLogger("pdfminer").setLevel(logging.ERROR)

except ImportError:

    PDFMINER_AVAILABLE = False


# ============================================================
# CONFIGURATION
# ============================================================

FOLDER_PATH = sys.argv[1] if len(sys.argv) > 1 else r"D:\python\testing"

EXCEL_PATH = os.path.join(
    FOLDER_PATH,
    "Accessibility_Report.xlsx"
)

# Output goes to the SECOND sheet of the workbook (0 = first, 1 = second).
# Other sheets are never touched. If the workbook has only one sheet,
# a new sheet named OUTPUT_SHEET_NAME is added as the second sheet.
# If the workbook does not exist, a new one is created with that single sheet.
OUTPUT_SHEET_INDEX = 1
OUTPUT_SHEET_NAME = "Figure Caption"

# Tag(s) allowed to hold a caption - either as the tag directly following
# <Figure>, or nested inside <Figure> itself. Some PDFs use a plain <P>,
# others use the dedicated PDF structure type <Caption>.
CAPTION_TAGS = {"P", "Caption"}

# Wrapper tags skipped while looking for "the tag after <Figure>"
CONTAINER_TAGS = {"Part", "Art", "Sect", "Div", "NonStruct", "Private"}

# Caption must START with a label like:
#   Figure 1 | Figure 10 | Figure 2.1 | Figure A.1 | Fig. 1 | Fig. 10 | Fig 3a
# group(1) = the label  ->  goes to the "Figure No." column
CAPTION_LABEL_REGEX = re.compile(
    r"^\s*("
    r"(?i:f\s?i\s?g\s?u\s?r\s?e|f\s?i\s?g)\.?\s*"
    r"(?:[A-Z]{1,2}[.\-\u2013\u2014]?)?"
    r"\d+(?:[.\-\u2013\u2014]\d+)*"
    r"(?:[a-z](?![A-Za-z]))?"
    r")"
)

# "caption"  : Figure No. = label found in the caption ("Figure 2.1", "Fig. 10")
# "sequence" : Figure No. = running number in the PDF ("Figure 1", "Figure 2" ...)
FIGURE_NO_MODE = "caption"

# characters that may sit in front of the caption label (soft hyphen,
# zero-width, anchored-object / private-use glyphs ...)
LEADING_JUNK_REGEX = re.compile(r"^[\s\u00ad\u200b\u200c\u200d\u2060\ufeff\ufffc\ue000-\uf8ff]+")

# Written into Excel when something is missing
MISSING_ALT_MARKER = "[MISSING ALT TEXT]"
MISSING_CAPTION_MARKER = "[CAPTION NOT FOUND]"
MISSING_FIGURE_NO_MARKER = "[NO FIGURE NO.]"

# False : figure without a numbered caption -> "[NO FIGURE NO.]"  (recommended,
#         the running number of the PDF does not match the book numbering)
# True  : write the running number in the PDF instead ("Figure 5", "Figure 6"...)
USE_RUNNING_NUMBER_WHEN_NO_CAPTION = False

RED_FILL = PatternFill(fill_type="solid", fgColor="FFC7CE")       # missing
ORANGE_FILL = PatternFill(fill_type="solid", fgColor="FFEB9C")    # Figure No. = running number

# ---- "Figure Page Number" column ------------------------------
# "label" : page number printed in the book / shown first in Acrobat's page
#           box (e.g. 126). Falls back to the PDF page number when the PDF
#           has no page labels.
# "index" : PDF page number, 1 = first page of the PDF (e.g. 28)
PAGE_NUMBER_SOURCE = "label"

MISSING_PAGE_MARKER = "[PAGE NOT FOUND]"

# ---- the six output columns ----------------------------------
FIELD_ORDER = ["date", "chapter", "figure_no", "page", "caption", "alt"]

CANONICAL_HEADERS = {
    "date": "Date",
    "chapter": "Chapter",
    "figure_no": "Figure No.",
    "page": "Figure Page Number",
    "caption": "Figure Caption",
    "alt": "Figure Alt Text",
}

# existing headers are matched case-insensitively
HEADER_ALIASES = {
    "date": {"date"},
    "chapter": {"chapter"},
    "figure_no": {"figure no.", "figure no", "figure number", "fig no.", "fig no"},
    "page": {"figure page number", "figure page no.", "figure page no",
             "figure page", "page number", "page no.", "page no", "page"},
    "caption": {"figure caption", "caption"},
    "alt": {"figure alt text", "alt text", "alt"},
}


# ============================================================
# BUILD PAGE MAP
# ============================================================

def build_page_map(pdf):
    """
    Map PDF page object -> page number (starts from 1).
    """

    page_map = {}

    for page_number, page in enumerate(pdf.pages, start=1):

        try:
            page_map[page.obj.objgen] = page_number
        except Exception:
            pass

    return page_map


# ============================================================
# PAGE LABELS  (printed page numbers, e.g. 126)
# ============================================================

def build_page_labels(pdf):
    """
    {pdf page number (1-based): page label}.
    Uses the /PageLabels of the PDF; without them label == PDF page number.
    """

    labels = {}

    for page_number, page in enumerate(pdf.pages, start=1):

        try:
            labels[page_number] = str(page.label)

        except Exception:
            labels[page_number] = str(page_number)

    return labels


def figure_no_value(figure):
    """
    Text for the "Figure No." cell.
    """

    if FIGURE_NO_MODE == "sequence":
        return figure["running_no"]

    if figure["figure_no"]:
        return figure["figure_no"]

    if USE_RUNNING_NUMBER_WHEN_NO_CAPTION:
        return figure["running_no"]

    return MISSING_FIGURE_NO_MARKER


def page_cell_value(figure):
    """
    Value for the "Figure Page Number" cell.
    Numbers are written as numbers, roman numerals / "A-3" as text.
    """

    if not figure["page"]:
        return MISSING_PAGE_MARKER

    if PAGE_NUMBER_SOURCE == "index":
        text = str(figure["page"])
    else:
        text = figure["page_label"] or str(figure["page"])

    return int(text) if text.isdigit() else text


# ============================================================
# GET STRUCTURE ELEMENT PAGE NUMBER
# ============================================================

def get_struct_page_number(struct_elem, page_map, inherited_page=None):
    """
    Page number from /Pg. If /Pg is missing, use parent/inherited page.
    """

    current_page = inherited_page

    try:

        if "/Pg" in struct_elem:

            pg = struct_elem["/Pg"]

            try:

                objgen = pg.objgen

                if objgen in page_map:
                    current_page = page_map[objgen]

            except Exception:
                pass

    except Exception:
        pass

    return current_page


# ============================================================
# ROLE MAP  (custom tag -> standard tag)
# ============================================================

def get_role_map(struct_root):

    role_map = {}

    try:

        if "/RoleMap" in struct_root:

            raw = struct_root["/RoleMap"]

            for key in raw.keys():
                role_map[str(key).lstrip("/")] = str(raw[key]).lstrip("/")

    except Exception:
        pass

    return role_map


def resolve_tag(raw_tag, role_map):

    tag = str(raw_tag).lstrip("/").strip()

    seen = set()

    while tag in role_map and tag not in seen:
        seen.add(tag)
        tag = role_map[tag]

    return tag


# ============================================================
# FLATTEN STRUCTURE TREE INTO DOCUMENT (READING) ORDER
# ============================================================

def flatten_structure(node, page_map, role_map, inherited_page, depth, out):
    """
    Pre-order walk (same order as the Tags panel in Acrobat).
    Every structure element -> {"elem", "tag", "page", "depth"}
    """

    if isinstance(node, pikepdf.Array):

        for child in node:

            flatten_structure(
                child, page_map, role_map,
                inherited_page, depth, out
            )

        return

    if not isinstance(node, pikepdf.Dictionary):
        return

    # MCR / OBJR dictionaries have no /S -> not structure elements
    if "/S" not in node:
        return

    current_page = get_struct_page_number(node, page_map, inherited_page)

    out.append({
        "elem": node,
        "tag": resolve_tag(node["/S"], role_map),
        "page": current_page,
        "depth": depth
    })

    try:

        if "/K" in node:

            flatten_structure(
                node["/K"], page_map, role_map,
                current_page, depth + 1, out
            )

    except Exception:
        pass


# ============================================================
# TEXT OF MARKED CONTENT  (MCID -> text)  USING pdfminer.six
# ============================================================
#
# Every content item in the Tags panel (e.g. "Figure 2.1" and
# "Existing theoretical explanations") is a marked-content sequence
# with an MCID in the page content stream. We read the page with
# pdfminer and keep a proper STACK of marked-content ids, so text is
# found even when the MCID sequence contains nested tags such as
# /Span, /Lang, /ActualText, /Artifact ... (very common in
# InDesign / Acrobat exported PDFs).

if PDFMINER_AVAILABLE:

    class MarkedContentAggregator(PDFPageAggregator):
        """
        Collects characters and remembers the MCID each one belongs to.
        Paths and images are ignored (not needed, and they are what makes
        pdfminer slow / fail on figure-heavy pages).
        """

        def __init__(self, rsrcmgr):

            super().__init__(rsrcmgr, pageno=1, laparams=None)

            self.mcid_stack = []
            self.properties = {}

        def begin_tag(self, tag, props=None):

            mcid = None

            try:

                # /P /MC0 BDC  -> property list stored in page /Properties
                if isinstance(props, PSLiteral):
                    props = resolve1(self.properties.get(props.name))

                props = resolve1(props)

                if isinstance(props, dict) and "MCID" in props:
                    mcid = int(resolve1(props["MCID"]))

            except Exception:
                mcid = None

            self.mcid_stack.append(mcid)

        def end_tag(self):

            if self.mcid_stack:
                self.mcid_stack.pop()

        def current_mcid(self):

            # innermost marked-content that carries an MCID
            for mcid in reversed(self.mcid_stack):

                if mcid is not None:
                    return mcid

            return None

        def render_char(self, *args, **kwargs):

            advance = super().render_char(*args, **kwargs)

            try:

                item = self.cur_item._objs[-1]

                if isinstance(item, LTChar):
                    item.mcid = self.current_mcid()

            except Exception:
                pass

            return advance

        def render_string(self, *args, **kwargs):

            # one broken font / string must not lose the rest of the page
            try:
                super().render_string(*args, **kwargs)
            except Exception:
                pass

        def paint_path(self, *args, **kwargs):
            return

        def render_image(self, *args, **kwargs):
            return


    def iter_chars(container, in_form=False):
        """
        Yields (char, in_form). in_form is True for text drawn inside a
        Form XObject (placed artwork), whose MCIDs are numbered separately
        from the page content.
        """

        for obj in container:

            if isinstance(obj, LTChar):
                yield obj, in_form

            elif isinstance(obj, LTContainer):
                yield from iter_chars(obj, True)


def chars_to_text(chars):
    """
    Join pdfminer characters (content-stream order) into a string.
    A space is added on a visible gap or a line change.
    """

    text = ""
    prev = None

    for ch in chars:

        value = ch.get_text()

        if prev is not None and value.strip() and not text.endswith(" "):

            size = getattr(ch, "size", None) or 10

            same_line = abs(ch.y0 - prev.y0) < size * 0.5

            gap = ch.x0 - prev.x1

            if (not same_line) or gap > size * 0.2:
                text += " "

        text += value

        prev = ch

    return text


class PageTextCache:
    """
    Lazily builds the text of every marked-content id (MCID) of a page.

    MCIDs of page content and of Form XObjects (placed artwork with its own
    tagged text such as "= 24" or "a") can be the same numbers, so both are
    kept apart; the page content wins unless the structure says /Stm.
    """

    def __init__(self, pdf_path):

        self.pages = []
        self.cache = {}
        self.errors = []
        self.init_error = None
        self._fp = None

        if not PDFMINER_AVAILABLE:

            self.init_error = (
                "pdfminer.six is not installed - caption text cannot be "
                "read (pip install pdfminer.six)."
            )

            return

        try:

            self._fp = open(pdf_path, "rb")

            document = PDFDocument(PDFParser(self._fp))

            self.pages = list(PDFPage.create_pages(document))

        except Exception as e:

            self.init_error = f"pdfminer could not open the PDF: {e}"

    def close(self):

        if self._fp is not None:

            try:
                self._fp.close()
            except Exception:
                pass

    def get(self, page_number, mcid, form=False):

        if page_number not in self.cache:
            self.cache[page_number] = self._build(page_number)

        page_text, form_text = self.cache[page_number]

        if form:
            return form_text.get(mcid) or page_text.get(mcid) or ""

        return page_text.get(mcid) or form_text.get(mcid) or ""

    def _build(self, page_number):

        page_text = {}
        form_text = {}

        if self.init_error or not page_number:
            return page_text, form_text

        if page_number > len(self.pages):
            return page_text, form_text

        device = None

        try:

            page = self.pages[page_number - 1]

            resource_manager = PDFResourceManager()

            device = MarkedContentAggregator(resource_manager)

            try:
                device.properties = resolve1(
                    page.resources.get("Properties")
                ) or {}
            except Exception:
                device.properties = {}

            interpreter = PDFPageInterpreter(resource_manager, device)

            try:

                interpreter.process_page(page)

                layout = device.get_result()

            except Exception as e:

                self.errors.append(f"page {page_number}: {e}")

                # keep everything that was read before the problem
                layout = device.cur_item

            page_groups = defaultdict(list)
            form_groups = defaultdict(list)

            for ch, in_form in iter_chars(layout):

                mcid = getattr(ch, "mcid", None)

                if mcid is None:
                    continue

                (form_groups if in_form else page_groups)[mcid].append(ch)

            for mcid, group in page_groups.items():
                page_text[mcid] = chars_to_text(group)

            for mcid, group in form_groups.items():
                form_text[mcid] = chars_to_text(group)

        except Exception as e:

            self.errors.append(f"page {page_number}: {e}")

        return page_text, form_text


def _collect_text(node, page_map, text_cache, inherited_page, parts):

    if isinstance(node, bool):
        return

    # ---- MCID given directly in /K -------------------------
    if isinstance(node, int):

        if inherited_page is not None:

            text = text_cache.get(inherited_page, node)

            if text:
                parts.append(text)

        return

    if isinstance(node, pikepdf.Array):

        for child in node:
            _collect_text(child, page_map, text_cache, inherited_page, parts)

        return

    if not isinstance(node, pikepdf.Dictionary):
        return

    # ---- Marked-content reference (MCR) --------------------
    if "/MCID" in node:

        page = get_struct_page_number(node, page_map, inherited_page)

        try:
            mcid = int(node["/MCID"])
        except Exception:
            return

        if page is not None:

            text = text_cache.get(page, mcid, form="/Stm" in node)

            if text:
                parts.append(text)

        return

    # ---- /ActualText on the tag itself wins ----------------
    try:

        if "/ActualText" in node:

            actual = str(node["/ActualText"]).strip()

            if actual:
                parts.append(actual)
                return

    except Exception:
        pass

    page = get_struct_page_number(node, page_map, inherited_page)

    if "/K" in node:
        _collect_text(node["/K"], page_map, text_cache, page, parts)


LIGATURES = {
    "\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl",
    "\ufb03": "ffi", "\ufb04": "ffl", "\ufb05": "st", "\ufb06": "st",
}


def clean_text(text):

    # remove control characters (e.g. \x00 at the end of an alt text - Excel
    # refuses them) and zero-width characters, then collapse all whitespace
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    text = re.sub(r"[\u00ad\u200b\u200c\u200d\u2060\ufeff]", "", text)

    # ligature glyphs from the PDF font -> normal letters (ﬁ -> fi, ﬂ -> fl ...)
    for ligature, letters in LIGATURES.items():
        text = text.replace(ligature, letters)

    return re.sub(r"\s+", " ", text).strip()


def get_element_text(elem, page_map, text_cache, inherited_page):

    parts = []

    _collect_text(elem, page_map, text_cache, inherited_page, parts)

    return clean_text(" ".join(parts))


# ============================================================
# ALT TEXT  ("Alternate Text for Images" in Acrobat)
# ============================================================

def alt_from_dict(dictionary):
    """
    /Alt of one dictionary (structure element or attribute object).
    """

    try:

        if "/Alt" in dictionary:
            return clean_text(str(dictionary["/Alt"]))

    except Exception:
        pass

    return ""


def iter_attribute_dicts(elem):
    """
    Attribute objects (/A) can be one dictionary or an array of them.
    """

    try:

        if "/A" not in elem:
            return

        attrs = elem["/A"]

        items = list(attrs) if isinstance(attrs, pikepdf.Array) else [attrs]

        for item in items:

            if isinstance(item, pikepdf.Dictionary):
                yield item

    except Exception:
        return


def alt_from_descendants(node, budget):
    """
    Depth-first search for the first non-empty /Alt below a tag.
    budget = [remaining number of elements to look at] (safety limit).
    """

    if isinstance(node, pikepdf.Array):

        for child in node:

            alt = alt_from_descendants(child, budget)

            if alt:
                return alt

        return ""

    if not isinstance(node, pikepdf.Dictionary) or "/S" not in node:
        return ""

    if budget[0] <= 0:
        return ""

    budget[0] -= 1

    alt = alt_from_dict(node)

    if alt:
        return alt

    try:

        if "/K" in node:
            return alt_from_descendants(node["/K"], budget)

    except Exception:
        pass

    return ""


def read_alt_text(figure_elem):
    """
    Returns (alt_text, source)

    1. /Alt of the <Figure> tag itself      (Object Properties > Tag)
    2. /Alt inside its attribute objects
    3. /Alt of a child tag (e.g. nested <Figure>)
    """

    alt = alt_from_dict(figure_elem)

    if alt:
        return alt, "Figure tag"

    for attribute in iter_attribute_dicts(figure_elem):

        alt = alt_from_dict(attribute)

        if alt:
            return alt, "attribute object"

    try:

        if "/K" in figure_elem:

            alt = alt_from_descendants(figure_elem["/K"], [500])

            if alt:
                return alt, "child tag"

    except Exception:
        pass

    return "", ""


def describe_element(elem, role_map):
    """
    Short description of a tag, used in the warning for a missing alt text
    (shows which keys the tag has and what its children are).
    """

    try:
        keys = ", ".join(sorted(str(k) for k in elem.keys()))
    except Exception:
        keys = "?"

    kids = []

    try:

        if "/K" in elem:

            raw = elem["/K"]

            items = list(raw) if isinstance(raw, pikepdf.Array) else [raw]

            for child in items:

                if isinstance(child, pikepdf.Dictionary) and "/S" in child:
                    kids.append(resolve_tag(child["/S"], role_map))

                else:
                    kids.append("content")

    except Exception:
        pass

    shown = ", ".join(kids[:8]) + (" ..." if len(kids) > 8 else "")

    return f"tag keys: {keys}; children: {shown or 'none'}"


# ============================================================
# CAPTION / FIGURE NO. DETECTION
# - A caption may be a <P> or <Caption>.
# - A nested caption may be <Figure><P><Span>Figure 1 ...</Span></P></Figure>.
# - For Figure No., we also inspect <Span> directly as a fallback.
# ============================================================

def normalise_label(label):
    """
    Small-caps labels are stored in the PDF as lower case ("figure 4.1"),
    letter-spaced ones as "F igure 4.1".
    -> "Figure 4.1" / "Fig. 4.1"
    """

    match = re.match(r"(?i)(f\s?i\s?g\s?u\s?r\s?e|f\s?i\s?g)(.*)$", label)

    if not match:
        return label

    letters = re.sub(r"\s", "", match.group(1)).lower()

    word = "Figure" if letters == "figure" else "Fig"

    return word + match.group(2)


def describe_following(nodes, figure_index, page_map, text_cache, limit=3):
    """
    The next few tags after a <Figure> with their text - shown in the warning
    when no caption is found, so it is easy to see what the tags really are.
    """

    depth = nodes[figure_index]["depth"]

    j = figure_index + 1

    while j < len(nodes) and nodes[j]["depth"] > depth:
        j += 1

    shown = []

    while j < len(nodes) and len(shown) < limit:

        node = nodes[j]

        j += 1

        if node["tag"] in CONTAINER_TAGS:
            continue

        text = ""

        if node["tag"] in ("P", "Caption", "H1", "H2", "H3", "H4", "Lbl"):

            text = get_element_text(
                node["elem"], page_map, text_cache, node["page"]
            )

        text = text[:40] + ("..." if len(text) > 40 else "")

        shown.append(f"<{node['tag']}> \"{text}\"" if text else f"<{node['tag']}>")

    return ", ".join(shown) if shown else "nothing"


def _match_caption_text(text):
    """
    Given raw element text, strips leading junk and checks it against the
    caption label pattern. Returns (caption, figure_no) or (None, None).
    """

    text = LEADING_JUNK_REGEX.sub("", text)

    if not text:
        return None, None

    match = CAPTION_LABEL_REGEX.match(text)

    if not match:
        return None, None

    figure_no = normalise_label(re.sub(r"\s+", " ", match.group(1)).strip())

    caption = figure_no + text[match.end(1):]

    return caption, figure_no


def find_caption_sibling(nodes, figure_index, page_map, text_cache):
    """
    Caption strategy 1:
    Find the first <P> or <Caption> tag immediately FOLLOWING <Figure>.

    The text can contain nested <Span> tags; get_element_text() reads the
    complete text under the candidate tag.

    Returns (caption_text, figure_no, reason_if_not_found)
    """

    figure_depth = nodes[figure_index]["depth"]

    # Skip the Figure's own children.
    j = figure_index + 1

    while j < len(nodes) and nodes[j]["depth"] > figure_depth:
        j += 1

    # Skip wrapper tags (Sect / Div / ...).
    while j < len(nodes) and nodes[j]["tag"] in CONTAINER_TAGS:
        j += 1

    if j >= len(nodes):
        return None, None, "no tag follows the <Figure>"

    candidate = nodes[j]

    if candidate["tag"] not in CAPTION_TAGS:
        return None, None, (
            f"next tag after <Figure> is <{candidate['tag']}>, "
            f"not one of {sorted(CAPTION_TAGS)}"
        )

    page_hint = candidate["page"] or nodes[figure_index]["page"]

    text = get_element_text(
        candidate["elem"],
        page_map,
        text_cache,
        page_hint
    )

    raw_text = LEADING_JUNK_REGEX.sub("", text)

    if not raw_text:
        return None, None, (
            f"the <{candidate['tag']}> after <Figure> has no readable text "
            f"(PDF page {page_hint})"
        )

    caption, figure_no = _match_caption_text(text)

    if not caption:
        snippet = raw_text if len(raw_text) <= 60 else raw_text[:57] + "..."

        return None, None, (
            f"the <{candidate['tag']}> after <Figure> does not start with "
            f"'Figure N' / 'Fig. N' (\"{snippet}\")"
        )

    return caption, figure_no, None


def find_caption_in_figure(nodes, figure_index, page_map, text_cache):
    """
    Caption strategy 2:
    Some tagged PDFs put the caption INSIDE <Figure>, for example:

        <Figure>
            <P>
                <Span>Figure 1 ... caption text ...</Span>
            </P>
        </Figure>

    Both <P> and <Caption> are checked.  If the text on the P/Caption
    wrapper is not exposed correctly by a particular PDF, its nested
    <Span> is also checked as a direct text source.

    Returns (caption_text, figure_no, reason_if_not_found)
    """

    figure_depth = nodes[figure_index]["depth"]
    j = figure_index + 1

    checked_any_caption_tag = False

    while j < len(nodes) and nodes[j]["depth"] > figure_depth:

        node = nodes[j]

        if node["tag"] in CAPTION_TAGS:
            checked_any_caption_tag = True

            page_hint = node["page"] or nodes[figure_index]["page"]

            text = get_element_text(
                node["elem"], page_map, text_cache, page_hint
            )

            caption, figure_no = _match_caption_text(text)

            if caption:
                return caption, figure_no, None

        j += 1

    if not checked_any_caption_tag:
        return None, None, (
            "no <P> or <Caption> tag found nested inside the <Figure>"
        )

    return None, None, (
        "found <P>/<Caption> tag(s) nested inside the <Figure>, but none "
        "of their text starts with 'Figure N' / 'Fig. N'"
    )


def find_figure_no_in_figure(nodes, figure_index, page_map, text_cache):
    """
    Figure-number-only fallback.

    This is specifically for PDFs whose structure is like:

        <Figure>
            <P>
                <Span>Figure 1 ...</Span>
            </P>
        </Figure>

    If the <P>/<Caption> wrapper does not expose the text as expected,
    inspect the nested <Span> directly.  Only the leading Figure/Fig label
    is returned; the full caption is NOT required.

    Returns (figure_no, reason_if_not_found)
    """

    figure_depth = nodes[figure_index]["depth"]

    # Prefer P / Caption first, then Span.
    preferred_tags = {"P", "Caption", "Span"}

    j = figure_index + 1

    while j < len(nodes) and nodes[j]["depth"] > figure_depth:

        node = nodes[j]

        if node["tag"] in preferred_tags:
            page_hint = node["page"] or nodes[figure_index]["page"]

            text = get_element_text(
                node["elem"], page_map, text_cache, page_hint
            )

            if text:
                _, figure_no = _match_caption_text(text)

                if figure_no:
                    return figure_no, None

        j += 1

    return None, (
        "no Figure/Fig. number found in the nested <P>/<Caption>/<Span>"
    )


def find_caption(nodes, figure_index, page_map, text_cache):
    """
    Tries:
      1. sibling <P>/<Caption>
      2. nested <P>/<Caption>
      3. nested <Span> as a figure-number-only fallback

    Returns (caption_text, figure_no, reason_if_not_found).
    """

    caption, figure_no, reason_sibling = find_caption_sibling(
        nodes, figure_index, page_map, text_cache
    )

    if caption:
        return caption, figure_no, None

    caption, figure_no, reason_nested = find_caption_in_figure(
        nodes, figure_index, page_map, text_cache
    )

    if caption:
        return caption, figure_no, None

    # We could not recover the full caption, but the user still needs
    # Figure No. in the Excel C column.
    figure_no, reason_number = find_figure_no_in_figure(
        nodes, figure_index, page_map, text_cache
    )

    if figure_no:
        return None, figure_no, (
            f"{reason_sibling}; also checked inside <Figure>: "
            f"{reason_nested}; Figure No. recovered from nested tag."
        )

    return None, None, (
        f"{reason_sibling}; also checked inside <Figure>: "
        f"{reason_nested}; {reason_number}"
    )


# ============================================================
# EXTRACT ALL FIGURES FROM ONE PDF
# ============================================================

def extract_figures(pdf_path):
    """
    Returns (figures, warnings)

    figures = [
        {"page": 2, "page_label": "19", "figure_no": "Figure 2.1",
         "running_no": "Figure 1",
         "caption": "Figure 2.1 Existing ...", "alt": "..."},
        ...
    ]   (in PDF order)
    """

    figures = []
    warnings = []

    text_cache = PageTextCache(pdf_path)

    try:

        if text_cache.init_error:
            warnings.append(text_cache.init_error)

        with pikepdf.open(pdf_path) as pdf:

            if "/StructTreeRoot" not in pdf.Root:

                warnings.append(
                    "PDF is not tagged (no StructTreeRoot) - "
                    "no figures can be extracted."
                )

                return figures, warnings

            struct_root = pdf.Root["/StructTreeRoot"]

            page_map = build_page_map(pdf)

            page_labels = build_page_labels(pdf)

            role_map = get_role_map(struct_root)

            nodes = []

            if "/K" in struct_root:

                flatten_structure(
                    struct_root["/K"], page_map, role_map,
                    None, 0, nodes
                )

            nested_until = -1     # index where the current <Figure> subtree ends

            for index, item in enumerate(nodes):

                if item["tag"] != "Figure":
                    continue

                # a <Figure> inside another <Figure> belongs to the outer one
                if index < nested_until:
                    continue

                nested_until = index + 1

                while (
                    nested_until < len(nodes)
                    and nodes[nested_until]["depth"] > item["depth"]
                ):
                    nested_until += 1

                sequence = len(figures) + 1

                page = item["page"]

                where = (
                    f"Figure #{sequence} in PDF (PDF page {page})"
                    if page else f"Figure #{sequence} in PDF"
                )

                # ---- ALT TEXT ------------------------------
                alt, alt_source = read_alt_text(item["elem"])

                if not alt:

                    warnings.append(
                        f"{where}: Alt Text is missing/empty "
                        f"({describe_element(item['elem'], role_map)})."
                    )

                elif alt_source != "Figure tag":

                    warnings.append(
                        f"{where}: Alt Text was not on the <Figure> tag "
                        f"itself - taken from its {alt_source}."
                    )

                # ---- CAPTION + FIGURE NO. ------------------
                caption, figure_no, reason = find_caption(
                    nodes, index, page_map, text_cache
                )

                if not caption:

                    warnings.append(
                        f"{where}: caption not found - {reason}. "
                        f"Tags after the figure: "
                        f"{describe_following(nodes, index, page_map, text_cache)}"
                    )

                figures.append({
                    "page": page,
                    "page_label": page_labels.get(page, "") if page else "",
                    "figure_no": figure_no or "",
                    "running_no": f"Figure {sequence}",
                    "caption": caption or "",
                    "alt": alt
                })

            for message in text_cache.errors:
                warnings.append(f"Could not read page text - {message}")

            if not figures:
                warnings.append("No <Figure> tags found in this PDF.")

    finally:

        text_cache.close()

    return figures, warnings


# ============================================================
# EXCEL : CREATE / OPEN / SAVE
# ============================================================

def style_header_cells(cells):

    yellow_fill = PatternFill(fill_type="solid", fgColor="FFFF00")

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )

    for cell in cells:

        cell.fill = yellow_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border


def write_default_headers(worksheet):

    for col, field in enumerate(FIELD_ORDER, start=1):

        worksheet.cell(row=1, column=col).value = CANONICAL_HEADERS[field]

    style_header_cells(
        [worksheet.cell(row=1, column=c) for c in range(1, len(FIELD_ORDER) + 1)]
    )

    for column, width in {
        "A": 15, "B": 45, "C": 14, "D": 20, "E": 60, "F": 80
    }.items():
        worksheet.column_dimensions[column].width = width


def create_excel():
    """
    New workbook containing ONLY the output sheet.
    """

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = OUTPUT_SHEET_NAME

    write_default_headers(worksheet)

    workbook.save(EXCEL_PATH)


def open_workbook():
    """
    Returns (workbook, created_new)
    """

    created_new = False

    if not os.path.exists(EXCEL_PATH):

        print("\nExcel not found - creating a new workbook.")

        create_excel()

        created_new = True

    try:

        return load_workbook(EXCEL_PATH), created_new

    except PermissionError:

        raise RuntimeError(
            "Cannot open the Excel file. Close it in Excel/LibreOffice "
            f"and run again: {EXCEL_PATH}"
        )


def looks_like_output_sheet(worksheet):

    fields = set()

    for cell in worksheet[1]:

        if cell.value is None:
            continue

        key = str(cell.value).strip().lower()

        for field, aliases in HEADER_ALIASES.items():

            if key in aliases:
                fields.add(field)

    return {"chapter", "caption", "alt"} <= fields


def get_output_sheet(workbook, created_new):

    if created_new:
        return workbook.worksheets[0]

    if len(workbook.worksheets) > OUTPUT_SHEET_INDEX:
        return workbook.worksheets[OUTPUT_SHEET_INDEX]

    # a workbook created by this script in an earlier run has ONLY the output
    # sheet - keep using it instead of adding another sheet every run
    only_sheet = workbook.worksheets[0]

    if looks_like_output_sheet(only_sheet):
        return only_sheet

    print(
        f"\nWorkbook has only {len(workbook.worksheets)} sheet(s) - "
        f"adding sheet '{OUTPUT_SHEET_NAME}'."
    )

    sheet = workbook.create_sheet(OUTPUT_SHEET_NAME)

    write_default_headers(sheet)

    return sheet


def save_workbook(workbook):

    try:

        workbook.save(EXCEL_PATH)

    except PermissionError:

        raise PermissionError(
            "Cannot save the Excel file - it is open in another program. "
            f"Close it and run again: {EXCEL_PATH}"
        )


# ============================================================
# EXCEL : HEADER ROW  ->  exactly the six columns
# ============================================================

def get_header_cells(worksheet):

    return [
        c for c in worksheet[1]
        if c.value is not None and str(c.value).strip() != ""
    ]


def map_columns(worksheet):
    """
    {"date": col, "chapter": col, ...} from the existing header row.
    """

    columns = {}

    for cell in get_header_cells(worksheet):

        key = str(cell.value).strip().lower()

        for field, aliases in HEADER_ALIASES.items():

            if key in aliases and field not in columns:
                columns[field] = cell.column

    return columns


def insert_column(worksheet, position):
    """
    Insert an empty column and shift column widths with the data.
    """

    max_col = worksheet.max_column

    widths = {}

    for col in range(position, max_col + 1):

        letter = get_column_letter(col)

        if letter in worksheet.column_dimensions:
            widths[col] = worksheet.column_dimensions[letter].width

    worksheet.insert_cols(position)

    for col, width in widths.items():

        if width:
            worksheet.column_dimensions[
                get_column_letter(col + 1)
            ].width = width


def prepare_output_sheet(worksheet):
    """
    Makes sure the sheet has Date | Chapter | Figure No. | Figure Page Number |
    Figure Caption | Figure Alt Text. Existing headers/formatting are kept;
    only missing ones are added ("Figure No." is inserted right after
    "Chapter", "Figure Page Number" right after "Figure No.").
    """

    if not get_header_cells(worksheet):
        write_default_headers(worksheet)

    columns = map_columns(worksheet)

    # ---- Figure No. : insert after Chapter -----------------
    if "figure_no" not in columns:

        headers = get_header_cells(worksheet)

        if "chapter" in columns:

            position = columns["chapter"] + 1

            insert_column(worksheet, position)

            style_source = worksheet.cell(row=1, column=columns["chapter"])

        else:

            position = headers[-1].column + 1

            style_source = headers[-1]

        new_cell = worksheet.cell(row=1, column=position)

        new_cell.value = CANONICAL_HEADERS["figure_no"]

        new_cell._style = copy.copy(style_source._style)

        worksheet.column_dimensions[get_column_letter(position)].width = 14

        columns = map_columns(worksheet)

    # ---- Figure Page Number : insert after Figure No. ------
    if "page" not in columns:

        headers = get_header_cells(worksheet)

        position = columns["figure_no"] + 1

        insert_column(worksheet, position)

        style_source = worksheet.cell(row=1, column=columns["figure_no"])

        new_cell = worksheet.cell(row=1, column=position)

        new_cell.value = CANONICAL_HEADERS["page"]

        new_cell._style = copy.copy(style_source._style)

        worksheet.column_dimensions[get_column_letter(position)].width = 20

        columns = map_columns(worksheet)

    # ---- any other missing header : add at the end ---------
    for field in ("date", "chapter", "caption", "alt"):

        if field in columns:
            continue

        headers = get_header_cells(worksheet)

        position = headers[-1].column + 1

        new_cell = worksheet.cell(row=1, column=position)

        new_cell.value = CANONICAL_HEADERS[field]

        new_cell._style = copy.copy(headers[-1]._style)

        columns = map_columns(worksheet)

    # ---- warn about extra columns --------------------------
    known = set(columns.values())

    extra = [
        str(c.value) for c in get_header_cells(worksheet)
        if c.column not in known
    ]

    if extra:

        print(
            "  NOTE: output sheet has extra column(s) "
            f"{', '.join(extra)} - ignored (output should have only the "
            "six required columns)."
        )

    return columns


# ============================================================
# EXCEL : ROW HELPERS
# ============================================================

def find_chapter_rows(worksheet, chapter_col, chapter_name):

    rows = []

    for row in range(2, worksheet.max_row + 1):

        value = worksheet.cell(row=row, column=chapter_col).value

        if value is None:
            continue

        if str(value).strip().lower() == chapter_name.strip().lower():
            rows.append(row)

    return rows


def last_used_row(worksheet):

    for row in range(worksheet.max_row, 1, -1):

        if any(c.value not in (None, "") for c in worksheet[row]):
            return row

    return 1


def excel_safe(value):
    """
    Excel/openpyxl refuse control characters (e.g. \x00) in a cell.
    """

    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)

    return value


def get_column_width(worksheet, column):

    letter = get_column_letter(column)

    try:

        if letter in worksheet.column_dimensions:

            width = worksheet.column_dimensions[letter].width

            if width:
                return width

    except Exception:
        pass

    return 13


def estimate_row_height(worksheet, row, columns):
    """
    Long alt text / captions are wrapped, so the row is made tall enough to
    show all of it (otherwise Excel / LibreOffice can show a clipped or
    empty-looking cell).
    """

    lines_needed = 1

    for field in ("caption", "alt"):

        value = worksheet.cell(row=row, column=columns[field]).value

        if not value:
            continue

        chars_per_line = max(
            8,
            int(get_column_width(worksheet, columns[field]) * 1.05)
        )

        lines = sum(
            max(1, math.ceil(len(part) / chars_per_line))
            for part in str(value).split("\n")
        )

        lines_needed = max(lines_needed, lines)

    return min(409, max(15, lines_needed * 15))


# ============================================================
# EXCEL : WRITE FIGURES
# ============================================================

def write_figures_to_excel(worksheet, chapter, figures):
    """
    Re-running the same chapter UPDATES its rows (no duplicates).
    Rows are written in PDF order.
    """

    columns = prepare_output_sheet(worksheet)

    current_date = datetime.now().strftime("%d-%m-%Y")

    existing_rows = find_chapter_rows(worksheet, columns["chapter"], chapter)

    center = Alignment(horizontal="center", vertical="top")

    left_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    next_row = last_used_row(worksheet) + 1

    for position, figure in enumerate(figures):

        if position < len(existing_rows):
            row = existing_rows[position]            # reuse old row
        else:
            row = next_row                           # new row at bottom
            next_row += 1

        values = {
            "date": current_date,
            "chapter": chapter,
            "figure_no": figure_no_value(figure),
            "page": page_cell_value(figure),
            "caption": figure["caption"] or MISSING_CAPTION_MARKER,
            "alt": figure["alt"] or MISSING_ALT_MARKER
        }

        for field, value in values.items():

            cell = worksheet.cell(row=row, column=columns[field])

            cell.value = excel_safe(value)
            cell.border = thin_border
            cell.alignment = (
                left_wrap if field in ("caption", "alt") else center
            )
            cell.fill = PatternFill(fill_type=None)

        # highlight items that need checking
        if FIGURE_NO_MODE != "sequence" and not figure["figure_no"]:
            worksheet.cell(
                row=row, column=columns["figure_no"]
            ).fill = ORANGE_FILL

        if not figure["caption"]:
            worksheet.cell(row=row, column=columns["caption"]).fill = RED_FILL

        if not figure["alt"]:
            worksheet.cell(row=row, column=columns["alt"]).fill = RED_FILL

        worksheet.row_dimensions[row].height = estimate_row_height(
            worksheet, row, columns
        )

    # ---- PDF now has fewer figures than before -> drop extras
    for row in sorted(existing_rows[len(figures):], reverse=True):
        worksheet.delete_rows(row)


# ============================================================
# GET ALL PDF FILES
# ============================================================

def get_pdf_files():

    if not os.path.exists(FOLDER_PATH):
        raise RuntimeError(f"Folder does not exist: {FOLDER_PATH}")

    pdf_files = [
        f for f in os.listdir(FOLDER_PATH)
        if f.lower().endswith(".pdf")
    ]

    pdf_files.sort()

    return pdf_files


# ============================================================
# MAIN AUTOMATION
# ============================================================

def main():

    print("=" * 70)
    print("FIGURE CAPTION + ALT TEXT EXTRACTION")
    print("=" * 70)

    print(f"\nFolder: {FOLDER_PATH}")
    print(f"Excel : {EXCEL_PATH}")

    # ---- FIND PDFs -----------------------------------------
    try:
        pdf_files = get_pdf_files()

    except RuntimeError as e:
        print(f"\nERROR: {e}")
        return

    if not pdf_files:
        print("\nNo PDF files found.")
        return

    print(f"\nPDF files found: {len(pdf_files)}")

    # ---- OPEN WORKBOOK -------------------------------------
    try:

        workbook, created_new = open_workbook()

        output_sheet = get_output_sheet(workbook, created_new)

    except RuntimeError as e:

        print(f"\nERROR: {e}")
        return

    print(f"Output sheet: '{output_sheet.title}'")

    summary = []       # one line per PDF for the final summary
    errors = []        # PDFs that could not be processed

    # ---- PROCESS EACH PDF ----------------------------------
    for index, pdf_file in enumerate(pdf_files, start=1):

        print("\n" + "=" * 70)
        print(f"[{index}/{len(pdf_files)}] Processing: {pdf_file}")

        pdf_path = os.path.join(FOLDER_PATH, pdf_file)

        chapter = os.path.splitext(pdf_file)[0]

        try:

            print(f"Chapter : {chapter}")

            figures, warnings = extract_figures(pdf_path)

            print(f"Figures found: {len(figures)}")

            for figure in figures:

                figure_label = figure_no_value(figure)

                alt_preview = figure["alt"][:45] + (
                    "..." if len(figure["alt"]) > 45 else ""
                )

                print(
                    f"  {figure_label:<15} "
                    f"page {page_cell_value(figure)}  "
                    f"caption={'YES' if figure['caption'] else 'NO '}  "
                    f"alt={'YES' if figure['alt'] else 'NO '}"
                    + (f"  \"{alt_preview}\"" if figure["alt"] else "")
                )

            for message in warnings:
                print(f"  WARNING: {message}")

            write_figures_to_excel(output_sheet, chapter, figures)

            save_workbook(workbook)

            print("Excel updated successfully.")

            summary.append((
                pdf_file,
                len(figures),
                sum(1 for f in figures if not f["figure_no"]),
                sum(1 for f in figures if not f["caption"]),
                sum(1 for f in figures if not f["alt"]),
                [w for w in warnings if "not tagged" in w or "not installed" in w
                 or "could not" in w.lower()]
            ))

        except PermissionError as e:

            # Excel file is locked - no point continuing
            print(f"\nERROR: {e}")

            return

        except Exception as e:

            if isinstance(e, pikepdf.PasswordError):
                msg = "PDF is password protected - skipped."
            else:
                msg = f"{type(e).__name__}: {e}"

            print(f"ERROR processing {pdf_file}: {msg}")

            errors.append((pdf_file, msg))

            # throw away half-written rows of this PDF, start clean again
            try:

                workbook, created_new = open_workbook()

                output_sheet = get_output_sheet(workbook, created_new)

            except RuntimeError as reload_error:

                print(f"\nERROR: {reload_error}")

                return

    # ---- FINAL SUMMARY -------------------------------------
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)

    for name, total, no_number, no_caption, no_alt, notes in summary:

        print(
            f"- {name}: {total} figure(s) | "
            f"no figure no.: {no_number} | "
            f"no caption: {no_caption} | "
            f"no alt text: {no_alt}"
        )

        for note in notes:
            print(f"    NOTE: {note}")

    for name, msg in errors:
        print(f"- {name}: NOT WRITTEN TO EXCEL - {msg}")

    print("\n" + "=" * 70)
    print("COMPLETED")
    print("=" * 70)


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()