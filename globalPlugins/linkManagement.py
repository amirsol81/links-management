# -*- coding: UTF-8 -*-
"""
NVDA add-on: linkManagement

Goal:
- Add NVDA Settings panel "Links management" with one checkbox.
- When enabled AND Screen layout is OFF:
    Force each LINK onto its own "line" in browse mode by monkey-patching
    VirtualBufferTextInfo._getLineOffsets.
- Attach trailing punctuation/spaces after a link to the same line to avoid
  creating near-blank lines like "," between links.

Designed for NVDA 2026.1+ (64-bit) on Python 3.13.x.
"""

import re
import weakref

import addonHandler
import config
import controlTypes
import globalPluginHandler
import gui
from gui.settingsDialogs import SettingsPanel
import textInfos
import textInfos.offsets
import virtualBuffers
import wx

addonHandler.initTranslation()

CONF_SECTION = "linkManagement"
CONF_KEY_ENABLED = "enabled"

# Ensure our config spec exists
try:
    spec = config.conf.spec
    if CONF_SECTION not in spec:
        spec[CONF_SECTION] = {}
    if CONF_KEY_ENABLED not in spec[CONF_SECTION]:
        spec[CONF_SECTION][CONF_KEY_ENABLED] = "boolean(default=False)"
except Exception:
    # If config isn't ready for some reason, we'll behave as disabled.
    pass


def _isAddonEnabled() -> bool:
    try:
        return bool(config.conf[CONF_SECTION][CONF_KEY_ENABLED])
    except Exception:
        return False


class LinksManagementPanel(SettingsPanel):
    title = _("Links management")

    def makeSettings(self, sizer):
        helper = gui.guiHelper.BoxSizerHelper(self, sizer=sizer)
        self.chkEnable = helper.addItem(
            wx.CheckBox(
                self,
                label=_(
                    "Enable Links management (force links on separate lines when Screen layout is off)"
                ),
            )
        )
        self.chkEnable.SetValue(_isAddonEnabled())

    def onSave(self):
        config.conf[CONF_SECTION][CONF_KEY_ENABLED] = bool(self.chkEnable.GetValue())

    def onDiscard(self):
        # No special handling needed.
        pass


# --- Monkey patch: line breaking ---

_ORIG_getLineOffsets = None

# Per-virtual buffer cache so we don't re-scan fields for every arrow press.
# weak keys: VirtualBuffer object -> dict
_SEG_CACHE = weakref.WeakKeyDictionary()

# Attach trailing punctuation/spaces to the link line.
# Includes common punctuation and whitespace; also includes Unicode ellipsis (…)
# and common link separators such as slash and middle dot.
_PUNCT_RE = re.compile(r"""^[\s,.;:!?)\]\}»"\'\u2026،؛؟•|/·]+""")

# A narrow definition of visual bullet markers which may be joined to a
# following link.  Do not include numbered markers such as "1." here:
# table-of-contents widgets often expose "1." / "1.1" before an internal
# section link, and those numbers should remain separate browse-mode lines.
# The Freedom Scientific case that motivated this repair used a real bullet
# (•), so symbolic markers remain supported.
_LINK_LEADING_LIST_MARKER_RE = re.compile(r"""^\s*[•*\-–—]\s*$""")


def _isLinkLeadingListMarker(text: str) -> bool:
    return bool(text and _LINK_LEADING_LIST_MARKER_RE.match(text))

# Opening punctuation that may belong immediately before a following link, but
# can sometimes be exposed by NVDA as its own line. Keep this intentionally
# narrow: currently only an opening parenthesis is handled, for cases such as
# Wikipedia degree abbreviations:
#
#     (
#     BA)
#
# which should be browsed as:
#
#     (BA)
_LINK_LEADING_OPENER_RE = re.compile(r"""^[\s\u200b\u200c\u200d\ufeff\u2060\u200e\u200f]*\([\s\u200b\u200c\u200d\ufeff\u2060\u200e\u200f]*$""")


def _isLinkLeadingOpener(text: str) -> bool:
    return bool(text and _LINK_LEADING_OPENER_RE.match(text))


def _findTrailingLinkLeadingOpenerStart(text: str):
    """
    Return the index of a final standalone '(' that should move from the
    preceding non-link segment to the following link segment.

    This handles cases where NVDA's original line contains text before a linked
    parenthetical token, but our link segmentation would otherwise leave the
    opener as a separate non-link line:

        (
        JD)

    Only a final '(' with optional spaces/invisible chars after it is moved.
    Earlier prose remains in its own segment, so this does not merge arbitrary
    text before links.
    """
    if not text:
        return None
    i = len(text) - 1
    while i >= 0 and text[i] in _INVISIBLE_SPACE_CHARS:
        i -= 1
    if i < 0 or text[i] != "(":
        return None
    return i

# Single-character separators which may visually separate adjacent links (for
# example breadcrumbs or compact metadata lists) but should stay attached to the
# preceding link during line navigation. Keep this intentionally narrow to avoid
# merging real prose after links.
_LINK_TRAILING_SEPARATOR_RE = re.compile(r"""^\s*[/·|]\s*$""")


def _isLinkTrailingSeparator(text: str) -> bool:
    return bool(text and _LINK_TRAILING_SEPARATOR_RE.match(text))



def _getNextDistinctOriginalLine(
    ti: virtualBuffers.VirtualBufferTextInfo,
    baseStart: int,
    baseEnd: int,
    storyLen,
):
    """
    Return the next original NVDA line after [baseStart, baseEnd].

    Some virtual buffers are sensitive at exact line boundaries: asking the
    original _getLineOffsets for offset == baseEnd can return the same line or
    an empty/boundary range. Probe a few characters forward, but only accept a
    distinct non-empty line that starts at or after baseEnd. This keeps adjacent
    punctuation repairs narrow while making opener-before-link cases reliable.
    """
    if storyLen is not None and baseEnd >= storyLen:
        return None
    try:
        limit = baseEnd + 8 if storyLen is None else min(storyLen - 1, baseEnd + 8)
    except Exception:
        limit = baseEnd + 8
    seen = set()
    for probe in range(baseEnd, limit + 1):
        try:
            start, end = _ORIG_getLineOffsets(ti, probe)
        except Exception:
            continue
        rng = (start, end)
        if rng in seen:
            continue
        seen.add(rng)
        if end <= start:
            continue
        if start == baseStart and end == baseEnd:
            continue
        if start < baseEnd:
            continue
        return rng
    return None


def _getPrevDistinctOriginalLine(
    ti: virtualBuffers.VirtualBufferTextInfo, baseStart: int, baseEnd: int
):
    """Return the previous distinct non-empty original NVDA line, if nearby."""
    if baseStart <= 0:
        return None
    seen = set()
    for probe in range(baseStart - 1, max(-1, baseStart - 9), -1):
        try:
            start, end = _ORIG_getLineOffsets(ti, probe)
        except Exception:
            continue
        rng = (start, end)
        if rng in seen:
            continue
        seen.add(rng)
        if end <= start:
            continue
        if start == baseStart and end == baseEnd:
            continue
        if end > baseStart:
            continue
        return rng
    return None


# Wikipedia-style footnote/citation links may be exposed as two adjacent
# virtual-buffer lines, for example:
#
#     [۱
#     ]
#
# Keep this deliberately narrow: an opening bracket plus only digits on the
# link line, followed by a closing-bracket-only line. Supports Latin, Persian,
# and Arabic-Indic digits.
_CITATION_OPEN_RE = re.compile(r"""^\s*\[\s*[0-9۰-۹٠-٩]+\s*$""")
_CITATION_CLOSE_RE = re.compile(r"""^\s*\]\s*$""")


def _isCitationOpeningText(text: str) -> bool:
    return bool(text and _CITATION_OPEN_RE.match(text))


def _isCitationClosingText(text: str) -> bool:
    return bool(text and _CITATION_CLOSE_RE.match(text))

def _isScreenLayoutOff() -> bool:
    try:
        return not bool(config.conf["virtualBuffers"]["useScreenLayout"])
    except Exception:
        # If unknown, do not patch behavior.
        return False


def _isLinkControlField(attrs) -> bool:
    """
    attrs: ControlField dict.

    We target role=LINK (NVDA normalized role), but we deliberately IGNORE anchors
    that are exposed as buttons (e.g. <a role="button">). This prevents breaking
    layouts such as Google "Search instead for ..." where interactive tokens are
    anchor elements with button semantics.

    Keep this conservative to reduce side effects.
    """
    try:
        role = attrs.get("role")

        # If NVDA already normalized this to a button role, it is not a link for our purposes.
        if role in (
            controlTypes.Role.BUTTON,
            getattr(controlTypes.Role, "TOGGLEBUTTON", None),
            getattr(controlTypes.Role, "SPLITBUTTON", None),
            getattr(controlTypes.Role, "MENUBUTTON", None),
        ):
            return False

        # Many sites (notably Google) use <a role="button"> which may still appear
        # as a link control field in some buffers. Detect and skip via xml-roles.
        for k in (
            "xml-roles",
            "xmlRoles",
            "IAccessible2::attribute_xml-roles",
            "IAccessible2::attribute_xmlRoles",
            "aria-role",
            "ariaRole",
        ):
            v = attrs.get(k)
            if not v:
                continue
            v = str(v).lower()
            # xml-roles may be space-separated; treat any presence of 'button' as button semantics.
            if "button" in v.split() or "button" in v:
                return False

        return role == controlTypes.Role.LINK
    except Exception:
        return False


def _extendEndToIncludeTrailingPunctuation(
    ti: virtualBuffers.VirtualBufferTextInfo, end: int, limit: int
) -> int:
    """
    Extend [end, ...] to include immediate punctuation+spaces so we don't generate
    lines consisting of just "," etc between links.
    """
    if end >= limit:
        return end

    windowEnd = min(limit, end + 64)
    try:
        tail = ti._getTextRange(end, windowEnd)
    except Exception:
        return end

    if not tail:
        return end

    m = _PUNCT_RE.match(tail)
    if not m:
        return end

    return min(limit, end + m.end())


def _getLinkRangesInRange(
    ti: virtualBuffers.VirtualBufferTextInfo, rngStart: int, rngEnd: int
):
    """Return real link ranges inside [rngStart, rngEnd]."""
    try:
        rngTI = ti.obj.makeTextInfo(textInfos.offsets.Offsets(rngStart, rngEnd))
        fields = rngTI.getTextWithFields()
    except Exception:
        return []

    linkRanges = []
    for cmd in fields:
        if not (
            isinstance(cmd, textInfos.FieldCommand) and cmd.command == "controlStart"
        ):
            continue
        attrs = cmd.field
        if not _isLinkControlField(attrs):
            continue
        try:
            docHandle = int(attrs["controlIdentifier_docHandle"])
            ID = int(attrs["controlIdentifier_ID"])
            s, e = ti._getOffsetsFromFieldIdentifier(docHandle, ID)
        except Exception:
            continue

        # Keep only if inside our range window.
        if e <= rngStart or s >= rngEnd:
            continue
        s = max(s, rngStart)
        e = min(e, rngEnd)
        if e > s:
            linkRanges.append((s, e))

    return sorted(set(linkRanges))


def _computeSegmentsForParagraph(
    ti: virtualBuffers.VirtualBufferTextInfo, rngStart: int, rngEnd: int
):
    """
    Build a list of (start, end) segments for a text RANGE, splitting around links.

    Important: This must NOT change NVDA's existing line/paragraph boundaries when
    there are no links in the range. Therefore:
      - If no link ranges are detected, return None (caller should fall back).
      - Splitting is constrained to the provided range (often an existing NVDA line).

    Each link becomes its own segment; non-link text remains in segments between links.
    A symbolic bullet/list marker or a narrow opening punctuation marker immediately
    before a link may be joined to that link. Numbered markers such as "1." are
    deliberately left separate from following links.
    """
    linkRanges = _getLinkRangesInRange(ti, rngStart, rngEnd)
    if not linkRanges:
        return None

    segments = []
    cursor = rngStart
    for s, e in linkRanges:
        linkStart = s
        if cursor < s:
            try:
                preLinkText = ti._getTextRange(cursor, s)
            except Exception:
                preLinkText = ""

            if _isLinkLeadingListMarker(preLinkText) or _isLinkLeadingOpener(preLinkText):
                # The text before this link is only a symbolic bullet marker or
                # a narrow opening punctuation marker, so it belongs on the
                # same spoken/browsed line as the link. Numbered markers such
                # as "1." stay separate for table-of-contents lists.
                linkStart = cursor
            else:
                openerStartInText = _findTrailingLinkLeadingOpenerStart(preLinkText)
                if openerStartInText is not None:
                    openerStart = cursor + openerStartInText
                    if openerStart > cursor:
                        segments.append((cursor, openerStart))
                    linkStart = openerStart
                else:
                    segments.append((cursor, s))

        linkEnd = _extendEndToIncludeTrailingPunctuation(ti, e, rngEnd)
        segments.append((linkStart, linkEnd))
        cursor = linkEnd

    if cursor < rngEnd:
        segments.append((cursor, rngEnd))

    # Drop empty segments, then repair opener-before-compact-link cases such
    # as Wikipedia infobox degree abbreviations: (BA) and (JD).
    segments = [(a, b) for (a, b) in segments if b > a]
    return _mergeOpenParenBeforeCompactLinkSegments(ti, segments)



_INVISIBLE_SPACE_CHARS = " \t\r\n\u00a0\u200b\u200c\u200d\ufeff\u2060\u200e\u200f"
_PAREN_FOLLOWING_TOKEN_RE = re.compile(
    r"""^[\s\u00a0\u200b\u200c\u200d\ufeff\u2060\u200e\u200f]*([^\s\u00a0()\[\]{}]{1,32}\))"""
)


def _findCompactParenTailEndInRawText(text: str):
    """
    Return the character index just after a compact token following an opener.

    This intentionally accepts only a no-space token ending in ')' after optional
    whitespace/invisible characters, e.g. BA) or JD).  It is used only when the
    current original NVDA line is just '(', so it cannot glue arbitrary prose.
    """
    if not text:
        return None
    m = _PAREN_FOLLOWING_TOKEN_RE.match(text)
    if not m:
        return None
    token = m.group(1)
    if not _tokenLooksSafeAfterOpenParen(token):
        return None
    return m.end(1)


def _maybeExtendOpenerSegmentToCompactTail(
    ti: virtualBuffers.VirtualBufferTextInfo, start: int, end: int, storyLen
):
    """Extend an opener-only segment forward to a compact closing token.

    This covers the Wikipedia case where the add-on's own link segmentation
    leaves '(' as a small non-link segment at the end of the current baseline
    line, while the linked BA)/JD) token is immediately after it in the virtual
    buffer stream.  It is forward-only and only applies to an opener-only
    segment, so it avoids the cursor trap caused by mapping later text backward.
    """
    try:
        text = ti._getTextRange(start, end)
    except Exception:
        text = ""
    if not _isLinkLeadingOpener(text):
        return None
    if storyLen is not None and end >= storyLen:
        return None
    try:
        lookEnd = end + 40 if storyLen is None else min(storyLen, end + 40)
        followingText = ti._getTextRange(end, lookEnd)
    except Exception:
        return None
    tailEndInText = _findCompactParenTailEndInRawText(followingText)
    if tailEndInText is None:
        return None
    return start, end + tailEndInText


def _tokenLooksSafeAfterOpenParen(token: str) -> bool:
    """Return True for compact tokens such as BA), JD), or D-IL)."""
    if not token or not token.endswith(")"):
        return False
    body = token[:-1]
    if not body or len(body) > 32:
        return False
    # Avoid joining punctuation-only fragments.  Degree abbreviations and similar
    # compact Wikipedia parenthetical links contain at least one letter or digit.
    return any(ch.isalnum() for ch in body)


def _stripInvisibleText(text: str) -> str:
    return text.strip(_INVISIBLE_SPACE_CHARS)


def _findFinalOpenParenOffsetInSegment(
    ti: virtualBuffers.VirtualBufferTextInfo, start: int, end: int
):
    """
    Return the document offset of a final standalone '(' inside a segment.

    This is intentionally segment-local.  It is used after normal link
    segmentation has already happened, to catch Wikipedia infobox cases where
    the text between two links is just an opening parenthesis before a compact
    linked degree abbreviation:

        Columbia University
        (
        BA)

    If visible text exists before the final '(', that text remains a separate
    segment and only the '(' moves to the following compact/link segment.
    """
    if end <= start:
        return None
    try:
        text = ti._getTextRange(start, end)
    except Exception:
        return None
    if not text:
        return None
    i = len(text) - 1
    while i >= 0 and text[i] in _INVISIBLE_SPACE_CHARS:
        i -= 1
    if i < 0 or text[i] != "(":
        return None
    return start + i


def _segmentLooksLikeCompactParenTail(
    ti: virtualBuffers.VirtualBufferTextInfo, start: int, end: int
) -> bool:
    """Return True for compact following segments such as BA) or JD)."""
    if end <= start:
        return False
    try:
        text = ti._getTextRange(start, end)
    except Exception:
        return False
    compact = _stripInvisibleText(text)
    if not compact or any(ch.isspace() for ch in compact):
        return False
    return _tokenLooksSafeAfterOpenParen(compact)


def _mergeOpenParenBeforeCompactLinkSegments(
    ti: virtualBuffers.VirtualBufferTextInfo, segments
):
    """
    Merge a standalone '(' segment with the immediately following compact
    linked/token segment, without ever mapping a later line backward.

    This is safer than adjacent-line scanning because it operates only on the
    segments already produced for the current NVDA baseline line.  It therefore
    cannot create the Down Arrow trap that appeared when the following link line
    was mapped back to the previous opener line.
    """
    if not segments or len(segments) < 2:
        return segments

    normalized = []
    # First split any segment that ends with a final '(' so the opener itself
    # can move to the next segment while earlier visible text stays separate.
    for start, end in segments:
        parenOffset = _findFinalOpenParenOffsetInSegment(ti, start, end)
        if parenOffset is None:
            normalized.append((start, end))
            continue
        if parenOffset > start:
            normalized.append((start, parenOffset))
        normalized.append((parenOffset, end))

    merged = []
    i = 0
    while i < len(normalized):
        start, end = normalized[i]
        if i + 1 < len(normalized):
            nextStart, nextEnd = normalized[i + 1]
            try:
                currentText = ti._getTextRange(start, end)
            except Exception:
                currentText = ""
            if _isLinkLeadingOpener(currentText) and _segmentLooksLikeCompactParenTail(
                ti, nextStart, nextEnd
            ):
                merged.append((start, nextEnd))
                i += 2
                continue
        merged.append((start, end))
        i += 1

    return [(a, b) for (a, b) in merged if b > a]


def _maybeMergeAdjacentLeadingOpenerAndLink(
    ti: virtualBuffers.VirtualBufferTextInfo,
    baseStart: int,
    baseEnd: int,
    storyLen,
    offset: int,
):
    """
    Forward-only repair for an opener-only line immediately followed by a
    compact linked/token tail, as seen in Wikipedia infobox degree entries:

        (
        JD)

    Previous versions tried several adjacent-line strategies.  The safest and
    most reliable trigger is narrower: only when the current original NVDA line
    is itself just an opening parenthesis, look a few text characters forward
    and join a compact no-space token that closes the parenthesis.  We never map
    the following token line backward, so Down Arrow remains able to progress.
    """
    try:
        baseText = ti._getTextRange(baseStart, baseEnd)
    except Exception:
        baseText = ""

    if not _isLinkLeadingOpener(baseText):
        return None
    if storyLen is not None and baseEnd >= storyLen:
        return None

    # Primary path: raw-text lookahead from the opener boundary.  This handles
    # buffers where the following BA)/JD) line is not discoverable as a normal
    # next original line from this exact boundary, but the text is still adjacent
    # in the virtual buffer stream.
    try:
        lookEnd = baseEnd + 40 if storyLen is None else min(storyLen, baseEnd + 40)
        followingText = ti._getTextRange(baseEnd, lookEnd)
    except Exception:
        followingText = ""
    tailEndInText = _findCompactParenTailEndInRawText(followingText)
    if tailEndInText is not None:
        return baseStart, baseEnd + tailEndInText

    # Secondary path: ask NVDA's original line logic for the next distinct line
    # and join only a compact parenthesis-closing token or a real first link.
    nextLine = _getNextDistinctOriginalLine(ti, baseStart, baseEnd, storyLen)
    if not nextLine:
        return None
    nextStart, nextEnd = nextLine
    if nextEnd <= nextStart:
        return None

    try:
        nextText = ti._getTextRange(nextStart, nextEnd)
    except Exception:
        nextText = ""

    linkRanges = _getLinkRangesInRange(ti, nextStart, nextEnd)
    if linkRanges:
        firstStart, firstEnd = linkRanges[0]
        try:
            gap = ti._getTextRange(nextStart, firstStart)
        except Exception:
            gap = ""
        if not gap.strip(_INVISIBLE_SPACE_CHARS):
            linkEnd = _extendEndToIncludeTrailingPunctuation(ti, firstEnd, nextEnd)
            # Only accept compact tails like BA) / JD), not arbitrary link text.
            try:
                candidateText = ti._getTextRange(firstStart, linkEnd)
            except Exception:
                candidateText = ""
            if _segmentLooksLikeCompactParenTail(ti, firstStart, linkEnd) or _tokenLooksSafeAfterOpenParen(
                candidateText.strip(_INVISIBLE_SPACE_CHARS)
            ):
                return baseStart, linkEnd

    compact = nextText.strip(_INVISIBLE_SPACE_CHARS)
    if _tokenLooksSafeAfterOpenParen(compact) and not any(ch.isspace() for ch in compact):
        return baseStart, nextEnd

    return None

def _maybeMergeAdjacentListMarkerAndLink(
    ti: virtualBuffers.VirtualBufferTextInfo,
    baseStart: int,
    baseEnd: int,
    storyLen,
    offset: int,
):
    """
    Handle virtual-buffer layouts where NVDA exposes a symbolic list bullet as
    one line and the linked list-item text as the next line:

        •
        FSCompanion

    The normal link segmentation code only sees the current NVDA line. In this
    case the marker is outside the link line range, so we need one very narrow
    adjacent-line repair. This intentionally handles only symbolic marker-only
    lines next to lines that contain a real link; numbered markers such as
    "1." are excluded so TOC numbering remains on its own line.
    """
    try:
        baseText = ti._getTextRange(baseStart, baseEnd)
    except Exception:
        baseText = ""

    # Case 1: the current original line is only a marker. Join it with the first
    # real link on the immediately following original line.
    if _isLinkLeadingListMarker(baseText):
        if storyLen is not None and baseEnd >= storyLen:
            return None
        try:
            nextStart, nextEnd = _ORIG_getLineOffsets(ti, baseEnd)
        except Exception:
            return None
        if nextEnd <= nextStart or (nextStart == baseStart and nextEnd == baseEnd):
            return None
        linkRanges = _getLinkRangesInRange(ti, nextStart, nextEnd)
        if not linkRanges:
            return None
        firstStart, firstEnd = linkRanges[0]
        linkEnd = _extendEndToIncludeTrailingPunctuation(ti, firstEnd, nextEnd)
        return baseStart, linkEnd

    # Case 2: the current original line contains a link and the previous original
    # line is only a marker. Return the same combined range while the offset is
    # inside the first link segment, so Up/Down sees one stable line.
    if baseStart <= 0:
        return None
    linkRanges = _getLinkRangesInRange(ti, baseStart, baseEnd)
    if not linkRanges:
        return None
    firstStart, firstEnd = linkRanges[0]
    firstEnd = _extendEndToIncludeTrailingPunctuation(ti, firstEnd, baseEnd)
    if not (firstStart <= offset < firstEnd):
        return None
    try:
        prevStart, prevEnd = _ORIG_getLineOffsets(ti, baseStart - 1)
    except Exception:
        return None
    if prevEnd > baseStart or prevEnd <= prevStart:
        return None
    try:
        prevText = ti._getTextRange(prevStart, prevEnd)
    except Exception:
        prevText = ""
    if _isLinkLeadingListMarker(prevText):
        return prevStart, firstEnd
    return None



def _maybeMergeAdjacentLinkAndTrailingSeparator(
    ti: virtualBuffers.VirtualBufferTextInfo,
    baseStart: int,
    baseEnd: int,
    storyLen,
    offset: int,
):
    """
    Handle virtual-buffer layouts where NVDA exposes a breadcrumb separator as
    its own line immediately after a link:

        Home
        /
        Blindness and Low Vision

        Malia
        ·
        Sasha

        صفحه نخست
        |
        اخبار روز

    The regular in-line punctuation extension only sees text inside the current
    original NVDA line. If the slash is a separate original line, join it to the
    preceding link line. This is deliberately limited to single-character separator
    lines such as slash, middle dot, or vertical bar, so ordinary content after a link is not
    glued to the link.
    """
    try:
        baseText = ti._getTextRange(baseStart, baseEnd)
    except Exception:
        baseText = ""

    # Case 1: current original line contains a real link and the immediately
    # following original line is just a separator. Join the separator forward.
    linkRanges = _getLinkRangesInRange(ti, baseStart, baseEnd)
    if linkRanges and (storyLen is None or baseEnd < storyLen):
        currentLinkStart, currentLinkEnd = linkRanges[0]
        currentLinkEnd = _extendEndToIncludeTrailingPunctuation(
            ti, currentLinkEnd, baseEnd
        )
        # Only alter navigation while the caret is inside the first link segment.
        if currentLinkStart <= offset < currentLinkEnd:
            try:
                nextStart, nextEnd = _ORIG_getLineOffsets(ti, baseEnd)
            except Exception:
                nextStart = nextEnd = None
            if (
                nextStart is not None
                and nextEnd > nextStart
                and not (nextStart == baseStart and nextEnd == baseEnd)
            ):
                try:
                    nextText = ti._getTextRange(nextStart, nextEnd)
                except Exception:
                    nextText = ""
                if _isLinkTrailingSeparator(nextText):
                    return currentLinkStart, nextEnd

    # Case 2: current original line is the separator. If the previous original
    # line contains a real link, return the same combined range so Down Arrow
    # does not expose the separator as a separate stop.
    if not _isLinkTrailingSeparator(baseText) or baseStart <= 0:
        return None
    try:
        prevStart, prevEnd = _ORIG_getLineOffsets(ti, baseStart - 1)
    except Exception:
        return None
    if prevEnd > baseStart or prevEnd <= prevStart:
        return None
    prevLinkRanges = _getLinkRangesInRange(ti, prevStart, prevEnd)
    if not prevLinkRanges:
        return None
    prevLinkStart, prevLinkEnd = prevLinkRanges[0]
    prevLinkEnd = _extendEndToIncludeTrailingPunctuation(ti, prevLinkEnd, prevEnd)
    return prevLinkStart, baseEnd



def _maybeMergeAdjacentCitationBracket(
    ti: virtualBuffers.VirtualBufferTextInfo,
    baseStart: int,
    baseEnd: int,
    storyLen,
    offset: int,
):
    """
    Handle Wikipedia-style reference links split across two original NVDA
    lines, especially Persian/Arabic numeral citations:

        [۱
        ]

    This is intentionally narrower than general punctuation repair. It only
    joins a line that consists of '[' + digits to an immediately adjacent line
    that consists only of ']'. The opening line must contain a real link.
    """
    try:
        baseText = ti._getTextRange(baseStart, baseEnd)
    except Exception:
        baseText = ""

    # Case 1: current line is the opening citation link, next line is only ']'.
    if _isCitationOpeningText(baseText):
        if storyLen is not None and baseEnd >= storyLen:
            return None
        if not _getLinkRangesInRange(ti, baseStart, baseEnd):
            return None
        try:
            nextStart, nextEnd = _ORIG_getLineOffsets(ti, baseEnd)
        except Exception:
            return None
        if nextEnd <= nextStart or (nextStart == baseStart and nextEnd == baseEnd):
            return None
        try:
            nextText = ti._getTextRange(nextStart, nextEnd)
        except Exception:
            nextText = ""
        if _isCitationClosingText(nextText):
            return baseStart, nextEnd
        return None

    # Case 2: current line is the closing bracket. Return the same combined
    # range if the previous original line is a citation-opening link.
    if not _isCitationClosingText(baseText) or baseStart <= 0:
        return None
    try:
        prevStart, prevEnd = _ORIG_getLineOffsets(ti, baseStart - 1)
    except Exception:
        return None
    if prevEnd > baseStart or prevEnd <= prevStart:
        return None
    try:
        prevText = ti._getTextRange(prevStart, prevEnd)
    except Exception:
        prevText = ""
    if not _isCitationOpeningText(prevText):
        return None
    if not _getLinkRangesInRange(ti, prevStart, prevEnd):
        return None
    return prevStart, baseEnd

def _patched_getLineOffsets(self: virtualBuffers.VirtualBufferTextInfo, offset: int):
    # Only active when enabled + screen layout is OFF.
    if (not _isAddonEnabled()) or (not _isScreenLayoutOff()):
        return _ORIG_getLineOffsets(self, offset)

    # Preserve NVDA's existing line boundaries as the baseline.
    try:
        baseStart, baseEnd = _ORIG_getLineOffsets(self, offset)
    except Exception:
        return _ORIG_getLineOffsets(self, offset)

    # Use story length in cache invalidation in case buffer regenerates.
    try:
        storyLen = self._getStoryLength()
    except Exception:
        storyLen = None

    openerLinkRange = _maybeMergeAdjacentLeadingOpenerAndLink(
        self, baseStart, baseEnd, storyLen, offset
    )
    if openerLinkRange:
        return openerLinkRange

    markerLinkRange = _maybeMergeAdjacentListMarkerAndLink(
        self, baseStart, baseEnd, storyLen, offset
    )
    if markerLinkRange:
        return markerLinkRange

    separatorLinkRange = _maybeMergeAdjacentLinkAndTrailingSeparator(
        self, baseStart, baseEnd, storyLen, offset
    )
    if separatorLinkRange:
        return separatorLinkRange

    citationBracketRange = _maybeMergeAdjacentCitationBracket(
        self, baseStart, baseEnd, storyLen, offset
    )
    if citationBracketRange:
        return citationBracketRange

    buf = self.obj
    cache = _SEG_CACHE.get(buf)
    key = (baseStart, baseEnd, storyLen)

    if not cache or cache.get("key") != key:
        segments = _computeSegmentsForParagraph(self, baseStart, baseEnd)
        cache = {"key": key, "segments": segments}
        _SEG_CACHE[buf] = cache

    segments = cache.get("segments")

    # If there are no links in this baseline line, do not alter behavior.
    if not segments:
        return baseStart, baseEnd

    for s, e in segments:
        if s <= offset < e:
            openerSegmentRange = _maybeExtendOpenerSegmentToCompactTail(
                self, s, e, storyLen
            )
            if openerSegmentRange:
                return openerSegmentRange
            return s, e

    # If offset is exactly at end, snap to last segment if present.
    if segments and offset == baseEnd:
        return segments[-1]

    return baseStart, baseEnd



def _installPatch():
    global _ORIG_getLineOffsets
    if _ORIG_getLineOffsets is not None:
        return
    _ORIG_getLineOffsets = virtualBuffers.VirtualBufferTextInfo._getLineOffsets
    virtualBuffers.VirtualBufferTextInfo._getLineOffsets = _patched_getLineOffsets


def _uninstallPatch():
    global _ORIG_getLineOffsets
    if _ORIG_getLineOffsets is None:
        return
    virtualBuffers.VirtualBufferTextInfo._getLineOffsets = _ORIG_getLineOffsets
    _ORIG_getLineOffsets = None
    _SEG_CACHE.clear()


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    def __init__(self):
        super().__init__()
        _installPatch()

        # Register settings panel
        try:
            gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(
                LinksManagementPanel
            )
        except Exception:
            # If registration path changes in a future NVDA build, add-on will still
            # run; only the panel might be missing.
            pass

    def terminate(self):
        # Unregister settings panel
        try:
            gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(
                LinksManagementPanel
            )
        except Exception:
            pass

        _uninstallPatch()
        super().terminate()
