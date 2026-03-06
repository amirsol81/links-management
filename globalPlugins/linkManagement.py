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
# Includes common punctuation and whitespace; also includes Unicode ellipsis (…).
_PUNCT_RE = re.compile(r"""^[\s,.;:!?)\]\}»"\'\u2026،؛؟•]+""")


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
    """
    try:
        rngTI = ti.obj.makeTextInfo(textInfos.offsets.Offsets(rngStart, rngEnd))
        fields = rngTI.getTextWithFields()
    except Exception:
        return None

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

        # Keep only if inside our range window
        if e <= rngStart or s >= rngEnd:
            continue
        s = max(s, rngStart)
        e = min(e, rngEnd)
        if e > s:
            linkRanges.append((s, e))

    if not linkRanges:
        return None

    linkRanges = sorted(set(linkRanges))

    segments = []
    cursor = rngStart
    for s, e in linkRanges:
        if cursor < s:
            segments.append((cursor, s))

        linkEnd = _extendEndToIncludeTrailingPunctuation(ti, e, rngEnd)
        segments.append((s, linkEnd))
        cursor = linkEnd

    if cursor < rngEnd:
        segments.append((cursor, rngEnd))

    # Drop empty segments
    return [(a, b) for (a, b) in segments if b > a]


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
