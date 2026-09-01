"""Verified official Moody's icon names, one per role.

Every value MUST have appeared in `md.find_icons()` output -- an unresolved
name degrades silently to a bright-blue circle placeholder. Verify with
`tools/verify_icons.py` after any edit.

The official set holds two visually distinct families: the line-drawn
`Icon ...` set and the filled `..._supportive_color` set. This deck stays in
the `Icon ...` family throughout so slides do not mix styles.
"""
ICONS = {
    # slide 6 - security landscape
    "regulation": "Icon Certificate Signed Document",
    "owasp": "Icon Website Security Protection",
    # slide 8 - three use cases
    "navigator": "Icon Analytics Report",
    "uva": "Icon File Document Approved Valid",
    "mcp": "Icon Cloud Connected Connectivity",
    # slides 9/11/13 - use case fields
    "today": "Icon List Clipboard Notes",
    "assistant": "Icon Brain Thinking Idea Intelligence",
    "governed": "Icon File Document Secure Encrypted",
    "status": "Icon Flag",
    # slide 15 - conclusion principles
    "grounded": "Icon Book Learning Reading",
    "access": "Icon Keys Door",
    "human": "Icon Group People Employees",
    "audit": "Icon Screen Computer Search",
    "invite": "Icon Chat Conversation Messages",
}
