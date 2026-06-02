"""Style profile management for the IB Word renderer.

Swaps the global ``ib_renderer.STYLE`` singleton between named profiles without
rewriting renderer code. Every renderer reads the module-global ``STYLE`` at
call time, so rebinding it here changes their output for subsequent renders.

Profiles:
    classic  Current production styling (default; reproduces existing output).
    ib-pro   IB-grade styling. Initially identical to classic; audited
             improvements are layered on in a later phase.

Usage:
    import style_profiles
    style_profiles.set_active_profile("ib-pro")   # before rendering
    ...
    style_profiles.set_active_profile("classic")  # reset (tests do this)
"""

from __future__ import annotations

import ib_renderer
from ib_renderer import IBStyle

#: Public, ordered tuple of valid profile names (also used for CLI choices).
VALID_PROFILES = ("classic", "ib-pro")

_FACTORIES = {
    "classic": IBStyle.classic,
    "ib-pro": IBStyle.ib_pro,
}


def set_active_profile(name: str) -> IBStyle:
    """Rebind the global ``STYLE`` singleton to the named profile.

    Args:
        name: One of ``VALID_PROFILES`` ("classic" or "ib-pro").

    Returns:
        The activated :class:`IBStyle` instance.

    Raises:
        ValueError: If ``name`` is not a known profile.
    """
    if name not in _FACTORIES:
        raise ValueError(
            f"Unknown style profile: {name!r}. Choose from {list(VALID_PROFILES)}."
        )
    profile = _FACTORIES[name]()
    ib_renderer.STYLE = profile
    return profile


def get_active_profile() -> str:
    """Return the name of the currently active profile."""
    return ib_renderer.STYLE.PROFILE
