"""Shared pytest fixtures for the IB report formatter test suite."""

import pytest


@pytest.fixture(autouse=True)
def _reset_style_profile():
    """Keep every test on the 'classic' style profile.

    The renderer reads a global ``ib_renderer.STYLE`` singleton that
    ``style_profiles.set_active_profile`` rebinds. Resetting to 'classic' before
    and after each test prevents a test that activates 'ib-pro' from leaking
    global state into unrelated tests.
    """
    import style_profiles

    style_profiles.set_active_profile("classic")
    yield
    style_profiles.set_active_profile("classic")
