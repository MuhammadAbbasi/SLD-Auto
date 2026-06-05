# -*- coding: utf-8 -*-
"""
sld_core.py - Loads the generation engine from code/smart_sld_generator.py
and re-exports the public interface needed by the web app.
"""
import importlib.util
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_SMART_PATH = os.path.join(_HERE, "code", "smart_sld_generator.py")

if not os.path.isfile(_SMART_PATH):
    raise FileNotFoundError(
        f"Core generator not found at {_SMART_PATH}. "
        "Make sure smart_sld_generator.py is present in the code/ directory."
    )

spec = importlib.util.spec_from_file_location("smart_sld_generator", _SMART_PATH)
_sld = importlib.util.module_from_spec(spec)
spec.loader.exec_module(_sld)

_orig_generate = _sld.generate
_INVERTER_POWERS = _sld._INVERTER_POWERS
_PANELS_PRESETS = _sld._PANELS_PRESETS
COL_SPACING_DEFAULT = _sld.COL_SPACING_DEFAULT
ROW_SPACING_DEFAULT = _sld.ROW_SPACING_DEFAULT


def generate(cfg, log_cb=print):
    """Wrapper that converts sys.exit() calls in the core to RuntimeError."""
    try:
        _orig_generate(cfg, log_cb=log_cb)
    except SystemExit as e:
        if e.code and e.code != 0:
            raise RuntimeError(
                "Generation failed - check the log above for the specific error."
            ) from None
