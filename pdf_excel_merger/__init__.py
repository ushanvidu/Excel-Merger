"""PDF -> Excel Merger.

A small toolkit that extracts tabular data from digital (text-based) PDFs and
merges it into the columns of an existing Excel template, under user-controlled
column mapping.

Modules
-------
- pdf_extractor : convert one or more PDFs into a single clean DataFrame
- excel_reader  : read a target Excel template and locate its header row/columns
- mapping       : auto-suggest source->target column mappings (fuzzy matching)
- merger        : write mapped source data into the template, preserving formatting
- utils         : shared helpers (header normalization, value coercion)
"""

__version__ = "0.1.0"
