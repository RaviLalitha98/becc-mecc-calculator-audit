"""Excel function availability reference used by inventory.py and recalc.py.

Version = first *perpetual* Excel release that supports the function on
Windows desktop. Microsoft 365 subscribers received most of these earlier.
"M365" = currently subscription-only. Always confirm edge cases against
Microsoft's function reference before stating them as fact in a report.
"""

FIRST_VERSION = {
    # Excel 2010
    "AGGREGATE": "2010", "NETWORKDAYS.INTL": "2010", "WORKDAY.INTL": "2010",
    # Excel 2013
    "IFNA": "2013", "XOR": "2013", "FORMULATEXT": "2013", "ISFORMULA": "2013",
    "SHEET": "2013", "SHEETS": "2013", "DAYS": "2013", "ISOWEEKNUM": "2013",
    "NUMBERVALUE": "2013", "CEILING.MATH": "2013", "FLOOR.MATH": "2013",
    "ENCODEURL": "2013", "WEBSERVICE": "2013", "FILTERXML": "2013",
    "RRI": "2013", "PDURATION": "2013",
    # Excel 2016
    "FORECAST.ETS": "2016", "FORECAST.LINEAR": "2016",
    # Excel 2019
    "CONCAT": "2019", "TEXTJOIN": "2019", "IFS": "2019", "SWITCH": "2019",
    "MAXIFS": "2019", "MINIFS": "2019",
    # Excel 2021 (dynamic-array generation)
    "XLOOKUP": "2021", "XMATCH": "2021", "FILTER": "2021", "SORT": "2021",
    "SORTBY": "2021", "UNIQUE": "2021", "SEQUENCE": "2021", "RANDARRAY": "2021",
    "LET": "2021",
    # Excel 2024
    "LAMBDA": "2024", "MAP": "2024", "REDUCE": "2024", "SCAN": "2024",
    "MAKEARRAY": "2024", "BYROW": "2024", "BYCOL": "2024", "ISOMITTED": "2024",
    "TEXTSPLIT": "2024", "TEXTBEFORE": "2024", "TEXTAFTER": "2024",
    "VSTACK": "2024", "HSTACK": "2024", "TAKE": "2024", "DROP": "2024",
    "CHOOSECOLS": "2024", "CHOOSEROWS": "2024", "TOCOL": "2024", "TOROW": "2024",
    "WRAPCOLS": "2024", "WRAPROWS": "2024", "EXPAND": "2024",
    "ARRAYTOTEXT": "2024", "VALUETOTEXT": "2024", "IMAGE": "2024",
    # Subscription-only at time of writing
    "GROUPBY": "M365", "PIVOTBY": "M365", "PERCENTOF": "M365",
    "REGEXTEST": "M365", "REGEXEXTRACT": "M365", "REGEXREPLACE": "M365",
    "TRIMRANGE": "M365", "TRANSLATE": "M365", "DETECTLANGUAGE": "M365",
    "PY": "M365", "COPILOT": "M365",
}

# Special internal names Excel writes into the file
SPECIAL_XLFN = {
    "SINGLE": "Implicit-intersection operator '@' written by dynamic-array Excel",
    "ANCHORARRAY": "Spilled-range reference 'A1#' (dynamic arrays)",
}

VERSIONS = ["2010", "2013", "2016", "2019", "2021", "2024", "M365"]
VERSION_RANK = {v: i for i, v in enumerate(VERSIONS)}

# LibreOffice support for functions likely to matter when LibreOffice is the
# recalculation engine (recalc.py). Value = first LibreOffice release.
LIBREOFFICE_FIRST = {
    "XLOOKUP": "24.8", "XMATCH": "24.8", "FILTER": "24.8", "SORT": "24.8",
    "SORTBY": "24.8", "UNIQUE": "24.8", "SEQUENCE": "24.8", "RANDARRAY": "24.8",
    "LET": "24.8", "CHOOSECOLS": "25.2", "CHOOSEROWS": "25.2", "DROP": "25.2",
    "TAKE": "25.2", "EXPAND": "25.2", "HSTACK": "25.2", "VSTACK": "25.2",
    "TOCOL": "25.2", "TOROW": "25.2", "WRAPCOLS": "25.2", "WRAPROWS": "25.2",
    "TEXTSPLIT": "25.2", "TEXTBEFORE": "25.2", "TEXTAFTER": "25.2",
    "LAMBDA": "unsupported", "MAP": "unsupported", "REDUCE": "unsupported",
    "SCAN": "unsupported", "BYROW": "unsupported", "BYCOL": "unsupported",
    "MAKEARRAY": "unsupported", "GROUPBY": "unsupported", "PIVOTBY": "unsupported",
}

VOLATILE = {"OFFSET", "INDIRECT", "NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY",
            "CELL", "INFO"}

LOOKUP_FUNCS = {"VLOOKUP", "HLOOKUP", "LOOKUP", "MATCH", "XLOOKUP", "XMATCH", "INDEX",
                "FILTER", "CHOOSE", "OFFSET", "INDIRECT", "SUMIF", "SUMIFS", "COUNTIF",
                "COUNTIFS", "AVERAGEIF", "AVERAGEIFS", "SUMPRODUCT", "MAXIFS", "MINIFS"}

AGG_FUNCS = {"SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUBTOTAL", "AGGREGATE",
             "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN"}

ROUNDING_FUNCS = {"ROUND", "ROUNDUP", "ROUNDDOWN", "MROUND", "INT", "TRUNC", "CEILING",
                  "FLOOR", "CEILING.MATH", "FLOOR.MATH", "ODD", "EVEN"}

# Error-masking wrappers: they can silently convert a broken lookup into 0/""
MASKING_FUNCS = {"IFERROR", "IFNA", "ISERROR", "ISNA", "ISERR"}


def base_name(fn: str) -> str:
    """Strip _xlfn. / _xlws. / _xlpm. prefixes."""
    fn = fn.upper()
    for p in ("_XLFN._XLWS.", "_XLFN.", "_XLWS.", "_XLPM."):
        if fn.startswith(p):
            return fn[len(p):]
    return fn


def supported_in(fn: str, version: str):
    first = FIRST_VERSION.get(base_name(fn))
    if first is None:
        return None  # long-standing or unknown
    return VERSION_RANK[version] >= VERSION_RANK[first]
