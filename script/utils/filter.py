# -*- coding: utf-8 -*-

from operator import eq, lt, le, gt, ge, ne

methods = {
    '<': lt, '<=': le, '>': gt, '>=': ge, '=': eq, '!=': ne,
}


def filter(cells, rules):
    """
    Filter cells on matching rules.
    Returns matching results for a cells.
    """
    for field, cell in cells:
        if field not in rules: continue

        for symbol, expect in rules[field]:
            if not methods[symbol](*retype(cell.value, expect)):
                return False

    return True


def retype(source, target):
    """
    Reasonably convert the type of the value
    used for comparison.
    """
    try:return int(source), int(target)
    except:pass
    try:return float(source), float(target)
    except:pass
    return str(source), str(target)
