# -*- coding: utf-8 -*-

def mixed(source, target):
    """
    Returns whether the source sequence
    contains elements of the target sequence.
    """
    if not target: return False
    for item in source:
        if item in target: return True
    return False


def merge(source, target):
    """Merge source sequence to target sequence."""
    target.extend([item for item in source if item not in target])
    return target


def match(refers, target):
    """
    Returns whether the reference string sequence
    has a substring of the target string sequence.
    """
    if not target: return False
    for item in target:
        for refer in refers:
            if str(refer) in str(item): return False
    return True


def extract(cells, filter=False):
    """
    Returns the values of the specified cell sequence.
    Filter parameter to specify whether to filter empty values.
    """
    return [cell.value for cell in cells if not filter or cell.value]


def remove(target, value):
    """Delete all specified elements in sequence."""
    while value in target: target.remove(value)
    return target


def replace(target, old, new):
    """Replace all matching substrings in a string sequence."""
    for i in range(len(target)):
        if isinstance(target[i], str): target[i] = target[i].replace(old, new)
    return target


def eliminate(target):
    """Remove duplicate elements from a sequence."""
    return sorted(set(target), key=target.index)
