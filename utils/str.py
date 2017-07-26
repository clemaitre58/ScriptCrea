import difflib


def similarity(seq1, seq2):
    seq = difflib.SequenceMatcher(a=seq1.lower(), b=seq2.lower())
    return seq.ratio()
