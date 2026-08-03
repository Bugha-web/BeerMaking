"""Guard against the Streamlit-Cloud (Python 3.11) TokenError class of failure.

Streamlit hashes every @st.cache_* function with inspect.getsource(), which
slices the file from the function's line to EOF and tokenizes it. On Python
3.11 inspect.getblock() catches only EndOfBlock/IndentationError, so a
tokenize.TokenError ("EOF in multi-line string") propagates and crashes the
app. Python 3.12+ catches it, so this never reproduces locally.

This script emulates the 3.11 behaviour: for every line of the file it slices
and tokenizes, and reports any slice that dies inside an unterminated string.
Run it before pushing.

    python tools/check_py311_safe.py [file]
"""
import io
import sys
import tokenize

MAX_STRING_LINES = 12  # long literals are what make mis-sliced getblock explode


def slices_that_break(path):
    """Return line numbers whose slice-to-EOF fails to tokenize (3.11 style)."""
    lines = io.open(path, encoding="utf-8").readlines()
    bad = []
    for start in range(len(lines)):
        chunk = lines[start:]
        try:
            list(tokenize.generate_tokens(iter(chunk).__next__))
        except tokenize.TokenError as e:
            if "multi-line string" in str(e):     # the 3.11 crash signature
                bad.append((start + 1, str(e)))
        except (IndentationError, SyntaxError):
            pass  # harmless: 3.11's getblock catches IndentationError, and a
                  # mid-block slice legitimately isn't valid standalone code
    return bad


def long_string_literals(path):
    """Multi-line string literals longer than MAX_STRING_LINES."""
    out = []
    with open(path, "rb") as f:
        for tok in tokenize.tokenize(f.readline):
            if tok.type == tokenize.STRING:
                span = tok.end[0] - tok.start[0] + 1
                if span > MAX_STRING_LINES:
                    out.append((tok.start[0], tok.end[0], span))
    return out


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "streamlit_app.py"
    print(f"checking {path} (local python {sys.version.split()[0]})\n")

    longs = long_string_literals(path)
    print(f"[1] multi-line string literals over {MAX_STRING_LINES} lines")
    for a, b, n in longs:
        print(f"    lines {a}-{b} ({n} lines)  <-- getblock hazard")
    print("    none" if not longs else f"    TOTAL: {len(longs)}")

    bad = slices_that_break(path)
    print("\n[2] slices that raise TokenError('EOF in multi-line string')")
    for ln, msg in bad[:20]:
        print(f"    from line {ln}: {msg}")
    print("    none" if not bad else f"    TOTAL: {len(bad)}")

    ok = not longs and not bad
    print("\nRESULT:", "PY311-SAFE" if ok else "UNSAFE — fix before pushing")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
