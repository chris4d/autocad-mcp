"""Report COM backend coverage against the advertised backend interface."""
import sys

sys.path.insert(0, "src")

from autocad_mcp.backends.base import AutoCADBackend
from autocad_mcp.backends.com_backend import COMBackend

declared = {
    n for n in dir(AutoCADBackend)
    if not n.startswith("_") and callable(getattr(AutoCADBackend, n, None))
}
implemented = set(dir(COMBackend))

missing = sorted(
    n for n in declared
    if n not in implemented
    or getattr(COMBackend, n, None) is getattr(AutoCADBackend, n, None)
)

print(f"declared : {len(declared)}")
print(f"missing  : {len(missing)}")
for n in missing:
    print("   ", n)
