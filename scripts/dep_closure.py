"""List the exact pip distributions (with versions) needed to run DocuForge.

Imports the app's full module graph, then maps every loaded top-level package
to its installed distribution. Used to generate the deployment requirements.
"""

import importlib
import importlib.metadata as md
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

ENTRY_MODULES = [
    "src.graph.graph_builder",
    "src.auth.session",
    "src.auth.supabase_auth",
    "src.auth.guards",
    "src.app.compat",
    "src.app.components.sidebar",
    "src.app.components.message_bubble",
    "src.app.pages.login",
    "src.app.pages.chat",
]

for m in ENTRY_MODULES:
    importlib.import_module(m)

pkg2dist = md.packages_distributions()
tops = {name.split(".")[0] for name in sys.modules if "." not in name}
dists: set[str] = set()
for t in tops:
    for d in pkg2dist.get(t, []):
        dists.add(d)

for d in sorted(dists, key=str.lower):
    try:
        print(f"{d}=={md.version(d)}")
    except Exception:
        print(d)
