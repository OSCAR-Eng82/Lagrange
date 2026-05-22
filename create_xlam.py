"""
create_xlam.py
Creates LagrangeRegression.xlam by:
  1. Starting LibreOffice headless on a UNO pipe
  2. Creating a blank Calc document
  3. Injecting LagrangeVBA.bas as a VBA-compatible module
  4. Saving as Excel Add-in (.xlam) via the VBA XML filter
  5. Closing LibreOffice

Run:  python3 create_xlam.py
Output: LagrangeRegression.xlam  (same directory as this script)
"""

import os
import sys
import time
import subprocess
import random

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VBA_SRC    = os.path.join(SCRIPT_DIR, "LagrangeVBA.bas")
OUTPUT     = os.path.join(SCRIPT_DIR, "LagrangeRegression.xlam")

# ---------------------------------------------------------------------------
# Add LibreOffice's Python libs to sys.path
# ---------------------------------------------------------------------------
LO_PROGRAM = "/usr/lib/libreoffice/program"
if LO_PROGRAM not in sys.path:
    sys.path.insert(0, LO_PROGRAM)

import uno
from com.sun.star.beans       import PropertyValue
from com.sun.star.connection  import NoConnectException


# ---------------------------------------------------------------------------
# Helper: build a PropertyValue
# ---------------------------------------------------------------------------
def prop(name, value):
    p = PropertyValue()
    p.Name  = name
    p.Value = value
    return p


# ---------------------------------------------------------------------------
# Start LibreOffice headless on a named pipe, return (process, pipe_name)
# ---------------------------------------------------------------------------
def start_lo():
    pipe_name = f"uno{random.randint(10**9, 10**10)}"
    cmd = [
        "soffice",
        "--headless", "--nologo", "--norestore", "--nofirststartwizard",
        f"--accept=pipe,name={pipe_name};urp;",
    ]
    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return proc, pipe_name


# ---------------------------------------------------------------------------
# Connect to a running LO instance (retry up to ~30 s)
# ---------------------------------------------------------------------------
def connect(pipe_name, retries=60, delay=0.5):
    local_ctx = uno.getComponentContext()
    resolver  = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    url = f"uno:pipe,name={pipe_name};urp;StarOffice.ComponentContext"
    for _ in range(retries):
        try:
            return resolver.resolve(url)
        except NoConnectException:
            time.sleep(delay)
    raise RuntimeError("Could not connect to LibreOffice after 30 s.")


# ---------------------------------------------------------------------------
# Main: create the XLAM
# ---------------------------------------------------------------------------
def main():
    # Read VBA source
    if not os.path.isfile(VBA_SRC):
        sys.exit(f"VBA source not found: {VBA_SRC}")
    with open(VBA_SRC, encoding="utf-8") as f:
        vba_code = f.read()

    print("Starting LibreOffice headless …")
    lo_proc, pipe_name = start_lo()

    try:
        print("Connecting via UNO …")
        ctx  = connect(pipe_name)
        smgr = ctx.ServiceManager

        desktop = smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx
        )

        # Create a blank Calc document (hidden)
        load_props = (
            prop("Hidden",       True),
            prop("MacroExecutionMode", 4),  # allow macros
        )
        doc = desktop.loadComponentFromURL(
            "private:factory/scalc", "_default", 0, load_props
        )

        print("Injecting VBA module …")
        # Access the document's Basic library container
        libs = doc.BasicLibraries
        lib_name = "LagrangeLib"
        if not libs.hasByName(lib_name):
            libs.createLibrary(lib_name)
        lib = libs.getByName(lib_name)

        mod_name = "LagrangeRegression"
        if lib.hasByName(mod_name):
            lib.replaceByName(mod_name, vba_code)
        else:
            lib.insertByName(mod_name, vba_code)

        # Mark the document as modified so LO writes the module on save
        doc.setModified(True)

        print(f"Saving as XLAM → {OUTPUT} …")
        doc.storeToURL(
            uno.systemPathToFileUrl(OUTPUT),
            (
                prop("FilterName", "Calc MS Excel 2007 VBA XML"),
                prop("Overwrite",  True),
            )
        )

        doc.close(False)
        print(f"\nSuccess!  Add-in written to:\n  {OUTPUT}")

    finally:
        lo_proc.terminate()
        lo_proc.wait(timeout=10)


if __name__ == "__main__":
    main()
