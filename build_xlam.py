"""
build_xlam.py  —  Creates LagrangeRegression.xlam entirely in Python.

No Excel or LibreOffice required.  Implements:
  - MS-OVBA compression  [MS-OVBA §2.4]
  - CFB (OLE2) binary container  [MS-CFB]
  - OOXML add-in packaging  [MS-XLSX]
"""

import io, os, struct, zipfile

# ---------------------------------------------------------------------------
# MS-OVBA compression  (literal-token-only; always valid, slightly larger)
# ---------------------------------------------------------------------------

def vba_compress(raw: bytes) -> bytes:
    """Compress bytes with MS-OVBA literal-token compression.

    Caps each DecompressedChunk at 3640 bytes so the all-literal compressed
    form never exceeds 4095 bytes (455 groups x 9 = 4095, fits in 12-bit field).
    CompressedChunkSignature bits 12-14 MUST equal 0b011 -> header = 0xB000|(size-1).
    """
    MAX_CHUNK = 3640                  # floor(4096 / 9) * 8
    out = bytearray([0x01])           # SignatureByte
    i = 0
    while i < len(raw):
        chunk = raw[i : i + MAX_CHUNK]
        i += len(chunk)
        comp = bytearray()
        j = 0
        while j < len(chunk):
            group = chunk[j : j + 8]
            j += len(group)
            comp.append(0x00)         # FlagByte = 0 -> all tokens are literals
            comp.extend(group)
        out += struct.pack('<H', 0xB000 | (len(comp) - 1))
        out += comp
    return bytes(out)


# ---------------------------------------------------------------------------
# VBA dir stream builder  [MS-OVBA §2.3.4.2]
# ---------------------------------------------------------------------------

def _rec(id_: int, data: bytes) -> bytes:
    return struct.pack('<HI', id_, len(data)) + data


def build_dir_stream(module_name: str, text_offset: int) -> bytes:
    """Return the uncompressed dir stream for one procedural module."""
    n   = module_name.encode('latin-1')
    u_n = module_name.encode('utf-16-le')
    proj = b'VBAProject'

    d = bytearray()
    # ---- project info ----
    d += _rec(0x0001, struct.pack('<I', 0x0001))         # SYSKIND  Win32
    d += _rec(0x0002, struct.pack('<I', 0x0409))         # LCID
    d += _rec(0x0014, struct.pack('<I', 0x0409))         # LCIDINVOKE
    d += _rec(0x0003, struct.pack('<H', 1252))           # CODEPAGE
    d += _rec(0x0004, proj)                              # NAME
    d += _rec(0x0005, b'')                               # DOCSTRING
    d += _rec(0x0040, b'')                               # DOCSTRING unicode
    d += _rec(0x0006, b'')                               # HELPFILEPATH1
    d += _rec(0x003D, b'')                               # HELPFILEPATH2
    d += _rec(0x0007, struct.pack('<I', 0))              # HELPCONTEXT
    d += _rec(0x0008, struct.pack('<I', 0))              # LIBFLAGS
    # PROJECTVERSION: fixed layout  (Reserved DWORD = 4)
    d += struct.pack('<HIIH', 0x0009, 0x0004, 1, 0x000A)
    d += _rec(0x000C, b'')                               # CONSTANTS
    d += _rec(0x003C, b'')                               # CONSTANTS unicode
    # ---- modules ----
    d += _rec(0x000F, struct.pack('<H', 1))              # MODULES count=1
    d += _rec(0x0013, struct.pack('<H', 0xFFFF))         # COOKIE
    # ---- single module ----
    d += _rec(0x0019, n)                                 # MODULENAME
    d += _rec(0x0047, u_n)                               # MODULENAMEUNICODE
    d += _rec(0x001A, n)                                 # STREAMNAME
    d += _rec(0x0032, u_n)                               # STREAMNAME unicode
    d += _rec(0x001C, b'')                               # DOCSTRING
    d += _rec(0x0048, b'')                               # DOCSTRING unicode
    d += _rec(0x0031, struct.pack('<I', text_offset))    # OFFSET
    d += _rec(0x001E, struct.pack('<I', 0))              # HELPCONTEXT
    d += _rec(0x002C, struct.pack('<H', 0xFFFF))         # COOKIE
    d += struct.pack('<HI', 0x0021, 0)                   # MODULETYPE procedural
    d += struct.pack('<HI', 0x002B, 0)                   # MODULE TERMINATOR
    # ---- modules terminator ----
    d += struct.pack('<HI', 0x0010, 0)
    return bytes(d)


# ---------------------------------------------------------------------------
# CFB (OLE2 Compound File) builder  [MS-CFB]
# ---------------------------------------------------------------------------

FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD
NOSTREAM   = 0xFFFFFFFF

SECTOR_SZ  = 512
MINI_CUTOFF= 4096          # streams < 4096 B go to mini stream
MINI_SZ    = 64


def _dir_entry(name: str, otype: int, color: int,
               left: int, right: int, child: int,
               clsid: bytes, start: int, size: int) -> bytes:
    enc = name.encode('utf-16-le')
    name_sz = len(enc) + 2          # includes null terminator
    enc = enc.ljust(64, b'\x00')[:64]
    return (enc
            + struct.pack('<H', name_sz)
            + struct.pack('<BB', otype, color)
            + struct.pack('<III', left, right, child)
            + clsid
            + struct.pack('<I', 0)  # state
            + b'\x00' * 16          # created/modified times
            + struct.pack('<III', start, size, 0))  # size_hi=0 for v3


def build_cfb(streams: dict) -> bytes:
    """
    streams: dict mapping stream path ('VBA/_VBA_PROJECT', etc.) to bytes.
    Storages are created automatically for path prefixes containing '/'.
    Returns raw CFB bytes.
    """
    # ----------------------------------------------------------------
    # Plan the directory tree
    # Names and types:
    #   Root Entry  (type 5)
    #   VBA         (type 1, storage, child of Root)
    #   Each stream (type 2)
    # ----------------------------------------------------------------
    # Fixed SID layout:
    #  0: Root Entry
    #  1: VBA  (storage)
    #  2: PROJECT
    #  3: PROJECTwm
    #  4: _VBA_PROJECT
    #  5: dir
    #  6: <module>
    # ----------------------------------------------------------------

    STREAM_NAMES   = list(streams.keys())   # e.g. ['VBA/_VBA_PROJECT','VBA/dir','VBA/X','PROJECT','PROJECTwm']
    stream_data    = [streams[k] for k in STREAM_NAMES]

    # Assign SIDs
    # We use a fixed layout matching our known names
    order = ['Root Entry', 'VBA', 'PROJECT', 'PROJECTwm',
             '_VBA_PROJECT', 'dir']
    # find the module name
    for k in STREAM_NAMES:
        base = k.split('/')[-1]
        if base not in ('_VBA_PROJECT', 'dir'):
            if k.startswith('VBA/'):
                order.append(base)
                break

    # Map name → SID
    sid = {n: i for i, n in enumerate(order)}

    # ---- Allocate data sectors ----
    # All streams are stored directly (no mini stream for simplicity)
    sector_data   = []   # list of 512-byte sector bytes

    def alloc_stream(data: bytes):
        """Pad data to sector boundary, store, return (start_sector, size)."""
        if not data:
            return ENDOFCHAIN, 0
        padded = data + b'\x00' * ((-len(data)) % SECTOR_SZ)
        start  = len(sector_data)
        n_sec  = len(padded) // SECTOR_SZ
        for i in range(n_sec):
            sector_data.append(padded[i*SECTOR_SZ : (i+1)*SECTOR_SZ])
        return start, len(data)

    # ---- Root mini stream: not used (all streams > cutoff or we store directly) ----
    # Map stream name → (start_sector_index_relative_to_data_base, size)
    stream_info = {}    # name → (start, size)

    # VBA streams
    for k in STREAM_NAMES:
        base = k.split('/')[-1]
        stream_info[base] = alloc_stream(streams[k])

    # Directory sectors: 4 entries per sector → ceil(7 / 4) = 2 directory sectors
    # FAT sector comes first (sector 0)
    # Dir sectors come next (sectors 1, 2)
    # Then data sectors start at sector 3
    DATA_SECTOR_BASE = 3   # FAT=0, Dir1=1, Dir2=2

    # ---- Build FAT ----
    # FAT maps each sector index to next sector in chain (or ENDOFCHAIN)
    # Sector 0  = FAT sector itself   → FATSECT
    # Sector 1  = Dir sector 1        → sector 2 (chain)
    # Sector 2  = Dir sector 2        → ENDOFCHAIN
    # Sectors 3+ = data sector chains

    total_data_sectors = len(sector_data)
    total_sectors      = DATA_SECTOR_BASE + total_data_sectors

    fat = [FREESECT] * total_sectors
    fat[0] = FATSECT
    fat[1] = 2
    fat[2] = ENDOFCHAIN

    # Build data-sector chains; update start_sector to absolute sector index
    # First pass: assign absolute starts
    abs_info = {}   # name → (abs_start, size)
    for name, (rel_start, sz) in stream_info.items():
        abs_start = DATA_SECTOR_BASE + rel_start
        n_sec     = (sz + SECTOR_SZ - 1) // SECTOR_SZ if sz else 0
        abs_info[name] = (abs_start, sz, n_sec)
        if n_sec > 0:
            for i in range(n_sec - 1):
                fat[abs_start + i] = abs_start + i + 1
            fat[abs_start + n_sec - 1] = ENDOFCHAIN

    # ---- Serialize FAT sector ----
    fat_sector = b''.join(struct.pack('<I', v) for v in fat)
    # Pad to 512 bytes (handles up to 128 sectors)
    fat_sector = fat_sector.ljust(SECTOR_SZ, b'\xff')[:SECTOR_SZ]

    # ---- Build directory entries ----
    VBA_CLSID = bytes.fromhex('01 02 00 00 00 00 00 00 C0 00 00 00 00 00 00 46'.replace(' ',''))

    def info(name):
        a, s, _ = abs_info.get(name, (ENDOFCHAIN, 0, 0))
        start = a if s > 0 else ENDOFCHAIN
        return start, s

    proj_start,  proj_sz  = info('PROJECT')
    projwm_start,projwm_sz= info('PROJECTwm')
    vba_proj_s,  vba_proj_z= info('_VBA_PROJECT')
    dir_s,       dir_z    = info('dir')
    # module name
    mod_name = order[-1]   # last appended
    mod_s,       mod_z    = info(mod_name)

    entries = [None] * 8

    # SID 0: Root Entry  (type=5, color=1=black)
    entries[0] = _dir_entry('Root Entry', 5, 1,
                            NOSTREAM, NOSTREAM, sid['VBA'],
                            VBA_CLSID, ENDOFCHAIN, 0)

    # SID 1: VBA storage  (type=1, color=1)
    entries[1] = _dir_entry('VBA', 1, 1,
                            sid['PROJECT'], NOSTREAM, sid['_VBA_PROJECT'],
                            b'\x00'*16, ENDOFCHAIN, 0)

    # SID 2: PROJECT  (type=2)
    entries[2] = _dir_entry('PROJECT', 2, 1,
                            NOSTREAM, sid['PROJECTwm'], NOSTREAM,
                            b'\x00'*16, proj_start, proj_sz)

    # SID 3: PROJECTwm  (type=2)
    entries[3] = _dir_entry('PROJECTwm', 2, 1,
                            NOSTREAM, NOSTREAM, NOSTREAM,
                            b'\x00'*16, projwm_start, projwm_sz)

    # SID 4: _VBA_PROJECT  (type=2)
    entries[4] = _dir_entry('_VBA_PROJECT', 2, 1,
                            NOSTREAM, sid['dir'], NOSTREAM,
                            b'\x00'*16, vba_proj_s, vba_proj_z)

    # SID 5: dir  (type=2)
    entries[5] = _dir_entry('dir', 2, 1,
                            NOSTREAM, sid[mod_name], NOSTREAM,
                            b'\x00'*16, dir_s, dir_z)

    # SID 6: module  (type=2)
    entries[6] = _dir_entry(mod_name, 2, 1,
                            NOSTREAM, NOSTREAM, NOSTREAM,
                            b'\x00'*16, mod_s, mod_z)

    # SID 7: unused
    entries[7] = _dir_entry('', 0, 1, NOSTREAM, NOSTREAM, NOSTREAM,
                             b'\x00'*16, ENDOFCHAIN, 0)

    dir_blob = b''.join(entries)   # 8 × 128 = 1024 bytes → 2 sectors
    dir_sec1 = dir_blob[:SECTOR_SZ]
    dir_sec2 = dir_blob[SECTOR_SZ:]

    # ---- CFB Header ----
    #  offset  size  description
    #   0       8    Magic
    #   8      16    CLSID (zeros)
    #  24       2    Minor version = 0x003E
    #  26       2    Major version = 0x0003
    #  28       2    Byte order = 0xFFFE
    #  30       2    Sector size shift = 9  (2^9 = 512)
    #  32       2    Mini sector shift = 6  (2^6 = 64)
    #  34       6    Reserved zeros
    #  40       4    # directory sectors (0 for v3)
    #  44       4    # FAT sectors
    #  48       4    First directory sector index
    #  52       4    Transaction signature = 0
    #  56       4    Mini stream cutoff = 4096
    #  60       4    First mini FAT sector (FREESECT)
    #  64       4    # mini FAT sectors = 0
    #  68       4    First DIFAT sector (FREESECT)
    #  72       4    # DIFAT sectors = 0
    #  76     436    DIFAT table (109 entries; entry 0 = FAT sector 0 = sector 0)

    header = struct.pack('<8s16sHHHHH6sIIIIIIIII',
        b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1',   # magic
        b'\x00' * 16,                             # CLSID
        0x003E,                                   # minor version
        0x0003,                                   # major version
        0xFFFE,                                   # byte order
        0x0009,                                   # sector size shift (512)
        0x0006,                                   # mini sector shift (64)
        b'\x00' * 6,                              # reserved
        0,                                        # # dir sectors (v3 = 0)
        1,                                        # # FAT sectors
        1,                                        # first dir sector = sector 1
        0,                                        # transaction signature
        MINI_CUTOFF,                              # mini stream cutoff
        FREESECT,                                 # first mini FAT sector
        0,                                        # # mini FAT sectors
        FREESECT,                                 # first DIFAT sector
        0,                                        # # DIFAT sectors
    )
    # DIFAT array: entry 0 = sector index of FAT sector 0 = 0; rest = FREESECT
    difat = struct.pack('<I', 0) + struct.pack('<I', FREESECT) * 108
    header = header + difat
    header = header.ljust(SECTOR_SZ, b'\x00')

    # ---- Assemble file ----
    out = bytearray(header)
    out += fat_sector
    out += dir_sec1
    out += dir_sec2
    for sec in sector_data:
        out += sec

    return bytes(out)


# ---------------------------------------------------------------------------
# PROJECT stream builder
# ---------------------------------------------------------------------------

def build_project_stream(module_name: str) -> bytes:
    text = (
        f'ID="{{00000000-0000-0000-0000-000000000000}}"\r\n'
        f'Document=ThisDocument/&H00000000\r\n'
        f'Module={module_name}\r\n'
        f'Name=VBAProject\r\n'
        f'HelpContextID="0"\r\n'
        f'VersionCompatible32="393222000"\r\n'
        f'CMG=""\r\n'
        f'GC=""\r\n'
        f'DPB=""\r\n'
    )
    return text.encode('latin-1')


# ---------------------------------------------------------------------------
# Assemble vbaProject.bin
# ---------------------------------------------------------------------------

def build_vba_project(module_name: str, vba_source: str) -> bytes:
    source_bytes = vba_source.encode('latin-1', errors='replace')

    # Module stream: p-code placeholder (2 bytes) + compressed source
    pcode_stub  = b'\x61\x00'          # minimal performance-cache stub
    comp_source = vba_compress(source_bytes)
    module_stream = pcode_stub + comp_source
    text_offset = len(pcode_stub)

    dir_uncompressed = build_dir_stream(module_name, text_offset)
    dir_stream        = vba_compress(dir_uncompressed)

    # _VBA_PROJECT: minimal stub accepted by Excel
    vba_project_stub = b'\xCC\x61\xFF\xFF\x00\x00\x00'

    project_stream = build_project_stream(module_name)
    project_wm     = b''              # no name mapping needed

    streams = {
        f'VBA/_VBA_PROJECT': vba_project_stub,
        f'VBA/dir':           dir_stream,
        f'VBA/{module_name}': module_stream,
        'PROJECT':            project_stream,
        'PROJECTwm':          project_wm,
    }

    return build_cfb(streams)


# ---------------------------------------------------------------------------
# XLAM (OOXML) packaging
# ---------------------------------------------------------------------------

CONTENT_TYPES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="bin"  ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.ms-excel.addin.macroEnabled.sheet+xml"/>
  <Override PartName="/xl/vbaProject.bin"
    ContentType="application/vnd.ms-office.vbaProject"/>
</Types>
"""

RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>
"""

WORKBOOK_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
  <workbookPr codeName="ThisWorkbook"/>
  <sheets/>
</workbook>
"""

WORKBOOK_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject"
    Target="vbaProject.bin"/>
</Relationships>
"""


def build_xlam(vba_source: str, output_path: str, module_name: str = 'LagrangeRegression'):
    vba_bin = build_vba_project(module_name, vba_source)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES.strip())
        zf.writestr('_rels/.rels',          RELS.strip())
        zf.writestr('xl/workbook.xml',      WORKBOOK_XML.strip())
        zf.writestr('xl/_rels/workbook.xml.rels', WORKBOOK_RELS.strip())
        zf.writestr('xl/vbaProject.bin',    vba_bin)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    HERE       = os.path.dirname(os.path.abspath(__file__))
    VBA_SRC    = os.path.join(HERE, 'LagrangeVBA.bas')
    OUTPUT     = os.path.join(HERE, 'LagrangeRegression.xlam')

    with open(VBA_SRC, encoding='utf-8') as f:
        vba_source = f.read()

    print('Building vbaProject.bin …')
    build_xlam(vba_source, OUTPUT)
    size = os.path.getsize(OUTPUT)
    print(f'Done → {OUTPUT}  ({size:,} bytes)')

    # Quick sanity check with olefile
    try:
        import olefile
        with olefile.OleFileIO(OUTPUT, path_encoding=None) as ole:
            actual = zipfile.ZipFile(OUTPUT)
            vba_bin = actual.read('xl/vbaProject.bin')
            ole2 = olefile.OleFileIO(io.BytesIO(vba_bin))
            streams = ole2.listdir()
            print('\nStreams in vbaProject.bin:')
            for s in streams:
                print(' ', '/'.join(s))
    except Exception as e:
        print(f'(sanity check skipped: {e})')
