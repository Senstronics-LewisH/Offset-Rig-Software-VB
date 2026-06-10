# twinBASIC Project File (.twinproj) Synchronization Guide

When migrating or maintaining a Visual Basic 6 project inside **twinBASIC**, the project can be saved in a monolithic `.twinproj` format. This guide documents the structure of the `.twinproj` archive format and explains how to synchronize modifications made to external VB6 `.bas` and `.frm` files back into twinBASIC.

## Structure of `.twinproj` File

The `.twinproj` file is a custom binary tree archive. 

### 1. File Header
The file begins with a 4-byte signature/magic number:
* **Magic Signature:** `1C A5 0B EA` (in hex)

### 2. Recursive Node Structure
Directly following the magic signature is the tree of project nodes. Each node represents either a **Directory** (like the root `Project1`, `Sources`, or `Resources`) or a **File** (like `.bas` modules, `.frm.tbform` assets, or settings).

Every node follows a strict binary layout:

| Offset | Size (Bytes) | Data Type | Field Description |
| :--- | :--- | :--- | :--- |
| `+0` | 2 | UInt16 (Little Endian) | **Item Type** (`1` = File / Root, `2` = Directory) |
| `+2` | 4 | UInt32 (Little Endian) | **Name Length** ($N$) |
| `+6` | $N$ | UTF-8 String | **Node Name** (e.g., `OffsetCheck.bas` or `Sources`) |
| `+6+N` | 13 | Bytes | **Metadata Block** (various internal flags/attributes) |
| `+19+N` | 4 | UInt32 (Little Endian) | **Value/Size Field** ($V$) <br> • For Files: Size of file data in bytes <br> • For Directories: Number of child nodes |

Depending on the **Item Type**, the node continues with:

#### A. File Node (`Item Type = 1`)
1. **File Data:** Directly following the size field $V$, there are $V$ bytes of raw file content (representing the source code).
2. **Padding:** Immediately after the file data, a fixed **4-byte null padding block** (`00 00 00 00`) is appended.
3. The next node starts immediately after the padding block.

#### B. Directory Node (`Item Type = 2`)
1. **Sub-Items:** There is **no file data** or padding. Instead, the next $V$ child nodes are packed sequentially, one after another.
2. The next node starts immediately after the last child node is fully parsed.

---

## Programmatic Synchronization Process

To edit source code in a `.twinproj` project from outside the twinBASIC IDE, you must parse the binary structure, replace the source code byte buffer for the target files, update the size fields ($V$), and rebuild the archive.

### Python Sync Algorithm Reference
The following logic (implemented in `sync_twinproj.py` in the brain scratchpad) provides a reliable, structure-preserving round-trip parser:

```python
import struct

# 1. Parsing Tree Nodes
def parse_item(offset, data):
    item_type = struct.unpack("<H", data[offset:offset+2])[0]
    name_len = struct.unpack("<I", data[offset+2:offset+6])[0]
    name = data[offset+6:offset+6+name_len].decode("utf-8")
    
    curr = offset + 6 + name_len
    meta = data[curr:curr+13]
    curr += 13
    
    val = struct.unpack("<I", data[curr:curr+4])[0]
    curr += 4
    
    if name == "Project1" or item_type == 2:
        children = []
        for _ in range(val):
            child, curr = parse_item(curr, data)
            children.append(child)
        node = Node(item_type, name, meta, val, children)
    else:
        file_data = data[curr:curr+val]
        curr += val
        if curr + 4 <= len(data) and data[curr:curr+4] == b"\x00\x00\x00\x00":
            curr += 4 # Skip file padding
        node = Node(item_type, name, meta, val, file_data)
        
    return node, curr

# 2. Packing Tree Nodes
def pack_node(node):
    res = bytearray()
    res.extend(struct.pack("<H", node.item_type))
    name_bytes = node.name.encode("utf-8")
    res.extend(struct.pack("<I", len(name_bytes)))
    res.extend(name_bytes)
    res.extend(node.meta)
    
    if node.name == "Project1" or node.item_type == 2:
        res.extend(struct.pack("<I", len(node.data_or_children)))
        for child in node.data_or_children:
            res.extend(pack_node(child))
    else:
        res.extend(struct.pack("<I", len(node.data_or_children)))
        res.extend(node.data_or_children)
        res.extend(b"\x00\x00\x00\x00") # Padding
    return res
```

## Developer Guidelines
1. **Always backup** `Project1.twinproj` before running any automated synchronization script.
2. **Never change** the order of children inside directories, as this can affect twinBASIC IDE display lists.
3. If you add or remove modules/forms, you must update the parent folder's child count (`val` field of the directory node).
