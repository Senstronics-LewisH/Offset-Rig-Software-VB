# sync_all_twinprojs.py
# Programmatic parser and packer to sync .bas and .frm modifications back into Project1.twinproj

import struct
import os
import json

class Node:
    def __init__(self, item_type, name, meta, val, data_or_children):
        self.item_type = item_type
        self.name = name
        self.meta = meta
        self.val = val
        self.data_or_children = data_or_children

def parse_item(offset, data):
    if offset >= len(data):
        return None, offset
    
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
            curr += 4
        node = Node(item_type, name, meta, val, file_data)
        
    return node, curr

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
        res.extend(b"\x00\x00\x00\x00")
    return res

def make_frm_twin_content(old_twin_bytes, frm_disk_path):
    old_twin = old_twin_bytes.decode("utf-8", errors="ignore")
    guid_lines = []
    for line in old_twin.splitlines():
        if line.strip().startswith("["):
            guid_lines.append(line)
        else:
            break
            
    header = "\r\n".join(guid_lines)
    
    with open(frm_disk_path, "r", encoding="utf-8", errors="ignore") as f:
        frm_content = f.read()
        
    idx = frm_content.find("Attribute VB_Name")
    if idx == -1:
        raise ValueError(f"Could not find Attribute VB_Name in {frm_disk_path}")
        
    code_part = frm_content[idx:]
    lines = code_part.splitlines()
    indented_lines = []
    
    class_name = "Form"
    for line in lines:
        if "Attribute VB_Name" in line:
            parts = line.split("=")
            if len(parts) >= 2:
                class_name = parts[1].replace('"', '').strip()
                
    for line in lines:
        if line.strip() == "":
            indented_lines.append("")
        else:
            indented_lines.append("    " + line)
            
    body = "\r\n".join(indented_lines)
    return f"{header}\r\nClass {class_name}\r\n{body}\r\nEnd Class\r\n".encode("utf-8")

def sync_project(project_dir, fix_references_level=0):
    print(f"\nSyncing twinproj in: {project_dir}")
    twinproj_path = os.path.join(project_dir, "Project1.twinproj")
    backup_path = twinproj_path + ".bak"
    
    if not os.path.exists(twinproj_path):
        print(f"ERROR: twinproj file not found at {twinproj_path}")
        return False
        
    # Create a fresh backup of the current twinproj file before writing
    try:
        with open(twinproj_path, "rb") as sf, open(backup_path, "wb") as df:
            df.write(sf.read())
    except Exception as e:
        print(f"Warning: Could not create backup: {e}")
            
    source_path = twinproj_path
    with open(source_path, "rb") as f:
        orig_data = f.read()
        
    magic = orig_data[0:4]
    root_node, _ = parse_item(4, orig_data)
    
    logfile_path = os.path.join(project_dir, "LogFile.bas")
    offsetcheck_path = os.path.join(project_dir, "OffsetCheck.bas")
    controlpsu_path = os.path.join(project_dir, "ControlPSU.bas")
    dmmcontrol_path = os.path.join(project_dir, "DigitalMultimeterControl.bas")
    inputsoutputs_path = os.path.join(project_dir, "InputsOutputs.bas")
    csvfile_path = os.path.join(project_dir, "CSVFile.bas")
    vision_path = os.path.join(project_dir, "Vision.bas")
    ieeevb_path = os.path.join(project_dir, "ieeevb.bas")
    barcode_path = os.path.join(project_dir, "Barcode.bas")
    
    mainfrm_path = os.path.join(project_dir, "MainForm.frm")
    cableusagefrm_path = os.path.join(project_dir, "CableUsage.frm")
    pwordfrm_path = os.path.join(project_dir, "PWord.frm")
    
    with open(logfile_path, "rb") as f:
        logfile_content = f.read()
    with open(offsetcheck_path, "rb") as f:
        offsetcheck_content = f.read()
    with open(controlpsu_path, "rb") as f:
        controlpsu_content = f.read()
    with open(dmmcontrol_path, "rb") as f:
        dmmcontrol_content = f.read()
    with open(inputsoutputs_path, "rb") as f:
        inputsoutputs_content = f.read()
    with open(csvfile_path, "rb") as f:
        csvfile_content = f.read()
    with open(vision_path, "rb") as f:
        vision_content = f.read()
    with open(ieeevb_path, "rb") as f:
        ieeevb_content = f.read()
    with open(barcode_path, "rb") as f:
        barcode_content = f.read()
        
    def update_nodes(node):
        if node.name == "LogFile.bas":
            node.data_or_children = logfile_content
            node.val = len(logfile_content)
            print("  Updated LogFile.bas")
        elif node.name == "OffsetCheck.bas":
            node.data_or_children = offsetcheck_content
            node.val = len(offsetcheck_content)
            print("  Updated OffsetCheck.bas")
        elif node.name == "ControlPSU.bas":
            node.data_or_children = controlpsu_content
            node.val = len(controlpsu_content)
            print("  Updated ControlPSU.bas")
        elif node.name == "DigitalMultimeterControl.bas":
            node.data_or_children = dmmcontrol_content
            node.val = len(dmmcontrol_content)
            print("  Updated DigitalMultimeterControl.bas")
        elif node.name == "InputsOutputs.bas":
            node.data_or_children = inputsoutputs_content
            node.val = len(inputsoutputs_content)
            print("  Updated InputsOutputs.bas")
        elif node.name == "CSVFile.bas":
            node.data_or_children = csvfile_content
            node.val = len(csvfile_content)
            print("  Updated CSVFile.bas")
        elif node.name == "Vision.bas":
            node.data_or_children = vision_content
            node.val = len(vision_content)
            print("  Updated Vision.bas")
        elif node.name == "ieeevb.bas":
            node.data_or_children = ieeevb_content
            node.val = len(ieeevb_content)
            print("  Updated ieeevb.bas")
        elif node.name == "Barcode.bas":
            node.data_or_children = barcode_content
            node.val = len(barcode_content)
            print("  Updated Barcode.bas")
        elif node.name == "MainForm.frm.twin":
            twin_data = make_frm_twin_content(node.data_or_children, mainfrm_path)
            node.data_or_children = twin_data
            node.val = len(twin_data)
            print("  Updated MainForm.frm.twin")
        elif node.name == "MainForm.frm.tbform":
            tbform_data = node.data_or_children
            tbform_data = tbform_data.replace(b'"Caption": "POST CALIBRATION OFFSET CHECKER V34"', b'"Caption": "POST CALIBRATION OFFSET CHECKER"')
            node.data_or_children = tbform_data
            node.val = len(tbform_data)
            print("  Updated MainForm.frm.tbform (removed V34 default)")
        elif node.name == "CableUsage.frm.twin":
            twin_data = make_frm_twin_content(node.data_or_children, cableusagefrm_path)
            node.data_or_children = twin_data
            node.val = len(twin_data)
            print("  Updated CableUsage.frm.twin")
        elif node.name == "PWord.frm.twin":
            twin_data = make_frm_twin_content(node.data_or_children, pwordfrm_path)
            node.data_or_children = twin_data
            node.val = len(twin_data)
            print("  Updated PWord.frm.twin")
        elif node.name == "Settings":
            settings_json = json.loads(node.data_or_children.decode("utf-8"))
            refs = settings_json.get("project.references", [])
            new_refs = []
            for ref in refs:
                name = ref.get("name", "")
                if fix_references_level == 2 and name == "Keithley SCPI IVI Driver":
                    print("  Removed reference: Keithley SCPI IVI Driver")
                    continue
                elif fix_references_level == 2 and name == "IVI Keithley2230 1.0 Type Library":
                    ref["path32"] = r"C:\Program Files (x86)\IVI Foundation\IVI\Bin\KE2230.dll"
                    ref["path64"] = r"C:\Program Files (x86)\IVI Foundation\IVI\Bin\KE2230.dll"
                    print("  Fixed reference: IVI Keithley2230 1.0 Type Library")
                
                if fix_references_level >= 1 and name == "Microsoft Excel 16.0 Object Library":
                    ref["path32"] = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
                    ref["path64"] = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
                    print("  Cleaned reference: Microsoft Excel 16.0 Object Library")
                new_refs.append(ref)
            settings_json["project.references"] = new_refs
            new_settings_data = json.dumps(settings_json).encode("utf-8")
            node.data_or_children = new_settings_data
            node.val = len(new_settings_data)
            print("  Updated Settings references")
            
        if isinstance(node.data_or_children, list):
            for child in node.data_or_children:
                update_nodes(child)
                
    update_nodes(root_node)
    
    rebuilt_data = magic + pack_node(root_node)
    with open(twinproj_path, "wb") as f:
        f.write(rebuilt_data)
    print(f"Successfully wrote rebuilt twinproj to: {twinproj_path}")
    return True

if __name__ == "__main__":
    dir_target = os.path.dirname(os.path.abspath(__file__))
    sync_project(dir_target, fix_references_level=1)
