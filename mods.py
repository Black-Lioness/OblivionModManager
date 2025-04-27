import zipfile
import os
import sys
import shutil
import tempfile
import ctypes
import time
import re
from pathlib import Path # Using pathlib for easier path manipulation

# Attempt to import required libraries for RAR and 7z
try:
    import rarfile
    RAR_SUPPORT = True
except ImportError:
    RAR_SUPPORT = False

try:
    import py7zr
    SEVENZIP_SUPPORT = True
except ImportError:
    SEVENZIP_SUPPORT = False

# --- Configuration ---
GAME_BASE_PATH = r"C:\Program Files (x86)\Steam\steamapps\common\Oblivion Remastered"
OBLIVION_CONTENT_PATH = Path(GAME_BASE_PATH) / "OblivionRemastered" / "Content"
# ESP Mod Paths
ESP_DATA_PATH = OBLIVION_CONTENT_PATH / "Dev" / "ObvData" / "Data"
PLUGINS_TXT_PATH = ESP_DATA_PATH / "plugins.txt"
# Pak Mod Path
PAK_MODS_PATH = OBLIVION_CONTENT_PATH / "Paks" / "~mods"

# Default plugins that should always be at the top and not reordered by the user
# Using a set for efficient lookup, converting to lowercase for case-insensitivity
DEFAULT_PLUGINS = {
    "oblivion.esm", # Base game master file
    # Official DLCs (adjust if Oblivion Remastered names them differently)
    "dlcbattlehorncastle.esp",
    "dlcfrostcrag.esp",
    "dlchorsearmor.esp",
    "dlcmehrunesrazor.esp",
    "dlcorrery.esp",
    "dlcshiveringisles.esp", # Expansion often treated as DLC
    "dlcspelltomes.esp",
    "dlcthievesden.esp",
    "dlcvilelair.esp",
    "knights.esp", # Knights of the Nine
    # Added Altar mods as per user request - assume they are 'default' for this setup
    "altarespmain.esp",
    "altardeluxe.esp",
    "altaresplocal.esp",
}

SUPPORTED_EXTENSIONS = ['.zip']
if RAR_SUPPORT: SUPPORTED_EXTENSIONS.append('.rar')
if SEVENZIP_SUPPORT: SUPPORTED_EXTENSIONS.append('.7z')

PAK_EXTENSIONS = ('.pak', '.ucas', '.utoc')
# Files/Dirs to ignore when detecting single-folder structure
IGNORE_ITEMS = {'.ds_store', '__macosx', '.git', '.gitignore', 'thumbs.db'}


# --- Helper Functions (Admin, Path Checks) ---

def is_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception: return False

def run_as_admin():
    if sys.platform == 'win32':
        try:
            result = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join([f'"{arg}"' for arg in sys.argv]), os.getcwd(), 1)
            if result > 32: return True
            else: print(f"ERROR: Failed elevation (Code: {result}). Run manually as Admin."); return False
        except Exception as e: print(f"ERROR: Elevation attempt failed: {e}"); return False
    else: print("ERROR: Auto-elevation only on Windows."); return False

def is_supported_archive(filepath):
    return Path(filepath).suffix.lower() in SUPPORTED_EXTENSIONS

def ensure_directory_exists(dir_path):
    try:
        Path(dir_path).mkdir(parents=True, exist_ok=True)
        return True
    except OSError as e:
        print(f"ERROR: Could not create directory '{dir_path}': {e}")
        return False

# --- Archive Reading Functions ---

def _get_archive_member_list(archive_path):
    ext = Path(archive_path).suffix.lower()
    members = []
    try:
        if ext == '.zip':
            with zipfile.ZipFile(archive_path, 'r') as arc:
                for mi in arc.infolist(): members.append((mi.filename.replace('\\', '/'), mi.is_dir()))
        elif ext == '.rar' and RAR_SUPPORT:
            with rarfile.RarFile(archive_path, 'r') as arc:
                for mi in arc.infolist(): members.append((mi.filename.replace('\\', '/'), getattr(mi, 'is_dir', lambda: False)()))
        elif ext == '.7z' and SEVENZIP_SUPPORT:
             with py7zr.SevenZipFile(archive_path, mode='r') as arc:
                 for mi in arc.list(): members.append((mi.filename.replace('\\', '/'), mi.is_directory))
        else: raise ValueError(f"Unsupported/unavailable type: {ext}")
    except FileNotFoundError: raise FileNotFoundError(f"Archive not found: '{archive_path}'")
    except (zipfile.BadZipFile, rarfile.BadRarFile, py7zr.exceptions.Bad7zFile, ValueError) as e: raise ValueError(f"Error reading '{Path(archive_path).name}': {e}")
    except rarfile.UnrarNotFound: raise RuntimeError("Unrar executable not found.")
    except Exception as e: raise RuntimeError(f"Unexpected error reading '{Path(archive_path).name}': {e}")
    return members

def _detect_single_folder_prefix(member_list):
    top_level_dirs = set()
    has_relevant_files_at_root = False
    for path_str, is_dir in member_list:
        path_parts = path_str.strip('/').split('/')
        if not path_parts or path_parts[0].lower() in IGNORE_ITEMS: continue
        if len(path_parts) == 1:
            if is_dir: top_level_dirs.add(path_parts[0])
            elif path_parts[0].lower().endswith(('.esp',) + PAK_EXTENSIONS): has_relevant_files_at_root = True
        elif len(path_parts) > 1: top_level_dirs.add(path_parts[0])
    valid_top_dirs = [d for d in top_level_dirs if d.lower() not in IGNORE_ITEMS]
    if len(valid_top_dirs) == 1 and not has_relevant_files_at_root:
        single_folder_name = valid_top_dirs[0]
        all_inside = all(p.strip('/').split('/')[0] == single_folder_name for p, _ in member_list if p.strip('/') and p.strip('/').split('/')[0].lower() not in IGNORE_ITEMS)
        if all_inside:
            print(f"INFO: Detected single content folder: '{single_folder_name}'")
            return single_folder_name + "/"
    return None

def find_esps_in_archive(archive_path):
    try:
        members = _get_archive_member_list(archive_path)
        prefix = _detect_single_folder_prefix(members)
        esps = []
        for path_str, is_dir in members:
            if is_dir: continue
            path_to_check = path_str
            if prefix and not path_str.startswith(prefix): continue
            if path_to_check.lower().endswith('.esp'): esps.append(path_str)
        return esps
    except (FileNotFoundError, ValueError, RuntimeError) as e: raise e
    except Exception as e: raise RuntimeError(f"Error searching ESPs in '{Path(archive_path).name}': {e}")

def find_pak_sets_in_archive(archive_path):
    pak_sets = []
    found_bases = {}
    try:
        members = _get_archive_member_list(archive_path)
        prefix = _detect_single_folder_prefix(members)
        for path_str, is_dir in members:
             if is_dir: continue
             path_to_check = path_str
             if prefix and not path_str.startswith(prefix): continue
             lower_path = path_to_check.lower()
             if lower_path.endswith(PAK_EXTENSIONS):
                  path_obj = Path(path_to_check)
                  base_key = (path_obj.parent / path_obj.stem).as_posix().lower()
                  if base_key not in found_bases: found_bases[base_key] = {'pak': None, 'ucas': None, 'utoc': None}
                  ext = path_obj.suffix.lower()
                  if ext == '.pak': found_bases[base_key]['pak'] = path_str
                  elif ext == '.ucas': found_bases[base_key]['ucas'] = path_str
                  elif ext == '.utoc': found_bases[base_key]['utoc'] = path_str
        for base_key, files in found_bases.items():
            if files['pak'] and files['ucas'] and files['utoc']: pak_sets.append(files)
        return pak_sets
    except (FileNotFoundError, ValueError, RuntimeError) as e: raise e
    except Exception as e: raise RuntimeError(f"Error searching Pak sets in '{Path(archive_path).name}': {e}")

# --- Archive Extraction Functions ---

def _extract_files_from_archive(archive_path, files_to_extract, target_dir):
    ext = Path(archive_path).suffix.lower()
    extracted_paths = []
    if not ensure_directory_exists(target_dir): raise RuntimeError(f"Failed target dir: {target_dir}")
    try:
        if ext == '.zip':
            with zipfile.ZipFile(archive_path, 'r') as arc:
                for member in files_to_extract: arc.extract(member, target_dir); extracted_paths.append(Path(target_dir) / member)
        elif ext == '.rar' and RAR_SUPPORT:
             with rarfile.RarFile(archive_path, 'r') as arc:
                 for member in files_to_extract: arc.extract(member, target_dir); extracted_paths.append(Path(target_dir) / member)
        elif ext == '.7z' and SEVENZIP_SUPPORT:
             with py7zr.SevenZipFile(archive_path, mode='r') as arc:
                  arc.extract(path=target_dir, targets=files_to_extract)
                  for member in files_to_extract: extracted_paths.append(Path(target_dir) / member)
        else: raise ValueError(f"Unsupported type for extraction: {ext}")
        for p in extracted_paths:
             if not p.exists(): print(f"WARNING: Extracted file missing: {p}")
    except FileNotFoundError: raise FileNotFoundError(f"Archive not found: '{archive_path}'")
    except (KeyError, zipfile.BadZipFile, rarfile.BadRarFile, py7zr.exceptions.Bad7zFile, ValueError) as e: raise ValueError(f"Error extracting from '{Path(archive_path).name}': {e}")
    except rarfile.UnrarNotFound: raise RuntimeError("Unrar executable not found.")
    except Exception as e: raise RuntimeError(f"Unexpected extraction error from '{Path(archive_path).name}': {e}")
    return [str(p) for p in extracted_paths]

def extract_esp_to_temp(archive_path, esp_path_in_archive, temp_dir):
    try:
        temp_path = Path(temp_dir)
        extracted_list = _extract_files_from_archive(archive_path, [esp_path_in_archive], str(temp_path))
        extracted_path = Path(extracted_list[0])
        if not extracted_path.exists(): raise FileNotFoundError(f"Failed extraction: '{esp_path_in_archive}'")
        return str(extracted_path)
    except Exception as e: raise RuntimeError(f"Failed ESP extraction '{esp_path_in_archive}': {e}")

# --- Core Mod Manager Functions ---

def select_esp_from_list(esp_files, action="install"):
    if not esp_files: return None
    if len(esp_files) == 1: print(f"Found single ESP: {Path(esp_files[0]).name}"); return esp_files[0]
    print(f"\nMultiple ESP files found:"); [print(f"  {i + 1}: {esp}") for i, esp in enumerate(esp_files)]
    while True:
        try:
            choice = input(f"Enter ESP number (1-{len(esp_files)}), or 'c' cancel: ").lower().strip()
            if choice == 'c': print("Cancelled."); return None
            index = int(choice) - 1
            if 0 <= index < len(esp_files): print(f"Selected: {Path(esp_files[index]).name}"); return esp_files[index]
            else: print("Invalid choice.")
        except ValueError: print("Invalid input.")
        except (EOFError, KeyboardInterrupt): print("\nCancelled."); sys.exit(1)

def select_pak_set_from_list(pak_sets, action="install"):
    if not pak_sets: return None
    pak_filenames = [Path(s['pak']).name for s in pak_sets]
    if len(pak_sets) == 1: print(f"Found single Pak set: {pak_filenames[0]}"); return pak_sets[0]
    print(f"\nMultiple Pak sets found:"); [print(f"  {i + 1}: {name}") for i, name in enumerate(pak_filenames)]
    while True:
        try:
            choice = input(f"Enter Pak set number (1-{len(pak_sets)}), or 'c' cancel: ").lower().strip()
            if choice == 'c': print("Cancelled."); return None
            index = int(choice) - 1
            if 0 <= index < len(pak_sets): print(f"Selected Pak set: {pak_filenames[index]}"); return pak_sets[index]
            else: print("Invalid choice.")
        except ValueError: print("Invalid input.")
        except (EOFError, KeyboardInterrupt): print("\nCancelled."); sys.exit(1)

def install_esp_file(extracted_esp_path_str, esp_filename_in_archive):
    esp_basename = Path(esp_filename_in_archive).name
    dest_path = ESP_DATA_PATH / esp_basename
    extracted_esp_path = Path(extracted_esp_path_str)
    print(f"ESP Destination: {dest_path}")
    if not ensure_directory_exists(ESP_DATA_PATH): return False
    if dest_path.exists():
        print(f"WARNING: '{dest_path.name}' exists."); ovr = input("Overwrite? (y/n): ").lower().strip()
        if ovr != 'y': print("Cancelled."); return False
    try:
        print(f"Copying '{esp_basename}'..."); shutil.copy2(str(extracted_esp_path), str(dest_path)); print("Copy ok."); return True
    except PermissionError: print(f"ERROR: Permission denied copying to '{ESP_DATA_PATH}'."); return False
    except Exception as e: print(f"ERROR: Failed copy ESP: {e}"); return False

def read_plugins_file():
    plugins = []
    if not PLUGINS_TXT_PATH.exists(): return plugins
    try:
        with open(PLUGINS_TXT_PATH, 'r') as f: plugins = [ln.strip() for ln in f if ln.strip() and not ln.strip().startswith('#')]
    except PermissionError: print(f"ERROR: Permission denied reading '{PLUGINS_TXT_PATH}'."); raise
    except Exception as e: print(f"ERROR: Read fail '{PLUGINS_TXT_PATH}': {e}"); raise
    return plugins

def write_plugins_file(plugins_list):
    if not ensure_directory_exists(PLUGINS_TXT_PATH.parent): return False
    try:
        with open(PLUGINS_TXT_PATH, 'w') as f:
            for plugin in plugins_list:
                if plugin: f.write(plugin + '\n')
        return True
    except PermissionError: print(f"ERROR: Permission denied writing '{PLUGINS_TXT_PATH}'."); return False
    except Exception as e: print(f"ERROR: Write fail '{PLUGINS_TXT_PATH}': {e}"); return False

def register_plugin(esp_filename_in_archive):
    esp_basename = Path(esp_filename_in_archive).name
    print(f"Registering ESP '{esp_basename}'...")
    try: current_plugins = read_plugins_file()
    except Exception: return False
    if any(esp_basename.lower() == p.lower() for p in current_plugins): print(f"'{esp_basename}' already registered."); return True
    current_plugins.append(esp_basename)
    if write_plugins_file(current_plugins): print(f"Added '{esp_basename}' to {PLUGINS_TXT_PATH.name}."); return True
    else: return False

def remove_plugin_from_registry(esp_basename):
    print(f"Unregistering ESP '{esp_basename}'...")
    try: current_plugins = read_plugins_file()
    except Exception: return False
    original_len = len(current_plugins)
    updated_plugins = [p for p in current_plugins if p.lower() != esp_basename.lower()]
    if len(updated_plugins) == original_len: print(f"WARNING: '{esp_basename}' not in registry."); return True
    if write_plugins_file(updated_plugins): print(f"Removed '{esp_basename}' from registry."); return True
    else: return False

def delete_esp_file(esp_basename):
    esp_file_path = ESP_DATA_PATH / esp_basename
    print(f"Deleting ESP: {esp_file_path}")
    if not esp_file_path.exists(): print(f"WARNING: ESP File '{esp_basename}' not found."); return True
    try: esp_file_path.unlink(); print(f"Deleted '{esp_basename}'."); return True
    except PermissionError: print(f"ERROR: Permission denied deleting '{esp_basename}'."); return False
    except Exception as e: print(f"ERROR: Failed delete '{esp_basename}': {e}"); return False

def get_installed_custom_mods():
    try: return [p for p in read_plugins_file() if p.lower() not in DEFAULT_PLUGINS]
    except Exception: print("Could not read plugins.txt."); return []

def get_installed_pak_mods():
    pak_mods = []
    if not PAK_MODS_PATH.is_dir(): return pak_mods
    try:
        for item in PAK_MODS_PATH.iterdir():
            if item.is_file() and item.suffix.lower() == '.pak': pak_mods.append(item.name)
    except PermissionError: print(f"ERROR: Perms reading Pak dir '{PAK_MODS_PATH}'.")
    except Exception as e: print(f"ERROR: List Pak fail '{PAK_MODS_PATH}': {e}")
    return pak_mods

def delete_pak_mod_files(pak_basename):
    print(f"Deleting Pak set for: {pak_basename}")
    base = Path(pak_basename).stem
    files_del = [PAK_MODS_PATH / pak_basename, PAK_MODS_PATH / (base + ".ucas"), PAK_MODS_PATH / (base + ".utoc")]
    all_ok = True
    for fp in files_del:
        if fp.exists():
            try: fp.unlink(); print(f"  Deleted: {fp.name}")
            except PermissionError: print(f"  ERROR: Perms delete '{fp.name}'."); all_ok = False
            except Exception as e: print(f"  ERROR: Failed delete '{fp.name}': {e}"); all_ok = False
        else:
             if fp.suffix.lower() != '.pak': print(f"  INFO: Missing companion '{fp.name}'.")
             else: print(f"  WARNING: Main Pak missing '{fp.name}'."); all_ok = False
    return all_ok


# --- Processing Functions ---

def process_esp_installation(archive_path):
    print(f"\n--- ESP Install: {Path(archive_path).name} ---")
    temp_dir_obj = None
    try:
        esps = find_esps_in_archive(archive_path)
        if not esps: print("ERROR: No .esp files found."); return False
        selected_esp = select_esp_from_list(esps)
        if not selected_esp: return False
        print(f"Extracting '{Path(selected_esp).name}'...")
        temp_dir_obj = Path(tempfile.mkdtemp(prefix="obv_esp_"))
        extracted_path = extract_esp_to_temp(archive_path, selected_esp, str(temp_dir_obj))
        if not install_esp_file(extracted_path, selected_esp): print("--- ESP Install Fail (Copy) ---"); return False
        if not register_plugin(selected_esp):
            print(f"\nWARNING: ESP copied, registry failed."); print(f"  Manually add '{Path(selected_esp).name}' to '{PLUGINS_TXT_PATH.name}'")
            print("--- ESP Install Incomplete (Registry) ---"); return False
        else: print("--- ESP Install OK ---"); return True
    except (ValueError, FileNotFoundError, RuntimeError) as e: print(f"\nERROR ESP Install: {e}"); return False
    except Exception as e: print(f"\nUNEXPECTED ESP Install ERROR: {e}"); import traceback; traceback.print_exc(); return False
    finally:
        if temp_dir_obj and temp_dir_obj.exists():
            try: shutil.rmtree(str(temp_dir_obj))
            except Exception as e: print(f"WARN: Temp cleanup fail '{temp_dir_obj}': {e}")

def process_pak_installation(archive_path):
    print(f"\n--- Pak Install: {Path(archive_path).name} ---")
    try:
        pak_sets = find_pak_sets_in_archive(archive_path)
        if not pak_sets: print("ERROR: No complete Pak sets found."); return False
        selected_set = select_pak_set_from_list(pak_sets)
        if not selected_set: return False
        files_in_arc = [selected_set['pak'], selected_set['ucas'], selected_set['utoc']]
        target_names = [Path(f).name for f in files_in_arc]
        pak_base = Path(selected_set['pak']).name
        print(f"Target Pak dir: {PAK_MODS_PATH}")
        if not ensure_directory_exists(PAK_MODS_PATH): return False
        existing = [f for f in target_names if (PAK_MODS_PATH / f).exists()]
        if existing:
            print(f"WARN: Pak files exist: {', '.join(existing)}"); ovr = input("Overwrite? (y/n): ").lower().strip()
            if ovr != 'y': print("Pak install cancelled."); return False
        print(f"Extracting Paks for '{pak_base}'...")
        _extract_files_from_archive(archive_path, files_in_arc, str(PAK_MODS_PATH))
        all_exist = all((PAK_MODS_PATH / tn).exists() for tn in target_names)
        if not all_exist: print("ERROR: Post-extract verification failed."); print("--- Pak Install Fail (Verify) ---"); return False
        print(f"--- Pak Install OK: '{pak_base}' ---"); return True
    except (ValueError, FileNotFoundError, RuntimeError) as e: print(f"\nERROR Pak Install: {e}"); return False
    except Exception as e: print(f"\nUNEXPECTED Pak Install ERROR: {e}"); import traceback; traceback.print_exc(); return False

def process_esp_uninstallation():
    print("\n--- ESP Mod Uninstallation ---")
    mods = get_installed_custom_mods();
    if not mods: print("No custom ESP mods found."); return True
    print("\nInstalled Custom ESP Mods:"); [print(f"  {i + 1}: {mod}") for i, mod in enumerate(mods)]
    while True:
        try:
            choice = input(f"Uninstall ESP # (1-{len(mods)}), or 'c' cancel: ").lower().strip()
            if choice == 'c': print("Cancelled."); return False
            idx = int(choice) - 1;
            if 0 <= idx < len(mods): esp_to_del = mods[idx]; break
            else: print("Invalid choice.")
        except ValueError: print("Invalid input.")
        except (EOFError, KeyboardInterrupt): print("\nCancelled."); sys.exit(1)
    print(f"\nSelected '{esp_to_del}'."); conf = input(f"CONFIRM uninstall '{esp_to_del}'? (y/n): ").lower().strip()
    if conf != 'y': print("Cancelled."); return False
    reg_ok = remove_plugin_from_registry(esp_to_del)
    file_ok = delete_esp_file(esp_to_del)
    if reg_ok and file_ok: print("\n--- ESP Uninstall OK ---"); return True
    else: print("\n--- ESP Uninstall Fail/Warnings ---"); return False

def process_pak_uninstallation():
    print("\n--- Pak Mod Uninstallation ---")
    paks = get_installed_pak_mods();
    if not paks: print("No Pak mods found."); return True
    print("\nInstalled Pak Mods:"); [print(f"  {i + 1}: {pak}") for i, pak in enumerate(paks)]
    while True:
        try:
            choice = input(f"Uninstall Pak # (1-{len(paks)}), or 'c' cancel: ").lower().strip()
            if choice == 'c': print("Cancelled."); return False
            idx = int(choice) - 1;
            if 0 <= idx < len(paks): pak_to_del = paks[idx]; break
            else: print("Invalid choice.")
        except ValueError: print("Invalid input.")
        except (EOFError, KeyboardInterrupt): print("\nCancelled."); sys.exit(1)
    print(f"\nSelected '{pak_to_del}'."); conf = input(f"CONFIRM uninstall '{pak_to_del}'? (y/n): ").lower().strip()
    if conf != 'y': print("Cancelled."); return False
    files_ok = delete_pak_mod_files(pak_to_del)
    if files_ok: print("\n--- Pak Uninstall OK ---"); return True
    else: print("\n--- Pak Uninstall Fail/Warnings ---"); return False

def process_load_order():
    """Allows user to reorder the custom plugins in plugins.txt."""
    print("\n--- Change ESP Load Order ---")
    try:
        all_plugins = read_plugins_file()
    except Exception:
        print("ERROR: Failed to read plugins.txt. Cannot change load order.")
        return # Cannot proceed

    default_section = []
    custom_section = []
    # Ensure DEFAULT_PLUGINS uses lowercase for comparison
    default_plugins_lower = {p.lower() for p in DEFAULT_PLUGINS}

    # Separate plugins, preserving original relative order within sections
    for plugin in all_plugins:
        if plugin.lower() in default_plugins_lower:
            default_section.append(plugin)
        else:
            custom_section.append(plugin)

    if not custom_section:
        print("No custom ESP mods found in plugins.txt to reorder.")
        return

    print("\nCurrent Custom ESP Load Order:")
    for i, plugin in enumerate(custom_section):
        print(f"  {i + 1}: {plugin}")

    print("\nEnter the numbers corresponding to the desired new order, separated by spaces.")
    print("Example: If current is 1:A 2:B 3:C, and you want C B A, enter: 3 2 1")

    while True:
        try:
            new_order_input = input("New order: ").strip()
            if not new_order_input:
                print("No order entered. Cancelling.")
                return

            # Parse the input numbers
            new_indices_str = re.findall(r'\d+', new_order_input) # Find all sequences of digits
            if not new_indices_str:
                print("Invalid input format. Please enter numbers separated by spaces.")
                continue

            new_indices = [int(i) - 1 for i in new_indices_str] # Convert to 0-based index

            # Validation
            num_custom = len(custom_section)
            if len(new_indices) != num_custom:
                print(f"Error: You entered {len(new_indices)} numbers, but there are {num_custom} custom mods.")
                continue
            if not all(0 <= idx < num_custom for idx in new_indices):
                print(f"Error: Numbers must be between 1 and {num_custom}.")
                continue
            if len(set(new_indices)) != num_custom:
                print("Error: Duplicate numbers entered. Each mod must appear exactly once.")
                continue

            # Build the new custom section list
            new_custom_section = [custom_section[i] for i in new_indices]

            print("\nNew Proposed Order:")
            for i, plugin in enumerate(default_section): print(f"  D{i+1}: {plugin} (Default - Unchanged)")
            for i, plugin in enumerate(new_custom_section): print(f"  {i + 1}: {plugin}")

            confirm = input("Save this new load order? (y/n): ").lower().strip()
            if confirm == 'y':
                # Combine default section (preserving its original order) and the new custom order
                final_plugin_list = default_section + new_custom_section
                if write_plugins_file(final_plugin_list):
                    print("Load order successfully updated.")
                else:
                    print("ERROR: Failed to write updated plugins.txt.")
                return # Exit load order function
            else:
                print("Load order change cancelled.")
                return # Exit load order function

        except ValueError:
            print("Invalid input. Please enter numbers only.")
        except (EOFError, KeyboardInterrupt):
            print("\nLoad order change cancelled.")
            return # Exit load order function

# --- Main Execution ---

def main():
    print("--- Oblivion Remastered Simple Mod Manager ---")
    # ... (initial print statements) ...
    print(f"ESP Data Path:   {ESP_DATA_PATH}")
    print(f"Plugins File:    {PLUGINS_TXT_PATH.name}") # Show only filename
    print(f"Pak Mods Path:   {PAK_MODS_PATH}")
    print(f"Supported Archives: {', '.join(SUPPORTED_EXTENSIONS)}")
    if not RAR_SUPPORT: print("WARNING: RAR support disabled.")
    if not SEVENZIP_SUPPORT: print("WARNING: 7z support disabled.")
    print("-" * 60)

    # Admin Check
    if not is_admin():
        print("INFO: Attempting admin relaunch...");
        if run_as_admin(): time.sleep(1); sys.exit(0)
        else: print("FATAL: Admin privileges needed."); input("Press Enter..."); sys.exit(1)

    # Batch Mode Check
    archive_files_batch = []
    is_batch_mode = False
    if len(sys.argv) > 1:
        # ... (batch mode detection logic) ...
        print("\nINFO: Checking command line arguments...")
        for arg_path in sys.argv[1:]:
            if Path(arg_path).is_file() and is_supported_archive(arg_path):
                print(f"  + Found archive: {Path(arg_path).name}")
                archive_files_batch.append(str(Path(arg_path).resolve()))
        if archive_files_batch:
            is_batch_mode = True
            print("INFO: Entering Batch Installation mode.")


    # Main Loop
    while True:
        if is_batch_mode:
            # Process batch files then exit loop
            # ... (batch processing logic) ...
            print(f"\n--- Starting Batch Installation ({len(archive_files_batch)} file(s)) ---")
            overall_success = True
            for archive_path in archive_files_batch:
                 success_file = install_mod_from_archive(archive_path)
                 if not success_file: overall_success = False
            print("\n--- Batch Processing Complete ---")
            if not overall_success: print("NOTE: One or more installations failed or had warnings.")
            break # Exit loop after batch

        else:
            # Interactive Mode
            print("\n--- Main Menu ---")
            # List mods
            print("Installed Custom ESP Mods (Order Matters):")
            esp_mods = get_installed_custom_mods();
            if esp_mods: [print(f" - {mod}") for mod in esp_mods]
            else: print("  None")
            print("\nInstalled Pak Mods (Order Ignored):")
            pak_mods = get_installed_pak_mods();
            if pak_mods: [print(f" - {mod}") for mod in pak_mods]
            else: print("  None")
            print("-" * 20)

            # Get action
            choice = input("Action: (I)nstall, Uninstall (E)SP, Uninstall (P)ak, Change (L)oad Order, (Q)uit? ").lower().strip()

            if choice == 'i':
                while True:
                    arc_path = input("Path to mod archive (or drag & drop, 'c' to cancel): ").strip().strip('"')
                    if arc_path.lower() == 'c': print("Cancelled."); break
                    if Path(arc_path).is_file() and is_supported_archive(arc_path):
                        install_mod_from_archive(str(Path(arc_path).resolve()))
                        break # Return to main menu
                    else: print("Invalid path or unsupported file.")
            elif choice == 'e':
                process_esp_uninstallation()
            elif choice == 'p':
                process_pak_uninstallation()
            elif choice == 'l':
                process_load_order() # Call the new function
            elif choice == 'q':
                print("Exiting.")
                break # Exit the main loop
            else:
                print("Invalid choice.")
            # Loop continues, re-displaying menu unless 'q' was chosen


# --- Refactored Installation Logic ---
def install_mod_from_archive(archive_path):
    """Detects content and installs ESP or Pak mod from a single archive."""
    print(f"\nProcessing: {Path(archive_path).name}")
    success_file = False
    try:
        pak_sets = find_pak_sets_in_archive(archive_path)
        esp_files = find_esps_in_archive(archive_path)
        if pak_sets: print("INFO: Pak content detected."); success_file = process_pak_installation(archive_path)
        elif esp_files: print("INFO: ESP content detected."); success_file = process_esp_installation(archive_path)
        else: print(f"ERROR: No installable content found in '{Path(archive_path).name}'."); success_file = False
    except (FileNotFoundError, ValueError, RuntimeError) as e:
        print(f"ERROR: Cannot read archive '{Path(archive_path).name}': {e}"); success_file = False
    except Exception as e:
        print(f"CRITICAL ERROR processing archive '{Path(archive_path).name}': {e}"); import traceback; traceback.print_exc(); success_file = False
    return success_file


# --- Script Entry Point ---
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n--- A CRITICAL ERROR OCCURRED IN MAIN ---"); print(f"Error: {e}"); import traceback; traceback.print_exc()
    finally:
        print("\n" + "-" * 60); input("Press Enter to exit.")