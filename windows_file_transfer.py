import os  
import csv  
import threading  
import shutil  
import configparser  
import getpass  
import re  
import sys  
from datetime import datetime  
import win32security  
import win32con  
  
INI_DEFAULTS = {  
    'OriginPath': '',  
    'TargetPath': '',  
    'Threads': '4',  
    'AutoTimestampDir': 'False',  
    'SanitizeDoubleExt': 'True',  
    'StripIllegalChars': 'True',  
    'MetadataFields': 'All'  
}  
  
ILLEGAL_CHARS = r'[<>:"/\\|?*]'  
  
def prompt_for_ini(ini_path):  
    if not os.path.exists(ini_path):  
        print(f"INI file '{ini_path}' not found.")  
        create = input("Create new INI file with defaults? (y/n): ").strip().lower()  
        if create == 'y':  
            config = configparser.ConfigParser()  
            config['DEFAULT'] = INI_DEFAULTS  
            with open(ini_path, 'w') as f:  
                config.write(f)  
            print(f"Created {ini_path} with defaults.")  
        else:  
            print("Exiting.")  
            sys.exit(1)  
  
def load_ini(ini_path):  
    config = configparser.ConfigParser()  
    config.read(ini_path)  
    return config['DEFAULT']  
  
def prompt_for_credentials():  
    print("Enter Windows credentials (leave blank to use current user):")  
    username = input("Username (DOMAIN\\user): ").strip()  
    if username:  
        password = getpass.getpass("Password: ")  
        return username, password  
    return None, None  
  
def impersonate_user(username, password):  
    # Only if credentials provided  
    if not username or not password:  
        return None  
    try:  
        domain, user = username.split('\\')  
        handle = win32security.LogonUser(  
            user, domain, password,  
            win32con.LOGON32_LOGON_INTERACTIVE,  
            win32con.LOGON32_PROVIDER_DEFAULT  
        )  
        win32security.ImpersonateLoggedOnUser(handle)  
        return handle  
    except Exception as e:  
        print(f"Impersonation failed: {e}")  
        return None  
  
def sanitize_filename(filename, strip_illegal=True, sanitize_ext=True):  
    orig = filename  
    if strip_illegal:  
        filename = re.sub(ILLEGAL_CHARS, '_', filename)  
    if sanitize_ext:  
        parts = filename.split('.')  
        if len(parts) > 2:  
            filename = parts[0] + '.' + parts[-1]  
    return filename, orig != filename  
  
def make_timestamped_dir(base_path):  
    now = datetime.now().strftime('%Y%m%d_%H%M%S')  
    new_dir = os.path.join(base_path, now)  
    os.makedirs(new_dir, exist_ok=True)  
    return new_dir  
  
def is_relative(path):  
    return not os.path.isabs(path)  
  
def make_hyperlink(path):  
    # For Excel, hyperlinks are like: =HYPERLINK("file:///C:/path/to/file")  
    return f'=HYPERLINK("file:///{path.replace("\\", "/")}")'  
  
def process_transfer(row, idx, config, meta_fields, origin_base, target_base, options, meta_writer, lock):  
    origin_path = row[0]  
    target_path = row[1]  
    illegal_chars_stripped = False  
  
    # Handle relative paths  
    if is_relative(origin_path):  
        origin_path = os.path.join(origin_base, origin_path)  
    if is_relative(target_path):  
        target_path = os.path.join(target_base, target_path)  
  
    # Sanitize filename if needed  
    filename = os.path.basename(target_path)  
    sanitized, was_sanitized = sanitize_filename(  
        filename,  
        strip_illegal=options['strip_illegal'],  
        sanitize_ext=options['sanitize_ext']  
    )  
    illegal_chars_stripped = was_sanitized  
  
    target_dir = os.path.dirname(target_path)  
    if not os.path.exists(target_dir):  
        os.makedirs(target_dir, exist_ok=True)  
    target_path = os.path.join(target_dir, sanitized)  
  
    # Copy file  
    status = 'Success'  
    error = ''  
    try:  
        shutil.copy2(origin_path, target_path)  
    except Exception as e:  
        status = 'Failure'  
        error = str(e)  
  
    # Prepare metadata row  
    meta_row = []  
    if meta_fields == 'All':  
        meta_row = row  
    else:  
        meta_row = [row[int(i)-1] for i in meta_fields.split(',') if i.isdigit() and int(i)-1 < len(row)]  
  
    meta_row += [  
        make_hyperlink(origin_path),  
        make_hyperlink(target_path),  
        f"{status}: {error}" if error else status,  
        str(illegal_chars_stripped)  
    ]  
  
    # Write metadata (thread-safe)  
    with lock:  
        meta_writer.writerow(meta_row)  
  
def main():  
    ini_path = 'config.ini'  
    prompt_for_ini(ini_path)  
    config = load_ini(ini_path)  
  
    # Prompt for CSV file  
    csv_path = input("Enter FileTransfers CSV path: ").strip()  
    if not os.path.exists(csv_path):  
        print(f"CSV file '{csv_path}' not found.")  
        sys.exit(1)  
  
    # Origin/target base paths  
    origin_base = config.get('OriginPath', '')  
    target_base = config.get('TargetPath', '')  
  
    # Prompt for credentials if not in ini  
    username = config.get('Username', '')  
    password = config.get('Password', '')  
    if not username:  
        username, password = prompt_for_credentials()  
  
    # Impersonate if credentials provided  
    handle = None  
    if username and password:  
        handle = impersonate_user(username, password)  
        if not handle:  
            print("Proceeding with current user credentials.")  
  
    # Threading  
    threads_count = int(config.get('Threads', '4'))  
  
    # Options  
    options = {  
        'timestamp_dir': config.get('AutoTimestampDir', 'False').lower() == 'true',  
        'sanitize_ext': config.get('SanitizeDoubleExt', 'True').lower() == 'true',  
        'strip_illegal': config.get('StripIllegalChars', 'True').lower() == 'true'  
    }  
  
    # Metadata fields  
    meta_fields = config.get('MetadataFields', 'All')  
  
    # Optionally timestamp target dir  
    if options['timestamp_dir']:  
        target_base = make_timestamped_dir(target_base)  
  
    # Prepare metadata file  
    meta_file = os.path.splitext(csv_path)[0] + '_metadata.csv'  
    lock = threading.Lock()  
  
    with open(csv_path, newline='', encoding='utf-8') as f_in, \  
         open(meta_file, 'w', newline='', encoding='utf-8') as f_meta:  
        reader = csv.reader(f_in)  
        meta_writer = csv.writer(f_meta)  
  
        # Write header  
        header = next(reader)  
        if meta_fields == 'All':  
            meta_header = header  
        else:  
            meta_header = [header[int(i)-1] for i in meta_fields.split(',') if i.isdigit() and int(i)-1 < len(header)]  
        meta_header += ['Origin Hyperlink', 'Target Hyperlink', 'Transfer Status', 'Illegal_Chars_Stripped']  
        meta_writer.writerow(meta_header)  
  
        # Prepare jobs  
        jobs = []  
        for idx, row in enumerate(reader):  
            t = threading.Thread(  
                target=process_transfer,  
                args=(row, idx, config, meta_fields, origin_base, target_base, options, meta_writer, lock)  
            )  
            jobs.append(t)  
  
        # Start threads in batches  
        for i in range(0, len(jobs), threads_count):  
            batch = jobs[i:i+threads_count]  
            for t in batch:  
                t.start()  
            for t in batch:  
                t.join()  
  
    # Revert impersonation  
    if handle:  
        win32security.RevertToSelf()  
        handle.Close()  
  
    print(f"All transfers complete. Metadata written to {meta_file}")  
  
if __name__ == '__main__':  
    main()  
