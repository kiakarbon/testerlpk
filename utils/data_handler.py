import json
import os
import tempfile
from datetime import datetime

DATA_FILES = {
    'catatan': 'catatan_praktik.json',
    'psa_results': 'hasil_psa.json'
}

def get_data_path(filename):
    """Get path for data file"""
    temp_dir = tempfile.gettempdir()
    return os.path.join(temp_dir, filename)

def save_data(data_type, data):
    """Save data to JSON file"""
    if data_type in DATA_FILES:
        filepath = get_data_path(DATA_FILES[data_type])
        try:
            # Convert data to serializable format
            serializable_data = []
            for item in data:
                serializable_item = {}
                for key, value in item.items():
                    if key == 'image_paths':
                        # Don't save image paths in JSON
                        continue
                    elif isinstance(value, datetime):
                        serializable_item[key] = value.isoformat()
                    else:
                        serializable_item[key] = value
                serializable_data.append(serializable_item)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(serializable_data, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error saving data: {e}")
            return False
    return False

def load_data(data_type):
    """Load data from JSON file"""
    if data_type in DATA_FILES:
        filepath = get_data_path(DATA_FILES[data_type])
        if os.path.exists(filepath):
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return data
            except Exception as e:
                print(f"Error loading data: {e}")
                return []
        else:
            # Return empty list if file doesn't exist
            return []
    return []

def clear_data():
    """Clear all data files"""
    for filename in DATA_FILES.values():
        filepath = get_data_path(filename)
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except:
                pass
    return True

def backup_data():
    """Create backup of data files"""
    backup_dir = os.path.join(tempfile.gettempdir(), 'backups')
    os.makedirs(backup_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    for filename in DATA_FILES.values():
        source = get_data_path(filename)
        if os.path.exists(source):
            backup_file = os.path.join(backup_dir, f"{filename}_{timestamp}.json")
            try:
                with open(source, 'r', encoding='utf-8') as f_src:
                    with open(backup_file, 'w', encoding='utf-8') as f_dst:
                        f_dst.write(f_src.read())
            except Exception as e:
                print(f"Error creating backup: {e}")
    
    return backup_dir
