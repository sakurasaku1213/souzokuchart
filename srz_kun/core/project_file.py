import json

def save_project(project_data: dict, filepath: str) -> bool:
    """
    Saves the project data to a file in JSON format.

    Args:
        project_data (dict): A dictionary containing all project data
                             (application_version, people, diagram_layout, checklist).
        filepath (str): The path to the file where the data will be saved.

    Returns:
        bool: True if saving was successful, False otherwise.
    """
    print(f"Attempting to save project data to {filepath}...")
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(project_data, f, ensure_ascii=False, indent=4, sort_keys=True)
        print(f"Project data successfully saved to {filepath}")
        return True
    except IOError as e:
        print(f"Error saving project to {filepath}: {e}")
        # Potentially raise a custom exception here if more specific error handling is needed upstream
        return False
    except Exception as e: # Catch any other unexpected errors during JSON serialization or file writing
        print(f"An unexpected error occurred during saving to {filepath}: {e}")
        return False


def load_project(filepath: str) -> dict | None:
    """
    Loads project data from a JSON file.

    The function expects the file to contain a JSON object structured
    according to the .srk format.

    Args:
        filepath (str): The path to the file from which to load the data.

    Returns:
        dict | None: A dictionary containing the project data if successful,
                     or None if an error occurs (e.g., file not found, invalid JSON).
    """
    print(f"Attempting to load project data from {filepath}...")
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        print(f"Project data successfully loaded from {filepath}")
        return data
    except FileNotFoundError:
        print(f"Error: Project file not found at {filepath}")
        return None
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON format in {filepath}. Details: {e}")
        return None
    except IOError as e: # Catch other I/O errors during reading
        print(f"Error loading project from {filepath}: {e}")
        return None
    except Exception as e: # Catch any other unexpected errors
        print(f"An unexpected error occurred during loading from {filepath}: {e}")
        return None

# Example Usage (for testing purposes - can be removed or commented out)
if __name__ == '__main__':
    sample_data = {
        "application_version": "1.0.0",
        "people": [
            {"id": "p1", "name": "Taro Yamada", "x": 10, "y": 20},
            {"id": "p2", "name": "Hanako Yamada", "x": 150, "y": 20}
        ],
        "diagram_layout": {
            "p1": {"x": 10, "y": 20},
            "p2": {"x": 150, "y": 20}
        },
        "checklist": [{"task": "Check will", "done": True}]
    }
    test_save_path = "test_project_save.srk"

    print(f"\n--- Testing save_project ({test_save_path}) ---")
    if save_project(sample_data, test_save_path):
        print("Save successful.")

        print(f"\n--- Testing load_project ({test_save_path}) ---")
        loaded_data = load_project(test_save_path)
        if loaded_data:
            print("Load successful. Data:")
            # print(json.dumps(loaded_data, indent=2, ensure_ascii=False))
            assert loaded_data["application_version"] == "1.0.0"
            assert len(loaded_data["people"]) == 2
        else:
            print("Load failed.")

        # Clean up
        if os.path.exists(test_save_path):
            os.remove(test_save_path)
            print(f"Cleaned up {test_save_path}")
    else:
        print("Save failed.")

    print("\n--- Testing load_project (file not found) ---")
    load_project("non_existent_file.srk")

    print("\n--- Testing load_project (invalid JSON) ---")
    invalid_json_path = "invalid_json_test.srk"
    with open(invalid_json_path, "w") as f:
        f.write("{'invalid_json': True,}") # Malformed JSON
    load_project(invalid_json_path)
    if os.path.exists(invalid_json_path):
        os.remove(invalid_json_path)
        print(f"Cleaned up {invalid_json_path}")
