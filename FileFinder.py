def get_first_line(file_path):
    try:
        with open(file_path, 'r') as file:
            first_line = file.readline().strip()
            return first_line
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None


