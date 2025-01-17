import subprocess
import sys

def install_requirements(requirements_file='requirements.txt'):
    """
    Install libraries listed in a requirements.txt file.
    """
    try:
        with open(requirements_file, 'r') as file:
            libraries = file.readlines()
            libraries = [lib.strip() for lib in libraries if lib.strip() and not lib.startswith('#')]

        print(f"Found {len(libraries)} libraries to install.")
        for lib in libraries:
            print(f"Installing {lib}...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', lib])
            print(f"{lib} installed successfully.")
    
    except FileNotFoundError:
        print(f"Error: {requirements_file} not found. Please ensure the file exists.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install a library: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    install_requirements()
