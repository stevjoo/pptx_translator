import os
import sys

# Ensure parent directory is in sys.path so pptx_translator package resolves
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from pptx_translator.cli import main

if __name__ == "__main__":
    main()
