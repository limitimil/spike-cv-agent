"""Cross-platform entry point: python resume.py [command]."""
import sys
from pathlib import Path

if __name__ == '__main__':
    from src.__main__ import main
    # Root is explicit and stable even when launched from another directory.
    if '--root' not in sys.argv:
        sys.argv[1:1] = ['--root', str(Path(__file__).resolve().parent)]
    if len(sys.argv) == 3:
        sys.argv.append('start')
    main()
