"""Allow running TsaraKonta as a package.

Usage:
    python -m finances.tsarakonta [--excel path] [--etat EtatName]
"""

from .main import main

if __name__ == "__main__":
    main()
