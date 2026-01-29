import unittest
import sys
from pathlib import Path


def run():
    # Ensure project root is on sys.path so tests can import `src` package
    project_root = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(project_root))

    loader = unittest.TestLoader()
    start_dir = project_root / "tests"
    suite = loader.discover(start_dir=str(start_dir))
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    if not result.wasSuccessful():
        sys.exit(1)


if __name__ == "__main__":
    run()
