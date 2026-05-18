import sys
from pathlib import Path
import pytest

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))


@pytest.fixture
def base_dir():
    return BASE_DIR


@pytest.fixture
def test_data_dir():
    return BASE_DIR / "test_data"


@pytest.fixture
def xchecks_instance():
    from strategies.x_checks.x_checks import XChecks
    from task_configs import X_CHECKS_UPLOAD_CONFIG
    instance = XChecks(X_CHECKS_UPLOAD_CONFIG)
    instance.log = []
    return instance
