import pytest


def pytest_addoption(parser):
    parser.addoption("--skip-updown", action="store_true")


@pytest.fixture
def skip_updown(request):
    return request.config.getoption("--skip-updown")
