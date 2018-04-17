import pytest


def pytest_addoption(parser):
    parser.addoption("--skip-updown", action="store_true")
    parser.addoption("--skip-up", action="store_true")
    parser.addoption("--skip-down", action="store_true")


@pytest.fixture
def skip_updown(request):
    return request.config.getoption("--skip-updown")


@pytest.fixture
def skip_up(request):
    return request.config.getoption("--skip-up")


@pytest.fixture
def skip_down(request):
    return request.config.getoption("--skip-down")
