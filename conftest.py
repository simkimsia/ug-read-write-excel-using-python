import pytest


def pytest_addoption(parser):
    parser.addoption("--skip-updown", default=False)


@pytest.fixture
def skip_updown(request):
    return request.config.getoption("--skip-updown")
