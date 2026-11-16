"""Real SDK discovery checks; no Excel or paid model messages."""

from pathlib import Path

import pytest

from conftest import build_excel_cli_eval, build_excel_mcp_eval
from skill_discovery import assert_skill_exposure, discover_skills


@pytest.mark.parametrize("transport", ("mcp", "cli"))
@pytest.mark.parametrize("with_skill", (False, True))
async def test_skill_availability_without_model_messages(
    github_auth, tmp_path: Path, request: pytest.FixtureRequest, transport, with_skill,
):
    builder = build_excel_mcp_eval if transport == "mcp" else build_excel_cli_eval
    if with_skill and request.config.getoption(f"--{transport}-skill-dir") is None:
        pytest.skip(f"Historical discovery requires --{transport}-skill-dir.")
    supplied = request.getfixturevalue(f"excel_{transport}_skill_dir") if with_skill else None
    agent = builder("discovery-only", servers={}, working_directory=str(tmp_path),
                    skill_dir=supplied)
    discovered = await discover_skills(agent)
    assert_skill_exposure(discovered, f"excel-{transport}-report-formatting" if with_skill else None,
                          supplied)
