from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from harbor.agents.installed.base import BaseInstalledAgent, CliFlag, with_prompt_template
from harbor.environments.base import BaseEnvironment
from harbor.models.agent.context import AgentContext
from harbor.models.trial.paths import EnvironmentPaths

TASK_SUBMISSION_PROMPT_TEMPLATE = """You are tasked with completing a specific assignment.

## Environment

The `run_shell` tool provides access to a Linux-based execution environment that includes a full file system where you can create, read, and modify files.

Your environment comes preinstalled with a comprehensive set of Python packages and system tools:

**Jupyter Ecosystem:**
- jupyter-client 8.6.1, jupyter-core 5.5.1, jupyter-server 2.14.0
- jupyterlab 4.1.8, jupyterlab-pygments 0.3.0, jupyterlab-server 2.27.1
- notebook 6.5.1, nbclassic 0.4.5

**Web Frameworks:**
- aiohttp 3.9.5, hypercorn 0.14.3, fastapi 0.95.2, websockets 10.3
- pydantic 1.10.2, gradio 2.2.15

**Core Data Science:**
- numpy 1.24.0, numpy-financial 1.0.0, scipy 1.14.1, pandas 1.5.3
- matplotlib 3.6.3, matplotlib-venn 0.11.6, seaborn 0.11.2
- plotly 5.3.0, plotnine 0.10.1, bokeh 2.4.0

**Statistics & Machine Learning:**
- statsmodels 0.13.5, scikit-learn 1.1.3, scikit-image 0.20.0
- xgboost 1.4.2, catboost ~1.2.7, lightgbm ~4.5.0
- imbalanced-learn ~0.12.3, shap 0.39.0

**NLP:**
- nltk 3.9.1, gensim 4.3.1, spacy 3.4.4, textblob 0.15.3

**Computer Vision:**
- opencv-python 4.5.5.62, Pillow 9.1.0
- pytesseract 0.3.8, qrcode 7.3, pyzbar 0.1.8, imgkit 1.2.2

**Audio Processing:**
- ffmpeg-python 0.2.0, pydub 0.25.1, moviepy 1.0.3, soundfile 0.10.2
- librosa 0.8.1, mutagen 1.45.1, gtts 2.2.3, pyttsx3 2.90
- pedalboard 0.9.9, pyloudnorm 0.1.1, mne 0.23.4

**Document Processing:**
- python-docx 0.8.11, python-pptx 0.6.21, openpyxl 3.0.10, xlrd 2.0.1
- PyMuPDF 1.21.1, pdf2image 1.16.3, pdfplumber 0.6.2, pdfkit 0.6.1
- pypandoc 1.6.3, docx2txt 0.8, odfpy 1.4.1, pyxlsb 1.0.8
- tabula 1.0.5, camelot-py 0.10.1

**PDF Generation:**
- fpdf2 2.8.3, reportlab 3.6.12, weasyprint 53.3, pdfrw 0.4

**Graphics & Visualization:**
- graphviz 0.17, pydot 1.4.2, networkx 2.8.8
- svglib 1.1.0, svgwrite 1.4.1, cairosvg 2.5.2, trimesh 3.9.29
- wordcloud 1.9.2, folium 0.12.1

**Geospatial:**
- shapely 2.0.6, fiona 1.9.2, geopandas 0.10.2
- geopy 2.2.0, rasterio 1.3.3, basemap 1.3.9
- GDAL system libraries

**Scientific Computing:**
- sympy 1.13.1, pymc 4.0.1, h5py 3.8.0, tables 3.8.0

**3D & CAD:**
- cadquery 2.4.0, cadquery-ocp 7.7.0

**Chemistry & Biology:**
- rdkit 2024.9.6, biopython 1.84

**Data Utilities:**
- xml-python 0.4.3, markdownify 0.9.3, anytree 2.8.0
- rarfile 4.0, chardet 3.0.4, srt 3.5.3

**General Utilities:**
- tqdm 4.64.0, tabulate 0.9.0, faker 8.13.2, loguru 0.5.3
- fuzzywuzzy 0.18.0, rapidfuzz ~3.10.1, einops 0.3.2
- pycountry 20.7.3, countryinfo 0.1.2, pronouncing 0.2.0
- kerykeion 2.1.16, exchange_calendars 3.4

**Math & Logic:**
- pylog 1.1, pyprover 0.5.6, nashpy 0.0.35

**Semantic Web:**
- rdflib 6.0.0

**Security & Networking:**
- cryptography 3.4.8, pyopenssl 21.0.0, requests 2.31.0

**Database Connectors:**
- snowflake-connector-python 2.7.12, databricks-sql-connector 0.9.1

**Testing & Monitoring:**
- pytest 8.2.0, pytest-cov 5.0.0, pytest-json-report 1.5.0
- coverage 7.5.1, pytest-asyncio 0.23.6
- ddtrace 2.8.1, datadog 0.49.1

**Document Generation:**
- aspose-words 25.8.0

**Other:**
- typing-extensions 4.10.0, pyth3 0.7

**System Tools:**
- Python 3.12 (base environment)
- LibreOffice + LibreOffice Writer (for office document conversion, includes fonts-dejavu-core)
- Tesseract OCR (text extraction from images)
- Pandoc (universal document converter)
- Poppler utilities (PDF tools such as `pdftotext`, `pdfimages`)
- Ghostscript (PostScript/PDF processing)
- FFmpeg (complete audio/video processing suite with all codecs)
- Graphviz (graph visualization with DOT language)
- OpenJDK 21 JRE (Java runtime for Tabula and other Java-based tools)
- GDAL/GEOS/Proj (geospatial data libraries and utilities)
- Build tools: gcc, g++, cmake, pkg-config, make

## Reference Files Location

The reference files for the task are available in your environment's file system.

Here are their paths:

<reference_files>
{reference_files}
</reference_files>

## Completing Your Work

In order to complete the task you must use the `{finish_tool_name}` tool to submit your work.  If you do not use the `{finish_tool_name}` tool you will fail this task!

**Required in your finish call:**
1. A brief summary of what you accomplished
2. A list of **ABSOLUTE file paths** (starting with `/home/user/`) for all files you want to submit.

## Task

Here is the task you need to complete:

<task>
{task}
</task>

Please begin working on the task now."""

MESSAGE_SUMMARIZER_PROMPT = """The context window is approaching its limit. Please create a concise summary of the conversation so far to preserve important information.

Your summary should include:

1. **Task Overview**: What is the main goal or objective?

2. **Progress Made**: What has been accomplished so far?
   - Key files created/modified (with paths)
   - Important functions/classes implemented
   - Tools used and their outcomes

3. **Current State**: Where are we now?
   - What is currently working?
   - What has been tested/verified?

4. **Next Steps**: What still needs to be done?
   - Outstanding TODOs (with specific file paths and line numbers if applicable)
   - Known issues or bugs to address
   - Features or functionality not yet implemented

5. **Important Context**: Any critical details that shouldn't be lost
   - Special configurations or setup requirements
   - Important variable names, API endpoints, or data structures
   - Edge cases or constraints to keep in mind
   - Dependencies or relationships between components

Keep the summary concise but comprehensive. Focus on actionable information that will allow smooth continuation of the work."""

MESSAGE_SUMMARIZER_BRIDGE_PROMPT = """**Context Continuation**

Due to context window limitations, the previous conversation has been summarized. Below is a summary of what happened before:

---
{summary}
---

You should continue working on this task from where it was left off. All the progress, current state, and next steps are described in the summary above. Proceed with completing any outstanding work."""

STIRRUP_RUNTIME_SCRIPT = r"""
import asyncio
import json
import os
import sys
import uuid
from pathlib import Path


def fail(message: str) -> None:
    print(message, file=sys.stderr)
    raise SystemExit(2)


def _stringify_content(content):
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for block in content:
            if isinstance(block, str):
                parts.append(block)
            elif isinstance(block, dict):
                text = block.get("text")
                if isinstance(text, str):
                    parts.append(text)
                else:
                    parts.append(json.dumps(block, ensure_ascii=False, default=str))
            else:
                parts.append(str(block))
        return "\n".join(parts)
    return str(content)


def _flatten_history(history):
    out = []
    for group in history:
        if isinstance(group, list):
            out.extend(group)
    return out


def _build_steps(messages, model_name: str):
    steps = []
    step_id = 1
    for message in messages:
        role = getattr(message, "role", "")
        if role == "assistant":
            source = "agent"
        elif role == "user":
            source = "user"
        else:
            source = "system"

        step = {
            "step_id": step_id,
            "source": source,
            "message": _stringify_content(getattr(message, "content", "")),
        }
        if role == "assistant":
            step["model_name"] = model_name
        steps.append(step)
        step_id += 1
    return steps


def _collect_token_usage(messages):
    input_tokens = 0
    output_tokens = 0
    reasoning_tokens = 0
    for message in messages:
        if getattr(message, "role", "") != "assistant":
            continue
        usage = getattr(message, "token_usage", None)
        if usage is None:
            continue
        input_tokens += int(getattr(usage, "input", 0) or 0)
        output_tokens += int(getattr(usage, "output", 0) or 0)
        reasoning_tokens += int(getattr(usage, "reasoning", 0) or 0)
    return {
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "reasoning_tokens": reasoning_tokens,
    }


async def _run() -> int:
    if sys.version_info < (3, 12):
        fail(
            "Stirrup GDPval-AA harness requires Python >= 3.12 in the task environment. "
            "Update the task Dockerfile base image to Python 3.12 or newer."
        )

    try:
        from stirrup import Agent
        from stirrup.clients.litellm_client import LiteLLMClient
        from stirrup.core.models import ImageContentBlock, Tool
        import stirrup.core.agent as stirrup_agent_module
        from stirrup.tools.code_backends.base import (
            SHELL_TIMEOUT,
            CodeExecToolProvider,
            CommandResult,
        )
        from stirrup.tools.finish import SIMPLE_FINISH_TOOL
        from stirrup.tools.view_image import ViewImageToolProvider
        from stirrup.tools.web import WebToolProvider
    except Exception as exc:
        fail(f"Unable to import Stirrup runtime dependencies: {exc}")

    class HarborContainerRunShellProvider(CodeExecToolProvider):
        def __init__(self) -> None:
            super().__init__(allowed_commands=None)
            self._working_dir = Path("/workspace")

        async def __aenter__(self):
            self._working_dir.mkdir(parents=True, exist_ok=True)
            home_user = Path("/home/user")
            if not home_user.exists():
                try:
                    home_user.parent.mkdir(parents=True, exist_ok=True)
                    home_user.symlink_to(self._working_dir, target_is_directory=True)
                except Exception:
                    pass
            return self.get_code_exec_tool(
                name="run_shell",
                description=(
                    "Executes bash commands in the Harbor task container; returns exit code, stdout, and stderr."
                ),
            )

        async def __aexit__(
            self,
            exc_type: type[BaseException] | None,
            exc_val: BaseException | None,
            exc_tb: object,
        ) -> None:
            return None

        def _resolve_path(self, path: str) -> Path:
            raw = Path(path)
            if not raw.is_absolute():
                return (self._working_dir / raw).resolve()
            if raw == Path("/home/user"):
                return self._working_dir
            try:
                relative = raw.relative_to(Path("/home/user"))
                return (self._working_dir / relative).resolve()
            except ValueError:
                return raw

        async def run_command(self, cmd: str, *, timeout: int = SHELL_TIMEOUT) -> CommandResult:
            if not self._check_allowed(cmd):
                return CommandResult(
                    exit_code=1,
                    stdout="",
                    stderr=f"Command not allowed: '{cmd}' does not match any allowed patterns",
                    error_kind="command_not_allowed",
                    advice="Only commands matching the allowlist patterns are permitted.",
                )

            process = None
            try:
                process = await asyncio.create_subprocess_exec(
                    "bash",
                    "-lc",
                    cmd,
                    cwd=self._working_dir.as_posix(),
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE,
                )
                stdout_bytes, stderr_bytes = await asyncio.wait_for(
                    process.communicate(),
                    timeout=timeout,
                )
                return CommandResult(
                    exit_code=process.returncode or 0,
                    stdout=stdout_bytes.decode("utf-8", errors="replace"),
                    stderr=stderr_bytes.decode("utf-8", errors="replace"),
                )
            except asyncio.TimeoutError:
                if process is not None:
                    process.kill()
                    await process.communicate()
                return CommandResult(
                    exit_code=1,
                    stdout="",
                    stderr=f"Command timed out after {timeout} seconds",
                    error_kind="timeout",
                )
            except Exception as exc:
                return CommandResult(
                    exit_code=1,
                    stdout="",
                    stderr=str(exc),
                    error_kind="execution_error",
                )

        async def read_file_bytes(self, path: str) -> bytes:
            resolved = self._resolve_path(path)
            if not resolved.exists():
                raise FileNotFoundError(f"File not found: {path}")
            return resolved.read_bytes()

        async def write_file_bytes(self, path: str, content: bytes) -> None:
            resolved = self._resolve_path(path)
            resolved.parent.mkdir(parents=True, exist_ok=True)
            resolved.write_bytes(content)

        async def file_exists(self, path: str) -> bool:
            resolved = self._resolve_path(path)
            return resolved.exists() and resolved.is_file()

        async def is_directory(self, path: str) -> bool:
            resolved = self._resolve_path(path)
            return resolved.exists() and resolved.is_dir()

        async def list_files(self, path: str) -> list[str]:
            resolved = self._resolve_path(path)
            if not resolved.exists() or not resolved.is_dir():
                return []
            return [str(child.relative_to(resolved)) for child in resolved.rglob("*") if child.is_file()]

        async def view_image(self, path: str) -> ImageContentBlock:
            return ImageContentBlock(data=await self.read_file_bytes(path))

    class GDPvalAAWebToolProvider(WebToolProvider):
        def get_tools(self):
            tools = super().get_tools()
            aliased = []
            for tool in tools:
                if tool.name == "fetch_web_page":
                    aliased.append(
                        Tool(
                            name="web_fetch",
                            description=(
                                "Fetches and extracts main content from a web page as markdown."
                            ),
                            parameters=tool.parameters,
                            executor=tool.executor,
                        )
                    )
                else:
                    aliased.append(tool)
            return aliased

    _litellm_timeout = int(os.environ.get("HARBOR_STIRRUP_LITELLM_TIMEOUT", "3600"))
    import litellm
    litellm.request_timeout = _litellm_timeout
    os.environ["LITELLM_REQUEST_TIMEOUT"] = str(_litellm_timeout)

    model_name = os.environ["HARBOR_STIRRUP_MODEL_NAME"]
    reasoning_effort = os.environ.get("HARBOR_STIRRUP_REASONING_EFFORT", "high")
    vision_enabled = os.environ.get("HARBOR_STIRRUP_VISION_ENABLED", "true").lower() in {
        "1",
        "true",
        "yes",
    }

    max_turns = int(os.environ.get("HARBOR_STIRRUP_MAX_TURNS", "100"))
    context_cutoff = float(os.environ.get("HARBOR_STIRRUP_CONTEXT_CUTOFF", "0.7"))
    warning_threshold = int(os.environ.get("HARBOR_STIRRUP_WARNING_THRESHOLD", "20"))

    instruction = os.environ.get("HARBOR_STIRRUP_INSTRUCTION", "").strip()
    if not instruction:
        fail("Received empty Harbor task instruction for Stirrup harness.")

    log_dir = Path(os.environ["HARBOR_STIRRUP_LOG_DIR"])
    log_dir.mkdir(parents=True, exist_ok=True)
    submission_dir = log_dir / "submission"
    submission_dir.mkdir(parents=True, exist_ok=True)
    (log_dir / ".harbor_submission_dir").write_text("submission\n")

    stirrup_agent_module.MESSAGE_SUMMARIZER = os.environ[
        "HARBOR_STIRRUP_SUMMARIZER_PROMPT"
    ]
    stirrup_agent_module.MESSAGE_SUMMARIZER_BRIDGE_TEMPLATE = os.environ[
        "HARBOR_STIRRUP_SUMMARIZER_BRIDGE_PROMPT"
    ]

    client = LiteLLMClient(
        model=model_name,
        reasoning_effort=reasoning_effort,
        kwargs={"timeout": _litellm_timeout},
    )

    tools = [
        HarborContainerRunShellProvider(),
        GDPvalAAWebToolProvider(),
    ]
    if vision_enabled:
        tools.append(ViewImageToolProvider())

    agent = Agent(
        client=client,
        name="stirrup",
        tools=tools,
        finish_tool=SIMPLE_FINISH_TOOL,
        max_turns=max_turns,
        context_summarization_cutoff=context_cutoff,
        turns_remaining_warning_threshold=warning_threshold,
    )

    input_root = Path("/workspace/inputs")
    input_files = [input_root] if input_root.exists() else None

    async with agent.session(output_dir=submission_dir, input_files=input_files) as session:
        from stirrup.core.agent import _SESSION_STATE

        state = _SESSION_STATE.get(None)
        reference_files = []
        if state is not None and isinstance(state.uploaded_file_paths, list):
            reference_files = [str(path) for path in state.uploaded_file_paths]

        reference_files_text = "\n".join(reference_files) if reference_files else "(none)"
        prompt_template = os.environ["HARBOR_STIRRUP_TASK_PROMPT_TEMPLATE"]
        prompt = prompt_template.format(
            reference_files=reference_files_text,
            finish_tool_name="finish",
            task=instruction,
        )
        (log_dir / "task-prompt.txt").write_text(prompt + "\n", encoding="utf-8")

        finish_params, history, run_metadata = await session.run(prompt)

    messages = _flatten_history(history)
    steps = _build_steps(messages, model_name=model_name)
    usage = _collect_token_usage(messages)

    summary_payload = {
        "finish_params": finish_params.model_dump() if finish_params is not None else None,
        "token_usage": usage,
        "model_name": model_name,
        "reasoning_effort": reasoning_effort,
        "reference_files": reference_files,
        "run_metadata_keys": sorted(run_metadata.keys()) if isinstance(run_metadata, dict) else [],
    }
    (log_dir / "stirrup-run.json").write_text(
        json.dumps(summary_payload, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    trajectory_payload = {
        "schema_version": "ATIF-v1.5",
        "session_id": str(uuid.uuid4()),
        "agent": {
            "name": "stirrup",
            "version": "0.1.8",
            "model_name": model_name,
            "extra": {
                "reasoning_effort": reasoning_effort,
                "vision_enabled": vision_enabled,
                "run_shell_backend": "harbor_container",
            },
        },
        "steps": steps,
        "final_metrics": {
            "total_prompt_tokens": usage["input_tokens"],
            "total_completion_tokens": usage["output_tokens"],
            "total_cached_tokens": None,
            "total_cost_usd": None,
            "total_steps": len(steps),
            "extra": {
                "reasoning_tokens": usage["reasoning_tokens"],
            },
        },
    }
    (log_dir / "trajectory.json").write_text(
        json.dumps(trajectory_payload, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    if finish_params is None:
        fail(
            "Stirrup run ended without calling finish; increase max_turns or inspect /logs/agent/stirrup.txt"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(_run()))
"""


class HarborGDPvalAAStirrup(BaseInstalledAgent):
    SUPPORTS_ATIF: bool = True
    GDPVAL_MAX_TURNS: int = 500
    GDPVAL_CONTEXT_SUMMARIZATION_CUTOFF: str = "0.7"
    GDPVAL_WARNING_THRESHOLD: int = 20

    CLI_FLAGS = [
        CliFlag(
            "reasoning_effort",
            cli="--reasoning-effort",
            type="enum",
            choices=["none", "minimal", "low", "medium", "high", "xhigh", "default"],
            default="high",
        ),
        CliFlag(
            "vision_enabled",
            cli="--vision-enabled",
            type="bool",
            default=True,
        ),
        CliFlag(
            "litellm_timeout",
            cli="--litellm-timeout",
            type="int",
            default=3600,
        ),
    ]

    @staticmethod
    def name() -> str:
        return "stirrup"

    async def install(self, environment: BaseEnvironment) -> None:
        await self.exec_as_agent(
            environment,
            command=(
                "python3 - <<'PY'\n"
                "import stirrup\n"
                "print(getattr(stirrup, '__version__', 'unknown'))\n"
                "PY"
            ),
        )

    def get_version_command(self) -> str | None:
        return (
            "python3 - <<'PY'\n"
            "import stirrup\n"
            "print(getattr(stirrup, '__version__', 'unknown'))\n"
            "PY"
        )

    def parse_version(self, stdout: str) -> str:
        text = stdout.strip()
        for line in text.splitlines():
            candidate = line.strip()
            if candidate:
                return candidate
        return text

    def _resolved_value(self, key: str, default: Any) -> Any:
        if key in self._resolved_flags:
            return self._resolved_flags[key]
        return default

    @with_prompt_template
    async def run(
        self,
        instruction: str,
        environment: BaseEnvironment,
        context: AgentContext,
    ) -> None:
        if not self.model_name:
            raise ValueError("Stirrup harness requires model_name in provider/model format.")

        env: dict[str, str] = {
            "HARBOR_STIRRUP_MODEL_NAME": self.model_name,
            "HARBOR_STIRRUP_REASONING_EFFORT": str(
                self._resolved_value("reasoning_effort", "high")
            ),
            "HARBOR_STIRRUP_VISION_ENABLED": (
                "true"
                if bool(self._resolved_value("vision_enabled", True))
                else "false"
            ),
            "HARBOR_STIRRUP_MAX_TURNS": str(self.GDPVAL_MAX_TURNS),
            "HARBOR_STIRRUP_CONTEXT_CUTOFF": self.GDPVAL_CONTEXT_SUMMARIZATION_CUTOFF,
            "HARBOR_STIRRUP_WARNING_THRESHOLD": str(self.GDPVAL_WARNING_THRESHOLD),
            "HARBOR_STIRRUP_INSTRUCTION": instruction,
            "HARBOR_STIRRUP_LOG_DIR": EnvironmentPaths.agent_dir.as_posix(),
            "HARBOR_STIRRUP_TASK_PROMPT_TEMPLATE": TASK_SUBMISSION_PROMPT_TEMPLATE,
            "HARBOR_STIRRUP_SUMMARIZER_PROMPT": MESSAGE_SUMMARIZER_PROMPT,
            "HARBOR_STIRRUP_SUMMARIZER_BRIDGE_PROMPT": MESSAGE_SUMMARIZER_BRIDGE_PROMPT,
            "HARBOR_STIRRUP_LITELLM_TIMEOUT": str(
                self._resolved_value("litellm_timeout", 1800)
            ),
            "PYTHONUNBUFFERED": "1",
        }

        passthrough_env_keys = {
            "OPENAI_API_KEY",
            "OPENAI_BASE_URL",
            "OPENAI_API_BASE",
            "ANTHROPIC_API_KEY",
            "GOOGLE_API_KEY",
            "GEMINI_API_KEY",
            "GOOGLE_GENERATIVE_AI_API_KEY",
            "GOOGLE_APPLICATION_CREDENTIALS",
            "GOOGLE_CLOUD_PROJECT",
            "GOOGLE_CLOUD_LOCATION",
            "GOOGLE_GENAI_USE_VERTEXAI",
            "XAI_API_KEY",
            "OPENROUTER_API_KEY",
            "BRAVE_API_KEY",
            "LITELLM_API_KEY",
        }
        for key in passthrough_env_keys:
            value = os.environ.get(key)
            if value:
                env[key] = value

        stirrup_log_path = (EnvironmentPaths.agent_dir / "stirrup.txt").as_posix()
        run_command = "\n".join(
            [
                "set -euo pipefail",
                "{",
                "python3 - <<'PY'",
                STIRRUP_RUNTIME_SCRIPT.strip(),
                "PY",
                f"}} 2>&1 </dev/null | stdbuf -oL tee {stirrup_log_path}",
            ]
        )

        await self.exec_as_agent(
            environment,
            command=run_command,
            env=env,
        )

    def populate_context_post_run(self, context: AgentContext) -> None:
        summary_path = self.logs_dir / "stirrup-run.json"
        if not summary_path.exists():
            return

        try:
            summary_payload = json.loads(summary_path.read_text(encoding="utf-8"))
        except Exception as exc:
            print(f"Failed to parse Stirrup summary payload: {exc}")
            return

        if isinstance(summary_payload, dict):
            usage = summary_payload.get("token_usage")
            if isinstance(usage, dict):
                context.n_input_tokens = int(usage.get("input_tokens") or 0)
                context.n_output_tokens = int(usage.get("output_tokens") or 0)
                context.n_cache_tokens = int(usage.get("cached_tokens") or 0)
            cost_usd = summary_payload.get("cost_usd")
            if isinstance(cost_usd, (int, float)):
                context.cost_usd = float(cost_usd)

        trajectory_path = self.logs_dir / "trajectory.json"
        if not trajectory_path.exists():
            fallback_payload = {
                "schema_version": "ATIF-v1.5",
                "session_id": "stirrup-missing-trajectory",
                "agent": {
                    "name": "stirrup",
                    "version": "unknown",
                    "model_name": self.model_name,
                },
                "steps": [],
            }
            trajectory_path.write_text(
                json.dumps(fallback_payload, indent=2, ensure_ascii=False) + "\n",
                encoding="utf-8",
            )
