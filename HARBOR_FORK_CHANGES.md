# Harbor Fork — Required Changes

All changes are relative to the Harbor package source code. These patches enable:
- E2B environment to work with COPY directives in Dockerfiles
- E2B free-tier sandbox timeout compatibility
- Runtime upload of task data files and skills (no template rebuild needed)

---

## File 1: `harbor/environments/e2b.py`

### Change 1A: Fix `file_context_path` for Dockerfile COPY resolution

**Problem**: E2B's `Template()` resolves COPY paths relative to the caller's Python module directory (Harbor's `environments/` package dir), not the task's `environment/` directory. This causes `COPY data/ /workspace/data/` to fail with "No files found".

**Location**: `_create_template` method (~line 98)

**Before:**
```python
async def _create_template(self):
    if self.task_env_config.docker_image:
        template = Template().from_image(
            image=self.task_env_config.docker_image,
        )
    else:
        template = Template().from_dockerfile(
            dockerfile_content_or_path=str(self._environment_definition_path),
        )
```

**After:**
```python
async def _create_template(self):
    if self.task_env_config.docker_image:
        template = Template().from_image(
            image=self.task_env_config.docker_image,
        )
    else:
        template = Template(
            file_context_path=str(self.environment_dir),
        ).from_dockerfile(
            dockerfile_content_or_path=str(self._environment_definition_path),
        )
```

### Change 1B: Reduce sandbox timeout for E2B free tier

**Problem**: Hardcoded `timeout=86_400` (24 hours) is rejected by E2B free tier which caps at 1 hour.

**Location**: `_create_sandbox` method (~line 122)

**Before:**
```python
self._sandbox = await AsyncSandbox.create(
    template=self._template_name,
    metadata=metadata,
    timeout=86_400,
    allow_internet_access=self.task_env_config.allow_internet,
)
```

**After:**
```python
self._sandbox = await AsyncSandbox.create(
    template=self._template_name,
    metadata=metadata,
    timeout=3_600,
    allow_internet_access=self.task_env_config.allow_internet,
)
```

> **Note**: On E2B Pro plan (up to 24h), you can increase this back or make it configurable.

---

## File 2: `harbor/models/task/paths.py`

### Change 2: Add `data_dir` and `skills_data_dir` properties

**Problem**: No path definitions for the task's `data/` and `skills/` directories.

**Location**: After the `config_path` property

**Add these three properties:**
```python
@property
def data_dir(self) -> Path:
    """Path to the optional data/ directory for runtime input files."""
    return self.task_dir / "data"

@property
def skills_data_dir(self) -> Path:
    """Path to the optional skills/ directory for agent skills."""
    return self.task_dir / "skills"

@property
def prompts_dir(self) -> Path:
    """Path to the optional prompts/ directory for agent system prompts."""
    return self.task_dir / "prompts"
```

**Also update the docstring** at the top of the class to document the new directories:
```
├── instruction.md
├── task.toml
├── data/             # optional input files, uploaded to /workspace/inputs at runtime
│   └── ...
├── skills/           # optional skills, uploaded to /workspace/skills at runtime
│   └── ...
├── prompts/          # optional prompts, uploaded to /workspace/prompts at runtime
│   └── ...
├── environment/
│   ├── [docker-compose.yaml | Dockerfile | ...]
│   └── ...
```

---

## File 3: `harbor/trial/trial.py`

### Change 3A: Add `_upload_task_data` method

**Problem**: No mechanism to upload task input files and skills to non-mounted environments (E2B) at runtime.

**Location**: Before the `_maybe_upload_agent_logs` method

**Add this method:**
```python
async def _upload_task_data(self) -> None:
    """Upload task data, skills, and prompts to the sandbox for non-mounted environments.

    Uploads:
    - data/ directory    → /workspace/inputs  (dynamic user input files)
    - skills/ directory  → /workspace/skills  (agent skill definitions)
    - prompts/ directory → /workspace/prompts (agent system prompts)

    Skipped for mounted environments (e.g. Docker) where the host
    filesystem is already accessible.
    """
    if self._environment.is_mounted:
        return

    upload_map = [
        (self._task.paths.data_dir, "/workspace/inputs", "task data"),
        (self._task.paths.skills_data_dir, "/workspace/skills", "skills"),
        (self._task.paths.prompts_dir, "/workspace/prompts", "prompts"),
    ]

    for local_dir, remote_dir, label in upload_map:
        if local_dir.exists() and any(local_dir.iterdir()):
            self._logger.info(f"Uploading {label} from {local_dir}")
            try:
                await self._environment.upload_dir(
                    source_dir=local_dir,
                    target_dir=remote_dir,
                )
            except Exception:
                self._logger.error(f"Failed to upload {label} to environment")
                raise
```

### Change 3B: Call `_upload_task_data` in the `run()` method

**Location**: In the `run()` method, after `_setup_environment()` and before `_setup_agent()`

**Before:**
```python
await self._setup_environment()
self._environment.default_user = self._task.config.agent.user
await self._setup_agent()
```

**After:**
```python
await self._setup_environment()
await self._upload_task_data()
self._environment.default_user = self._task.config.agent.user
await self._setup_agent()
```

---

## File 4: `harbor/agents/installed/claude_code.py`

### Change 4: Add `system_prompt_file` CLI flag

**Problem**: No way to pass `--system-prompt-file` to Claude Code via job YAML kwargs. The existing custom `HarborClaudeCodeSystemPromptAgent` writes a file on the host which doesn't work for non-mounted environments (E2B). With runtime upload of `prompts/`, we just need a way to reference the file path inside the sandbox.

**Location**: In the `CLI_FLAGS` list of the `ClaudeCode` class, after the `append_system_prompt` flag

**Add this flag:**
```python
CliFlag(
    "system_prompt_file",
    cli="--system-prompt-file",
    type="str",
),
```

**Usage in job YAML:**
```yaml
agents:
  - name: claude-code
    model_name: anthropic/claude-sonnet-4-20250514
    kwargs:
      system_prompt_file: /workspace/prompts/system-prompt.md
```

This replaces the need for the custom `HarborClaudeCodeSystemPromptAgent` and its `import_path` usage. The system prompt file lives in the task's `prompts/` directory, gets uploaded to `/workspace/prompts/` at runtime, and Claude Code reads it directly from there.

---

## Summary

| # | File | Change | Purpose |
|---|------|--------|---------|
| 1A | `environments/e2b.py` | Add `file_context_path` to `Template()` | Fix COPY path resolution in Dockerfiles |
| 1B | `environments/e2b.py` | `timeout=86_400` → `3_600` | E2B free-tier compatibility |
| 2 | `models/task/paths.py` | Add `data_dir` + `skills_data_dir` + `prompts_dir` properties | Define paths for runtime upload dirs |
| 3A | `trial/trial.py` | Add `_upload_task_data()` method | Upload data→inputs, skills→skills, prompts→prompts |
| 3B | `trial/trial.py` | Call `_upload_task_data()` in `run()` | Wire it into the trial execution flow |
| 4 | `agents/installed/claude_code.py` | Add `system_prompt_file` CliFlag | Enable `--system-prompt-file` via job YAML kwargs |

> **Note**: Skills registration does NOT require a Harbor code change. Harbor already supports
> `skills_dir` natively via `task.toml`. Set `skills_dir = "/workspace/skills"` under
> `[environment]` and the stock `ClaudeCode` agent will copy them into its config directory.

## Task Directory Structure (after changes)

```
tasks/my-task/
├── instruction.md          ← dynamic per request
├── data/                   ← dynamic per request → uploaded to /workspace/inputs
│   ├── input.csv
│   └── template.xlsx
├── skills/                 ← mostly static → uploaded to /workspace/skills
│   ├── xlsx/
│   ├── docx/
│   └── ...
├── prompts/                ← static/semi-static → uploaded to /workspace/prompts
│   └── system-prompt.md
├── task.toml               ← static
├── environment/            ← static (baked into E2B template, never rebuilt)
│   ├── Dockerfile
│   └── docker-compose.yaml
└── tests/
    └── test.sh
```

## Job YAML Example (E2B with system prompt)

```yaml
environment:
  type: e2b
  force_build: false
  delete: true
  env:
    - ANTHROPIC_API_KEY=${ANTHROPIC_API_KEY}

agents:
  - name: claude-code
    model_name: anthropic/claude-sonnet-4-20250514
    kwargs:
      system_prompt_file: /workspace/prompts/system-prompt.md

artifacts:
  - /workspace

tasks:
  - path: tasks/my-task
```

## Behavior

- **E2B**: Template built once from Dockerfile (dependencies only). `data/`, `skills/`, and `prompts/` uploaded at runtime via E2B SDK. No template rebuild when files change.
- **Docker**: Upload steps are no-ops (`is_mounted` returns `True`). Docker uses volume mounts, so host files are already accessible.
- **All directories are optional**: If `data/`, `skills/`, or `prompts/` don't exist, the upload step is silently skipped.
- **System prompt**: Passed via `--system-prompt-file` flag to Claude CLI, referencing the uploaded file inside the sandbox. No need for the custom `HarborClaudeCodeSystemPromptAgent`.
