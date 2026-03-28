# Harbor POC — Getting Started Guide

This guide explains how to run AI agent tasks using Harbor with E2B sandboxes, and how to create your own tasks.

---

## Prerequisites

1. **Harbor CLI** installed (`uv tool install git+https://github.com/dineshcx/harborcx.git` from the forked repo)
2. **E2B API key** — sign up at [e2b.dev](https://e2b.dev) and set `E2B_API_KEY`
3. **Anthropic API key** — set `ANTHROPIC_API_KEY` (or `OPENAI_API_KEY` for Codex)

```bash
export E2B_API_KEY=your_e2b_key
export ANTHROPIC_API_KEY=your_anthropic_key
export OPENAI_API_KEY=your_openai_key
```

---

## Quick Start

Run the example task:

```bash
harbor run -c job.yaml --job-name my-first-run -y
```

This will:
1. Build an E2B sandbox template from the Dockerfile (first run only)
2. Upload input files, skills, and prompts to the sandbox
3. Start Claude Code inside the sandbox with the task instruction
4. Download artifacts (output files) and agent trajectory when done

Results appear in `jobs/my-first-run/`.

---

## Project Structure

```
├── job.yaml                              ← Job configuration (which agent, environment, tasks)
├── jobs/                                 ← Where results are saved (created after first run)
└── tasks/
    └── simple-test/                      ← A single task
        ├── instruction.md                ← What the agent should do (the prompt)
        ├── task.toml                     ← Task metadata and config
        ├── data/                         ← Input files for the agent
        │   └── students.csv
        ├── skills/                       ← Claude Code skills (document generation helpers)
        │   ├── xlsx/SKILL.md
        │   ├── docx/SKILL.md
        │   ├── pptx/SKILL.md
        │   ├── pdf/SKILL.md
        │   └── jupyter-notebook/SKILL.md
        ├── prompts/                      ← System prompt for the agent
        │   └── system-prompt.md
        └── environment/                  ← Sandbox environment definition
            └── Dockerfile
```

---

## How Each Piece Works

### `instruction.md` — The Task Prompt

This is what the agent sees as its instruction. Write it like you're telling a person what to do.

```markdown
Read the CSV file at /workspace/inputs/students.csv which contains student names and scores.
Create an Excel file called report.xlsx with one worksheet named "Summary".
Save the file to /workspace/report.xlsx.
```

**Key rule**: Input files are available at `/workspace/inputs/` inside the sandbox. Reference them with that path.

---

### `task.toml` — Task Configuration

Minimal configuration for the task:

```toml
version = "1.0"

[metadata]
taskId = "my-task-name"
category = "Reports"
difficulty = "1"

[environment]
skills_dir = "/workspace/skills"
```

| Field | Purpose |
|---|---|
| `taskId` | Unique identifier for this task |
| `category` | Grouping label (for organization) |
| `difficulty` | Difficulty rating (informational) |
| `skills_dir` | Tells the agent where to find skills inside the sandbox. Set to `/workspace/skills` so Harbor automatically registers uploaded skills into Claude Code |

---

### `data/` — Input Files

Place any files the agent needs to work with here. These are uploaded at runtime to `/workspace/inputs/` inside the sandbox.

Examples:
- CSV/Excel files to process
- PDF documents to analyze
- Template files to fill in
- Images, text files, etc.

**No Dockerfile rebuild needed** when you change input files — they're uploaded fresh every run.

---

### `skills/` — Agent Skills

Skills teach the agent *how* to generate specific file types correctly. Each skill is a directory with a `SKILL.md` file that contains instructions and helper scripts.

Available skills in this POC:

| Skill | Purpose |
|---|---|
| `xlsx/` | Excel spreadsheet generation (openpyxl, formulas, formatting) |
| `docx/` | Word document generation (python-docx or JS docx library) |
| `pptx/` | PowerPoint presentation generation (python-pptx or pptxgenjs) |
| `pdf/` | PDF generation (reportlab, weasyprint, fpdf2) |
| `jupyter-notebook/` | Jupyter notebook creation |

Skills are uploaded to `/workspace/skills/` and registered into Claude Code's config directory automatically (via the `skills_dir` setting in `task.toml`).

**No Dockerfile rebuild needed** when you change skills.

---

### `prompts/` — System Prompt

The system prompt gives the agent its persona and behavioral guidelines. Place your system prompt markdown file here.

```
prompts/
└── system-prompt.md
```

Referenced in `job.yaml` via:
```yaml
kwargs:
  system_prompt_file: /workspace/prompts/system-prompt.md
```

**Important**: The system prompt should contain domain instructions only (how to format documents, quality standards, style guidelines). Do NOT include tool-specific instructions — Claude Code has its own tools (`Bash`, `Read`, `Write`, `Edit`, etc.) and will get confused if the prompt references other tool names.

**No Dockerfile rebuild needed** when you change the system prompt.

---

### `environment/Dockerfile` — Sandbox Environment

Defines what software is pre-installed in the sandbox. This is the **static** part — it gets built into an E2B template once and reused across runs.

The current Dockerfile includes:
- **Python 3.12** with data/document libraries (pandas, openpyxl, python-docx, python-pptx, matplotlib, reportlab, etc.)
- **Node.js 20** with document libraries (docx, pptxgenjs)
- **System tools**: LibreOffice, Pandoc, Graphviz, wkhtmltopdf, etc.
- **AI agents**: Claude Code, OpenAI Codex, Gemini CLI, OpenCode

**When to rebuild**: Only when you add/remove/update libraries in the Dockerfile. Set `force_build: true` in `job.yaml` for that run, then set it back to `false`.

---

## `job.yaml` — Job Configuration

This is the main config file that ties everything together:

```yaml
jobs_dir: jobs                      # Where results are saved
n_attempts: 1                       # Retry count per task
timeout_multiplier: 1.0
agent_timeout_multiplier: 5         # Agent timeout = base * this

orchestrator:
  type: local                       # Run locally
  n_concurrent_trials: 1            # One task at a time

environment:
  type: e2b                         # Use E2B sandboxes (not Docker)
  force_build: false                # true = rebuild template, false = reuse cached
  delete: true                      # Clean up sandbox after run
  env:
    - ANTHROPIC_API_KEY=${ANTHROPIC_API_KEY}

verifier:
  disable: true                     # No automated scoring

agents:
  - name: claude-code               # Agent to use
    model_name: anthropic/claude-opus-4-6
    kwargs:
      reasoning_effort: high
      system_prompt_file: /workspace/prompts/system-prompt.md

artifacts:
  - /workspace                      # Download everything from /workspace after run

tasks:
  - path: tasks/simple-test         # Which task(s) to run
```

### Key settings to know

| Setting | What it does |
|---|---|
| `environment.type` | `e2b` for cloud sandboxes, `docker` for local Docker |
| `environment.force_build` | `true` rebuilds the template from Dockerfile. Set to `false` after first successful build |
| `environment.env` | Environment variables passed into the sandbox |
| `agents[].name` | Agent name: `claude-code`, `codex`, `gemini-cli`, `opencode` |
| `agents[].model_name` | Model to use (e.g., `anthropic/claude-opus-4-6`, `openai/gpt-5.4`) |
| `agents[].kwargs.system_prompt_file` | Path to system prompt file inside the sandbox |
| `agents[].kwargs.reasoning_effort` | `low`, `medium`, `high` (model-dependent) |
| `artifacts` | Paths to download from sandbox after the run |
| `tasks[].path` | Path to the task directory |

---

## Running a Task

### First run (builds the E2B template):

```bash
# Set force_build: true in job.yaml, then:
harbor run -c job.yaml --job-name first-run -y
```

This takes several minutes as it builds the Docker image on E2B. After this, set `force_build: false`.

### Subsequent runs (reuses cached template):

```bash
harbor run -c job.yaml --job-name my-run-name -y
```

This is fast — just spins up a sandbox, uploads files, runs the agent.

### Run with a different job name each time:

```bash
harbor run -c job.yaml --job-name experiment-1 -y
harbor run -c job.yaml --job-name experiment-2 -y
```

---

## Output Structure

After a run, results appear in `jobs/<job-name>/`:

```
jobs/my-run-name/
├── config.json                              ← Job configuration snapshot
├── job.log                                  ← High-level job log
├── result.json                              ← Overall results and metrics
└── simple-test__<id>/                       ← Per-task trial directory
    ├── trial.log                            ← Detailed trial log
    ├── config.json                          ← Trial config
    ├── result.json                          ← Trial result (score, metrics)
    ├── agent/
    │   ├── claude-code.txt                  ← Raw agent output log
    │   └── trajectory.json                  ← ATIF-v1.2 trajectory (every step the agent took)
    └── artifacts/
        └── workspace/                       ← All files the agent created
            ├── report.xlsx
            ├── presentation.pptx
            └── ...
```

### Key files to check:
- **`agent/claude-code.txt`** — see what the agent did step by step
- **`agent/trajectory.json`** — structured trajectory in ATIF format (for programmatic analysis)
- **`artifacts/workspace/`** — the actual output files the agent generated

---

## Creating a New Task

1. Create a new directory under `tasks/`:

```bash
mkdir -p tasks/my-new-task/{data,prompts,environment}
```

2. **Write the instruction** (`instruction.md`):

```markdown
Analyze the sales data at /workspace/inputs/sales.csv and create:
1. An Excel summary with pivot tables at /workspace/sales_report.xlsx
2. A PDF chart showing monthly trends at /workspace/trends.pdf
```

3. **Add input files** to `data/`:

```bash
cp your-data-file.csv tasks/my-new-task/data/
```

4. **Create `task.toml`**:

```toml
version = "1.0"

[metadata]
taskId = "my-new-task"
category = "Analysis"
difficulty = "2"

[environment]
skills_dir = "/workspace/skills"
```

5. **Copy or symlink the shared environment and skills** (reuse from simple-test):

```bash
cp -r tasks/simple-test/environment/Dockerfile tasks/my-new-task/environment/
cp -r tasks/simple-test/skills tasks/my-new-task/skills
cp -r tasks/simple-test/prompts tasks/my-new-task/prompts
```

6. **Update `job.yaml`** to point to your task:

```yaml
tasks:
  - path: tasks/my-new-task
```

7. **Run it**:

```bash
harbor run -c job.yaml --job-name my-new-task-run -y
```

---

## What Changes Require a Template Rebuild?

| What changed | Rebuild needed? | Action |
|---|---|---|
| `instruction.md` | No | Just re-run |
| Files in `data/` | No | Just re-run |
| Files in `skills/` | No | Just re-run |
| Files in `prompts/` | No | Just re-run |
| `task.toml` | No | Just re-run |
| `job.yaml` (agent, model, kwargs) | No | Just re-run |
| `environment/Dockerfile` | **Yes** | Set `force_build: true`, run once, set back to `false` |

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `SandboxException` | E2B sandbox failed to start | Check E2B API key, check Dockerfile syntax |
| `Multi-stage Dockerfiles are not supported` | E2B doesn't support `FROM ... AS ...` | Use single-stage Dockerfile only |
| `Timeout cannot be greater than 1 hours` | E2B free tier limit | Our Harbor fork already fixes this (3600s) |
| Agent uses wrong tools (`view`, `create_file`) | System prompt has tool instructions from claude.ai web | Remove tool-specific content from system prompt |
| `Failed to download artifact '/workspace'` | Agent errored before creating files | Check `agent/claude-code.txt` for the error |
| Template rebuild every run | `force_build: true` left on | Set `force_build: false` after first successful build |

---

## Harbor Fork Changes

We maintain a small set of patches to Harbor for E2B compatibility. See `HARBOR_FORK_CHANGES.md` for the full list of changes with before/after code.
