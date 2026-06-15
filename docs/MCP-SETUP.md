# Registering AIOffice as an MCP server · 接入 MCP

> **中文摘要** — AIOffice 内置一个 stdio MCP server：`aioffice mcp --workspace <目录>`。它把 17 个 MCP 工具（与 18 个 CLI 动词一一对应）暴露给 AI 主机。`--workspace`（或环境变量 `AIOFFICE_WORKSPACE`，默认当前目录）定义一个沙箱根，所有文件读写都被限制在其中，越界即 `sandbox_denied`。下面给出 Claude Desktop、Claude Code、Cursor、通用 stdio 客户端、以及 TonoBraid（`~/.tonoagent/mcp.json`）的精确配置片段。建议把 agent 的系统提示指向 [SKILL.md](../SKILL.md)（仓库根，AI 上手指南）。

AIOffice exposes the **same surface twice**: as a CLI and as an MCP server. The MCP server runs over **stdio** and registers **17 tools** that mirror the CLI verbs 1:1, going through one shared internal command layer. One mental model, two entry points.

```sh
aioffice mcp --workspace /path/to/your/documents
```

- **`--workspace <dir>`** (or the `AIOFFICE_WORKSPACE` env var; default: current directory) is the **sandbox root**. Every `file` / `output` argument is realpath- and symlink-escape-checked on each call; anything resolving outside the workspace returns `sandbox_denied`. Point it at the folder the agent is allowed to touch.
- The server speaks JSON-RPC over stdin/stdout. Each tool result is the same JSON envelope you get on the CLI: `{ok, data, error{code,message,suggestion,candidates?}, meta{rev,elapsedMs,version,warnings?}}`.
- No network, no external engine, no binary downloads at runtime — OOXML read/write is 100% in-process.

If you installed via npm and want zero global state, you can use `npx` as the command (`npx aioffice mcp …`); if you installed globally or via Homebrew/script, use the `aioffice` command directly. The examples below show both.

---

## The 17 tools · 工具一览

| MCP tool | CLI verb | What it does |
|---|---|---|
| `office_create` | `create` | New document (or import `.md`→`.docx`, `.csv`→`.xlsx`) |
| `office_read` | `read` | Cheap inspection views (outline/text/stats/structure/properties/embeds/markdown/csv) |
| `office_query` | `query` | CSS-like selectors → canonical paths |
| `office_get` | `get` | One node + its properties |
| `office_edit` | `edit` | Atomic batch set/add/remove/move/replace/accept/reject/extract + find-replace |
| `office_render` | `render` | The *look* step — html/svg/text/png/pdf |
| `office_validate` | `validate` | OOXML validation + lint with fix suggestions |
| `office_audit` | `audit` | Accessibility + quality lint (findings are data; `--fix` safe autofixes) |
| `office_diff` | `diff` | Semantic compare vs another file or a snapshot |
| `office_convert` | `convert` | Cross-format conversion (docx/xlsx/pptx ↔, ↔md/csv, →pdf/png/svg/html) |
| `office_template` | `template` | `{{key}}` merge / mail-merge across docx/xlsx/pptx |
| `file_snapshot` | `snapshot` | Pre-edit snapshot ring (20), list / restore |
| `office_status` | `doctor` | Environment + runtime + `capabilities` self-report (incl. `surfaceVersion`) |
| `office_help` | `help` | Built-in docs: addressing, selectors, errors, per-feature topics |
| `office_schema` | `schema` | Machine-readable JSON of the whole surface |
| `preview_open` | `preview open` | Live localhost preview |
| `preview_selection` | `preview selection` | Human clicks in the preview → canonical paths |

Total tool-schema budget is capped at 3,500 tokens (enforced by a test). `surfaceVersion` is `1.0` — the frozen AI-facing contract in [CONTRACT.md](../CONTRACT.md).

---

## Claude Desktop

Edit `claude_desktop_config.json`:
- **macOS** — `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows** — `%APPDATA%\Claude\claude_desktop_config.json`

Add (or merge into) the `mcpServers` block:

```json
{
  "mcpServers": {
    "aioffice": {
      "command": "aioffice",
      "args": ["mcp", "--workspace", "/Users/you/Documents/aioffice"]
    }
  }
}
```

If `aioffice` is not on the GUI app's `PATH`, use an absolute path to the binary (e.g. `~/.local/bin/aioffice`, the Homebrew path from `which aioffice`, or the npm global bin) as `command`. Restart Claude Desktop after editing. You can also set the workspace via env instead of `args`:

```json
{
  "mcpServers": {
    "aioffice": {
      "command": "/Users/you/.local/bin/aioffice",
      "args": ["mcp"],
      "env": { "AIOFFICE_WORKSPACE": "/Users/you/Documents/aioffice" }
    }
  }
}
```

---

## Claude Code

Add the server with the CLI (project scope by default; add `--scope user` for all projects):

```sh
claude mcp add aioffice -- aioffice mcp --workspace "$PWD"
```

Using the npm package without a global install:

```sh
claude mcp add aioffice -- npx aioffice mcp --workspace "$PWD"
```

List / remove:

```sh
claude mcp list
claude mcp remove aioffice
```

This writes an `mcpServers` entry equivalent to the Claude Desktop JSON above into your Claude Code config.

---

## Cursor

Create `.cursor/mcp.json` in your project (or `~/.cursor/mcp.json` for a global server):

```json
{
  "mcpServers": {
    "aioffice": {
      "command": "aioffice",
      "args": ["mcp", "--workspace", "${workspaceFolder}"]
    }
  }
}
```

Cursor expands `${workspaceFolder}` to the open project root, which makes that the sandbox. Enable the server in **Cursor Settings → MCP**. Swap `command` to an absolute path (or `npx` + `["aioffice","mcp", …]`) if `aioffice` is not on Cursor's `PATH`.

---

## Generic stdio MCP client

Any MCP host that launches a stdio subprocess can use AIOffice. The contract is just:

- **command**: `aioffice` (or an absolute path to the binary, or `npx aioffice`)
- **args**: `["mcp", "--workspace", "<sandbox dir>"]`
- **transport**: stdio (JSON-RPC framed on stdin/stdout)
- **env** (optional): `AIOFFICE_WORKSPACE=<sandbox dir>` instead of the `--workspace` arg

Smoke-test the server by hand — initialize and list tools over stdio:

```sh
printf '%s\n%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"smoke","version":"0"}}}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' \
  | aioffice mcp --workspace .
```

You should see an `initialize` result followed by a `tools/list` result enumerating the 17 tools.

---

## TonoBraid

TonoBraid reads custom stdio MCP servers from `~/.tonoagent/mcp.json` (array form). Add an entry:

```json
[
  {
    "name": "aioffice",
    "command": "aioffice",
    "args": ["mcp", "--workspace", "/Users/you/Documents/aioffice"]
  }
]
```

Use an absolute `command` path if `aioffice` is not on TonoBraid's `PATH`. After editing, restart the agent so it re-reads `mcp.json`.

---

## Workspace sandbox · 沙箱说明

The `--workspace` directory is a hard boundary, not a default:

- Relative `file` paths resolve **inside** the workspace; absolute paths and `..` traversals that escape it are rejected with `sandbox_denied` (exit 4 on the CLI).
- Symlinks are resolved and re-checked, so a symlink inside the workspace can't be used to escape it.
- Snapshots, previews, and rendered outputs all stay within the sandbox.

Give each agent the narrowest workspace that still contains the documents it needs. To change it per session, restart the server with a different `--workspace` (or `AIOFFICE_WORKSPACE`).

---

## Agent system-prompt pointer · 系统提示建议

Point the agent's system prompt at the repo's **`SKILL.md`** (the AI-facing onboarding guide, at the repo root) so it knows the envelope shape, the addressing grammar (`/body/p[3]`, `/Sheet1/A1`, `/slide[2]/shape[3]`), and the read-before-write workflow. If you are wiring this up before `SKILL.md` lands, the same guidance is always available live from the server itself via `office_help` and from [docs/MCP.md](MCP.md). A minimal pointer:

> You have the **aioffice** MCP server (17 tools mirroring the AIOffice CLI). Operate on .docx/.xlsx/.pptx inside the workspace sandbox. Every tool returns `{ok,data,error,meta}`; on `ok:false`, read `error.suggestion` and `error.candidates`. Inspect before editing (`office_read` / `office_query` / `office_get`), then make atomic `office_edit` batches. See SKILL.md and `office_help` for addressing, selectors, and per-feature topics.

For the full machine-readable surface at runtime, the agent can call `office_schema` (whole surface) and `office_help {topic}` (per-feature docs). See also [docs/MCP.md](MCP.md) for the complete per-tool specification.
