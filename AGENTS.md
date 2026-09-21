# Repository Instructions

## Mission

Build a Windows-first agent that turns plain-language requests and user-provided materials into editable PowerPoint files, supports conversational revisions, and previews the actual exported PPTX.

## Read Before Changing Anything

Use the documents below as the source of truth. Read only the documents relevant to the task, plus the active execution plan for implementation work.

- Product scope and acceptance: [docs/PRODUCT.md](docs/PRODUCT.md)
- Architecture and technical boundaries: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
- Test strategy and quality gates: [docs/TESTING.md](docs/TESTING.md)
- Active MVP execution plan: [docs/exec-plans/active/mvp.md](docs/exec-plans/active/mvp.md)
- Original research and setup notes: [AI_PPT_Agent_开源项目与技术方案.md](AI_PPT_Agent_开源项目与技术方案.md)

## Long-Term Engineering Rules

- Keep the MVP Windows-native and runnable from PowerShell.
- Use Python 3.12 and `uv` for the initial application unless an approved architecture change says otherwise.
- Treat `DeckSpec` as the canonical source for outlines, rendering, revisions, previews, and versions.
- Give decks, slides, sources, and assets stable IDs. Never use a visible slide index as persistent identity.
- Convert conversational edits into a validated `PatchPlan`. Do not ask the model to regenerate untouched slides.
- Keep layout deterministic. Models choose supported layouts and content; application code owns coordinates, limits, and rendering.
- Generate native editable PPTX objects whenever the format supports them. Do not silently flatten a deck into full-slide images.
- Show final previews rendered from the exported PPTX, not from a separate approximation.
- Validate all model-produced structures before they can change project state or create a downloadable artifact.
- Preserve the last successful version when parsing, generation, rendering, preview, or validation fails.
- Keep model providers behind an adapter. Provider-specific behavior must not leak into domain models.
- Store secrets only in ignored local configuration. Never commit API keys, uploaded user material, generated workspaces, or private fonts/assets.
- Treat uploaded files and paths as untrusted input. Restrict reads and writes to the active project workspace.
- Do not introduce a database, queue, Node service, free-form canvas, or multi-agent framework into the MVP without updating the architecture document and active plan first.

## Change Discipline

- Keep tasks small enough to verify independently.
- Add or update tests with every behavior change.
- Run the narrowest relevant tests first, then the full required gate from [docs/TESTING.md](docs/TESTING.md).
- Do not mix unrelated refactors into a feature change.
- Preserve user-authored files and unrelated worktree changes.
- Record completed tasks, decisions, blockers, and newly discovered work in the active execution plan.
- A task is complete only when its documented test and acceptance condition pass.

## Documentation Ownership

- Change user outcomes, scope, or acceptance criteria in `docs/PRODUCT.md`.
- Change components, contracts, dependencies, or data flow in `docs/ARCHITECTURE.md`.
- Change test commands, fixtures, matrices, or quality gates in `docs/TESTING.md`.
- Change sequencing, progress, implementation notes, or task-level decisions in the active execution plan.
- Keep this file limited to durable repository rules and links. Do not copy detailed plans into it.

## Documentation and Command Style

- Write repository documentation in concise Chinese; keep code identifiers and command names in English.
- Use PowerShell examples for Windows commands.
- Use relative Markdown links for repository files.
- When the project becomes runnable, keep `README.md` quick-start commands accurate.

