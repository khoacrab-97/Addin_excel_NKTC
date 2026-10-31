# ConstructionDailyLog — Development Governance

## 1. Project Context

- **Project:** ConstructionDailyLog
- **Product:** Microsoft Excel Add-in for construction daily log management and template mapping.
- **Environment:**
  - Windows
  - Claude Code Desktop
  - GitHub Repository
  - Microsoft Excel Desktop
  - Office.js
  - React
  - TypeScript

## 2. Core Development Rules

Before implementing any task:

1. Read `MASTER_CONTEXT.md`.
2. Read `ROADMAP.md`.
3. Read relevant business and technical specifications.
4. Inspect the existing source code.
5. Check Git status and current branch.
6. Identify dependencies and affected modules.
7. Prepare an implementation plan.

Never assume a feature is missing without inspecting the codebase.

## 3. Code Quality

- Use TypeScript strict mode.
- Follow modular architecture.
- Keep business logic independent from Excel APIs.
- Use actual Office.js APIs.
- Do not invent unsupported APIs.
- Avoid duplicated business logic.
- Avoid unnecessary dependencies.
- Validate all external and user-entered data.
- Handle asynchronous operations correctly.
- Use clear types and meaningful names.
- Keep components maintainable and reusable.

## 4. Excel Safety

- Never modify cells outside the configured target mapping.
- Preserve existing formulas and formatting.
- Never overwrite worksheet data without explicit confirmation.
- Use batch Excel operations where appropriate.
- Handle missing worksheets and invalid ranges.
- Avoid unnecessary `Excel.run` calls.
- Read the latest source worksheet values before mapping.
- Do not rely on stale cached data.
- Test on Excel Desktop Windows.

## 5. Feature Protection

- Never remove an existing completed feature without approval.
- Never rewrite unrelated modules.
- Preserve existing public interfaces unless changes are justified.
- Add regression tests for critical business logic.
- Review the impact of changes on existing functionality.

## 6. Development Workflow

For every task:

1. Analyze requirements.
2. Inspect current implementation.
3. Identify affected files.
4. Define acceptance criteria.
5. Implement.
6. Run lint and type checks.
7. Run relevant unit tests.
8. Run build.
9. Perform integration tests when applicable.
10. Update documentation.
11. Report actual results.

## 7. Roadmap Status

Use only these statuses:

| Status | Definition |
|---|---|
| Not Started | Work has not begun. |
| In Progress | Work is underway. |
| Implemented | Code has been written. |
| Tested | Relevant automated tests have passed. |
| Verified | The feature has been tested in its intended real environment. |

Never mark a feature **Verified** based only on successful compilation.

## 8. Git Rules

- Never delete Git history.
- Never force push.
- Never commit secrets.
- Never push without user approval.
- Inspect `git diff` before committing.
- Use meaningful commit messages.
- Keep `main` stable.
- Use feature branches for substantial development when appropriate.

## 9. Completion Report

Every completed task must report:

1. Scope completed.
2. Files created.
3. Files modified.
4. Files deleted, if any.
5. Technical decisions.
6. Tests executed.
7. Test results.
8. Known limitations.
9. Documentation updated.
10. Recommended next task.

Do not claim success when tests have not been executed.
