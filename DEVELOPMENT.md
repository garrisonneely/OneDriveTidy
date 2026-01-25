# Development Guidelines & Best Practices

This document outlines the coding standards and best practices for the OneDriveTidy project.

## C# Coding Standards

### Naming Conventions
- **Classes/Interfaces**: PascalCase (e.g., `GraphService`, `IConfiguration`).
- **Methods**: PascalCase (e.g., `InitializeAsync`, `UploadFile`).
- **Variables/Parameters**: camelCase (e.g., `fileName`, `cancellationToken`).
- **Private Fields**: camelCase with underscore prefix (e.g., `_graphClient`, `_logger`).
- **Constants**: ALL_CAPS with underscores (e.g., `MAX_RETRIES`, `DEFAULT_TIMEOUT`).

### Asynchronous Programming
- Use `async`/`await` for all I/O bound operations.
- Appends `Async` suffix to asynchronous method names (e.g., `SaveDataAsync`).
- Always pass `CancellationToken` to async methods where supported.
- Avoid `async void` except for top-level event handlers.

### Dependency Injection
- Prefer Constructor Injection for all dependencies.
- Register services in `Program.cs` with appropriate lifetimes (`Singleton`, `Scoped`, or `Transient`).

### Formatting & Style
- Use file-scoped namespaces (C# 10+).
- Use `var` when the type is obvious from the right-hand side.
- Use explicit types when the type is not obvious.
- Braces should be on a new line (Allman style).

## Git Best Practices

### Commits
- **Atomic Commits**: Each commit should do one thing and do it well. Avoid bundling unrelated changes.
- **Commit Messages**:
  - Use the imperative mood (e.g., "Add feature" not "Added feature").
  - First line should be a short summary (50 chars or less).
  - Leave a blank line after the summary, then provide a detailed description if necessary.
- **Staging**: Review your staged changes (`git diff --staged`) before committing to ensure no unintended files are included.

### Branching Strategy
- **main**: The stable, production-ready branch. Do not commit directly to main.
- **Feature Branches**: Create a new branch for each feature or fix (e.g., `feature/add-excel-export`, `fix/upload-bug`).
- **Pull Requests**: Merge changes into `main` via Pull Requests (PRs) to allow for code review.
  
### Workflow
1. Pull latest `main`: `git checkout main && git pull`
2. Create branch: `git checkout -b feature/my-feature-name`
3. Make changes and commit: `git commit -m "Implement X"`
4. Push branch: `git push -u origin feature/my-feature-name`
5. Open Pull Request.

### Ignored Files
- Ensure `.gitignore` is properly configured to exclude:
  - Build artifacts (`bin/`, `obj/`)
  - User secrets and local configuration (`appsettings.Development.json`)
  - IDE specific files (`.vs/`, `.vscode/`)
  - Temporary files (`*.tmp`, `*.log`)
