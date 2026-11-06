# 0006. A built-in git client keeps its repository out of sight

## Status

Proposed.

## Context

Git sync (`docs/GitSyncCompare.md`) compares a project with a folder the user picked and keeps. Versioning is then done outside the editor: the user needs a folder, a repository in it, and another tool to commit, branch and look at history.

Upstream Rubberduck once had a Source Control panel built on LibGit2Sharp (Changes, Branches, Unsynced commits, Settings, Undo). It was off by default and marked experimental, was removed in February 2018 (`516d87160`, then `40f320525` took out LibGit2Sharp itself), and nothing of it is left. The surviving hints at why: it ran against the live VBE from UI code, so re-importing a module could throw COM exceptions in the Code Explorer, and a native dependency plus credential handling kept producing crashes.

What is different now: Git sync already solves the hard part, moving modules between the VBE and files through the VBE's own export and import (so hidden attributes survive), one module at a time on the calling thread, with `Scan` (changes nothing) and `Apply` (does what the user picked) as separate steps. It never calls git.

## Decision

1. **Git sync is the bridge; git is added on top.** A commit is `Scan`, `Apply` project to folder, then `git add -A` and `git commit`. A checkout is `git checkout`, then `Scan`, then `Apply` folder to project. The VBE is touched only by Git sync's code, never by the git client.
2. **The repository lives in `%LOCALAPPDATA%\Rubberduck\repos\<project-id>\`**, as two folders: `git` (the git directory) and `tree` (the work tree, which is the sync folder). Every call passes `--git-dir` and `--work-tree` explicitly. No `.git` file or folder is written anywhere the user sees. A setting chooses instead a visible folder or an existing repository, which is how a project joins CI, other tools or a team; the hidden store is the default, not the only way.
3. **The project id is kept with the project**, not derived from its path, so that moving or renaming the workbook does not lose the history. It is stored in the `rubberduck.json` kept in the document (see Resources in the glossary). Where the document has none (xls, doc, Access, an unsaved file), the id is looked up in a map from file path to project id in AppData, next to the one Git sync already keeps (`GitSyncLinks`). An unknown path whose project name matches a known repository asks whether it is that project moved, and relinks on yes. The map also holds each project's last-known path: if the id turns up at a different path and the old path no longer exists, it was a move and the path is updated silently; if the old path still exists as another file, it is a copy, and the user chooses between continuing the history there (a fork) and starting fresh.
4. **The engine is the `git` executable, run as a child process** with an argument list (never through a shell), `GIT_TERMINAL_PROMPT=0`, no pager, a timeout, and output read by porcelain formats (`status --porcelain=v2 -z`, `log` with a fixed format). Not LibGit2Sharp: it was dropped once, is a native dependency inside Excel's process, and lacks features users will expect (rebase, many merge strategies).
5. **Git is found, and if missing, installed once the user agrees.** Order: the path in settings, `git` on `PATH`, Rubberduck's own copy, `C:\Program Files\Git`. A version below the supported minimum counts as missing. If none is found, the first use of a git feature shows a dialog saying what will be downloaded, from where and how big, with **Install** and **Cancel**. Install then runs without further steps: it downloads **MinGit** (the Git for Windows project's portable build) from its GitHub release, verifies the SHA-256 against a hash pinned in Rubberduck for that version, unpacks it to `%LOCALAPPDATA%\Rubberduck\git\<version>`, and runs `git --version`. No administrator rights and nothing outside Rubberduck's folder; the user's system and `PATH` are untouched. It never installs without the click, and a failed download or hash mismatch deletes what it fetched and says so. Rejected: the full Git for Windows installer (needs elevation, changes `PATH`, wizard choices), and `winget` (absent on many corporate machines, still needs elevation).
6. **Off by default**, like the API: a setting turns the git client on, and nothing runs git or creates a repository until it is on.
7. **Same shape as Git sync**: scan first, show what will happen, apply what the user chose. The git client does not run unattended. API routes for it come after the editor works, and then need the approval writes need.
8. **Hostile repositories are expected.** Hooks run code. Rubberduck's own repositories are created with `core.hooksPath` pointing at an empty folder; a cloned or user-supplied repository is run with hooks off too, unless the user turns them on. `protocol.ext.allow=never`, and `safe.directory` is not widened.

9. **Rubberduck holds the credentials; git gets them for one call.** Git cannot prompt inside Excel (`GIT_TERMINAL_PROMPT=0`), and MinGit has no Git Credential Manager. When a remote call fails to authenticate, Rubberduck asks in its own dialog for a user name and a token, and saves them in the Windows Credential Manager (per Windows user, encrypted by the system), keyed by remote URL. For each remote call it passes them through environment variables (`GIT_CONFIG_COUNT`, `GIT_CONFIG_KEY_n`, `GIT_CONFIG_VALUE_n`, git 2.31 or later) as `http.<remote-url>.extraHeader`, so the header is sent to that remote only. The secret is never on a command line, in `.git/config`, in a remote URL, in the repository or in the log; the runner masks it in anything it logs. Settings lists the saved logins and deletes them. Optional, later: use the system's credential helper when a full Git for Windows is installed, and SSH through Windows' own OpenSSH client (`core.sshCommand`) with the user's `ssh-agent`. Rejected: `https://user:token@host/` remotes (the secret lands in the repository's config), `-c http.extraHeader=...` on the command line (visible to every process of the user), bundling GCM at first (a second program to pin and keep current).
10. **Staging is hidden at first.** A commit takes everything that changed; leaving a module out is `.rdsyncignore`. A staging UI waits for someone to ask.
11. **The installer works behind a proxy and offline.** It uses the system proxy with default credentials and TLS 1.2 or later. **Install from a file** takes a MinGit zip the user points at, held to the same pinned hash. **Choose git.exe** remains the last fallback.

## Phases

1. Hidden repository, commit, status, history, diff of one module, restore a module from a commit. (Needs the installer, since nothing works without git.)
2. Branches, checkout, revert.
3. Remotes, push, pull, credentials, merge with the existing compare window for conflicts.
4. Forms (`.frx`) and document modules edge cases, API routes.

## Consequences

- A history that is out of the user's sight is not in their backup, not in Dropbox and not on another machine. It is lost with the profile or `%LOCALAPPDATA%`. Phase 3's remote is the real answer; until then the UI says where the repository is and offers to open or move it.
- The project id must survive copying a workbook: a copy that keeps the id would share, and fork, the history. The copy has to be detected (the stored path no longer matches the file) and asked about.
- Git sync never deletes a module from a project. A checkout to a commit that lacks a module leaves it in the project, listed as *Only in project*, and the user decides. Checkout therefore cannot promise the project equals the commit.
- `rubberduck.json` and `rubberduck.sync.json` are in the tree and so in every commit. That is what they were designed for.
- Rubberduck now ships a hash it must keep current when it raises the MinGit version, and depends on the GitHub release being reachable. Behind a proxy or offline, the dialog offers **Choose git.exe** instead.
- Remotes authenticate with tokens (personal access tokens), not browser sign-in or MFA flows, until the system helper or GCM is added. A host that accepts nothing but those is out of reach at first. Whether MinGit brings its own `ssh.exe`, and whether GCM has a portable build, are to be checked before either is relied on.

## Open

- Whether the credential dialog should offer the browser/device flow for GitHub and Azure DevOps, and so need an OAuth app registered for Rubberduck.
