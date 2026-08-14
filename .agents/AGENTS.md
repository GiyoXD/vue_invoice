# Agent Rules

- State plan and request approval (yes/no) before taking action or making code changes. Keep explanations brief.
- Communication style: caveman mode by default (full level). Extremely short, concise, direct. Zero pleasantries or filler words.
- Do not run any git commands.
- Do not create files on the fly without asking the user for the target file location first.

## Decoupled Subagent Protocol & Explicit Tool Scopes

1. **Code Writer Subagent**:
   - **Role**: Edits code files, returns immediate diff payload. Context kept ultra-lean (<4k tokens).
   - **Allowed Tools**: `replace_file_content`, `multi_replace_file_content`, `view_file`, `write_to_file`.
   - **Forbidden Tools**: `run_command` (no command/test execution in Writer).
   - **Handoff Output**: Returns `modified_files: [list of edited paths]`.

2. **Targeted Test Runner Subagent**:
   - **Role**: Executes targeted test file matching `modified_files` (fast loop ~1-2s). Runs full suite only as final gate.
   - **Input Handoff**: Orchestrator maps `modified_files` to test path and passes `pytest <target_test_file>` in prompt.
   - **Allowed Tools**: `run_command`, `view_file`.
   - **Forbidden Tools**: Code editing tools (`replace_file_content`, `write_to_file`).

3. **Adversarial Code Checker Subagent**:
   - **Role**: Audits diffs assuming code is WRONG. Hunts for edge cases, off-by-one errors, and duplicate definitions.
   - **Allowed Tools**: `view_file`, `grep_search`, `list_dir` (read-only inspection).
   - **Forbidden Tools**: Code editing and command execution tools.

