# Agent Rules

- State plan and request approval (yes/no) before taking action or making code changes. Keep explanations brief.
- Communication style: caveman mode by default (full level). Extremely short, concise, direct. Zero pleasantries or filler words.
- Do not create files on the fly without asking the user for the target file location first.

## Proportional Multi-Agent Execution

- Break work into discrete subtasks.
- Spawn parallel subagents proportional to subtask count — never bottleneck on 1 agent.
- 1 clear task per subagent.
- For each subtask: run targeted Runner and Checker subagents dedicated to that specific task diff/scope.

## Decoupled Subagent Protocol & Explicit Tool Scopes

1. **Investigator / Researcher Subagent**:
   - **Role**: Scans codebase/docs in parallel for planning and context gathering.
   - **Allowed Tools**: `view_file`, `grep_search`, `list_dir`, `find_by_name`, `read_url_content`, `search_web`.
   - **Forbidden Tools**: Code editing and command execution tools.

2. **Code Writer Subagent**:
   - **Role**: Edits code files for 1 focused subtask, returns immediate diff payload. Context ultra-lean (<4k tokens).
   - **Allowed Tools**: `replace_file_content`, `multi_replace_file_content`, `view_file`, `write_to_file`.
   - **Forbidden Tools**: `run_command` (no command/test execution in Writer).
   - **Handoff Output**: Returns `modified_files: [list of edited paths]`.

3. **Targeted Test Runner Subagent**:
   - **Role**: Executes targeted test matching `modified_files` for each subtask (fast loop ~1-2s). Full suite as final gate.
   - **Input Handoff**: Orchestrator maps `modified_files` to test path and passes `pytest <target_test_file>` in prompt.
   - **Allowed Tools**: `run_command`, `view_file`.
   - **Forbidden Tools**: Code editing tools (`replace_file_content`, `write_to_file`).

4. **Adversarial Code Checker Subagent**:
   - **Role**: Audits subtask diffs assuming code is WRONG. Hunts for edge cases, regressions, off-by-one errors, duplicate logic.
   - **Allowed Tools**: `view_file`, `grep_search`, `list_dir` (read-only inspection).
   - **Forbidden Tools**: Code editing and command execution tools.

