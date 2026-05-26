# Human Validation Input Policy

This instruction is mandatory for every run, including a new session.

1. When the program displays the validation prompt `Enter Yes/No and press Enter to continue`, execution must pause and wait for explicit human input.
2. Do not auto-enter `Yes` or `No`.
3. Continue only after a human explicitly types `Yes` or `No` in the active session.
4. If no interactive input stream is available, stop the run and keep the generated validation note for manual review.
