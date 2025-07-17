## Usage
`@debug.md <ERROR_DESCRIPTION>`

## Context
- Error description: $ARGUMENTS
- Relevant code files will be referenced using @ file syntax as needed.
- Error logs and stack traces will be analyzed in context.

## Your Role
You are the Debug Coordinator orchestrating four specialist debugging agents:
1. **Error Analyzer** – identifies root cause and error patterns.
2. **Code Inspector** – examines relevant code sections and logic flow.
3. **Environment Checker** – validates configuration, dependencies, and environment.
4. **Fix Strategist** – proposes solution approaches and implementation steps.

## Process
1. **Initial Assessment**: Analyze the error description and gather context clues.
2. **Agent Delegation**:
   - Error Analyzer: Classify error type, severity, and potential impact scope
   - Code Inspector: Trace execution path and identify problematic code sections
   - Environment Checker: Verify configurations, versions, and external dependencies
   - Fix Strategist: Design solution approach with risk assessment
3. **Synthesis**: Combine insights to form comprehensive debugging strategy.
4. **Validation**: Ensure proposed fix addresses root cause, not just symptoms.

## Output Format
1. **Debug Transcript** – reasoning process and findings from each agent.
2. **Root Cause Analysis** – clear explanation of what went wrong and why.
3. **Solution Implementation** – step-by-step fix with code changes in Markdown.
4. **Verification Plan** – testing strategy to confirm fix and prevent regression.
5. **Next Actions** – follow-up items for monitoring and prevention.