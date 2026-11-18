# AI Survey Invite Helper

A small Python tool that writes personalised survey invitations for doctors from an Excel list, then **checks every invitation** before a human sends it.

**READ → WRITE → CHECK**

1. **READ**: loads each doctor from `doctors.xlsx` (name, specialty, survey topic, length, payment, deadline).
2. **WRITE**: creates a personal invitation, either from a simple template or with an AI model (AI mode).
3. **CHECK**: tests each invitation for the right name, payment and deadline, and makes sure it is under 150 words. Each one is marked **PASS** or **FIX**, with the reason.

The results are saved to `invites_output.xlsx`, with a **"Human approved? (Y/N)"** column. The tool never sends emails by itself; a person reviews everything first.

## Example result

The sample list has 8 doctors, and 2 mistakes were planted on purpose to test the checker:

```
PASS | Dr. Rao      | all checks OK
PASS | Dr. Mehta    | all checks OK
PASS | Dr. Iyer     | all checks OK
FIX  | Dr. Sen      | Payment wrong or missing
PASS | Dr. Kapoor   | all checks OK
FIX  | Dr. Ghosh    | Deadline wrong or missing
PASS | Dr. Hegde    | all checks OK
PASS | Dr. Das      | all checks OK

6/8 invites passed. Saved to invites_output.xlsx
```

The checker caught both planted mistakes.

## How to run

```bash
pip install -r requirements.txt
python3 invite_helper.py
```

## AI mode (optional)

By default the tool uses a template, so it runs without any API key.

To let an AI model write the emails instead:

1. Set `USE_AI = True` in `invite_helper.py`.
2. Set your Anthropic API key as an environment variable: `export ANTHROPIC_API_KEY=your_key_here`
3. Run the script as usual. The same checks run on the AI-written emails.

## Files

| File | What it is |
|---|---|
| `invite_helper.py` | The program (read, write, check) |
| `doctors.xlsx` | Sample input: 8 fictional doctors |
| `invites_output.xlsx` | Output: each invite + PASS/FIX + problems |
| `requirements.txt` | Python packages needed |

## Limits and next steps

- AI mode needs an API key, and the template mode is what was tested on the sample data.
- Next ideas: check for banned words (for example medical claims), write reminder emails for doctors who didn't respond, and test two subject lines to see which gets more opens.

All doctor data in this repository is fictional.
