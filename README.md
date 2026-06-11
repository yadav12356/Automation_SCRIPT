# FM Onboarding AI Automation

Single-file FM onboarding automation entrypoint for AI/Freshdesk integration.

## Run

```sh
pip install -r requirements.txt
python ai_fm_onboarding_single.py /path/to/ops_file.xlsx
```

The script validates the Ops file, resolves/creates loczipcode when configured,
generates new/existing partner upload CSVs, triggers the correct onboarding
Jenkins job, and prints one JSON response with `ticket_reply`.

## Secrets

Do not store credentials in this repository. Provide them as runtime environment
variables in the AI runner or secret manager.
