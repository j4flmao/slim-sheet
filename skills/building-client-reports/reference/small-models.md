# Prompts for small local models (7–14B)

Small models do the task when they only have to substitute names into a fixed call. Use the
`lite` tool tier (8 tools), temperature 0, and one tool call per message.

## Recipe

1. Ask the model for `overview(path)` only. Read the column names yourself or let it list them.
2. Paste one role preset from `roles/` and ask it to replace column names, nothing else.
3. Run `anomalies(scope='client')` and paste the findings back with "fix only these".

## Prompt (copy as is)

```text
You can call slim-sheet tools. Call exactly one tool, then stop.
Call build_report with these arguments. Replace each <...> with a column name from this list and
change nothing else:
COLUMNS: <paste the column names from overview>
ARGUMENTS:
path='<file>', data_sheet='<sheet>', title='<one sentence with a number>',
kpis=['Revenue=sum:<revenue column>|vs:<prior revenue column>:prior',
      'Orders=count:<id column>'],
charts=['trend:<date or month column>:<revenue column>',
        'rank:<category column>:<revenue column>']
```

## Why this works

- The grammar is closed: aggregates, intents and formats are fixed words, so the model cannot
  invent options.
- Errors teach: a wrong column name returns the list of valid names in the hint; paste it back.
- Everything visual is decided by the server, so a small model cannot make the report ugly.
