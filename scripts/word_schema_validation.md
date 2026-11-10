# Word schema validation

This checkout-only tool uses OOXML-Validator v2.1.6 (MIT), targeting
`Microsoft365` by default. It is neither a product nor a test dependency.
Only CI downloads the executable. Release URLs and SHA-256 values from the
first downloads on 2026-11-10 are pinned in `word_schema_validator_assets.json`.
Verify the archive before extracting or running its executable:

```sh
uv run --no-sync python -m scripts.validate_word_schema verify-asset /absolute/path/linux-x64.zip
uv run --no-sync python -m scripts.validate_word_schema validate /absolute/path/OOXMLValidatorCLI --output-dir /new/output/directory --report /output/word-schema-report.json
```

The script also supports direct invocation as `python scripts/validate_word_schema.py`.
`--office-version` changes the validation target; `--exceptions` selects an
exception file. The output directory must not already exist. The shared
`scripts.sample_corpus.sample_cases` definition discovers every `*.md` recursively
under the three sample directories. `render_corpus` preserves each source's
settings, forces strict mode, and adds a separate charts-on run. Its immutable
cases expose the source, render options, and output-relative path for reuse by
snapshot tests. Even if the charts sample already enables charts in frontmatter,
both requested cases are retained.

The script fails for uncovered errors, unused exceptions, malformed/incomplete
validator results, or execution/rendering failures. Validator exit code zero
does not mean validation passed. JSON records each document's errors (id,
part URI, XPath, description, error type, and exception coverage) and the decision.
It requires results for every rendered document using `--all --recursive`.

The Ubuntu `word-schema` job is deliberately report-only: job-level
`continue-on-error: true` keeps its failures from failing the workflow, and neither
job has a `needs` dependency on the other. The script retains failure exit codes.
An `always()` step supplies a failure JSON if setup prevented the script from
running; another `always()` step uploads the report. Removing job-level
`continue-on-error` is reserved for the later blocking-gate task.

## Exception format

`word_schema_exceptions.json` starts as an empty JSON list. Each entry must have
this shape (the values below are placeholders, not an accepted exception):

```json
{
  "id": "validator error Id",
  "part_uri": "/word/example.xml",
  "xpath": "/w:example[1]",
  "reason": "Why this Word-saved reproduction demonstrates a false positive",
  "evidence": {
    "kind": "word_saved_reproduction",
    "procedure": "Exact steps to create the same construction and save it in Word, then run the validator with the stated Office target",
    "word_version": "Exact Word version and build",
    "validator_output": [
      {
        "FilePath": "example-word-saved.docx",
        "ValidationErrors": [
          {
            "Id": "validator error Id",
            "Description": "Observed validator description",
            "ErrorType": "Schema",
            "Path": {
              "PartUri": "/word/example.xml",
              "XPath": "/w:example[1]"
            }
          }
        ]
      }
    ]
  }
}
```

Only a document actually saved by Word reproducing the same error is acceptable
evidence. Standards citations, assumptions, and engine-produced documents alone
are insufficient. Preserve the observed validator output (string-encoded or list
`ValidationErrors` are both accepted). The script checks the required fields and
that the evidence contains the exact identity; reviewers must verify the
reproduction itself. Matching uses the exact `(id, part_uri, xpath)` tuple across
the whole corpus, with no wildcards. Duplicate identities and entries matching no
current error fail the run. Never commit the Word reproduction document, which
can contain author information. Do not exempt real engine defects, including
the currently known `updateFields` ordering error in `/word/settings.xml`.
