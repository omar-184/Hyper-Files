# Op documentation format (authoring guide)

The `docs/*.md` files are the reference for every canonical edit op the
executor registers. `op-docs.ts` parses them at import time into `OP_DOCS`;
`opUsage()` derives the one-line signature that guided errors append.
`tests/op-docs-sync.test.ts` asserts that the docs track the op registry
exactly and that the files contain no CJK text.

## File layout

```
# <Group title>
> <one-line summary of the group>

<optional group-level prose: addressing, units, conventions>

### <opName>
`<signature>`

<prose, field table, examples, common mistakes, related ops>

### <nextOpName> (internal)
`<signature>`
...
```

- One file per group; the file name is the group id (`text`, `element`,
  `insert`, `table`, `slide`, `deck`).
- A block starts at `### <opName>`; `opName` must match a registered op.
- The first non-empty line after the heading is the **signature**, wrapped in
  single backticks. It is emitted verbatim as `Usage: <opName> <signature>` in
  guided errors, so keep it compact and free of backticks.
- Heading suffixes in parentheses set flags: `(internal)` marks an op whose
  payload (bytes, clipboard data, part paths) only the app's own UI produces;
  `(pending)` hides its usage line until the registering branch lands.
  Combine as `(internal, pending)`.
- ```json fenced blocks are examples. Use the placeholder ids below.

  ```

## Placeholder ids for examples

| Placeholder  | Fixture element                 |
| ------------ | ------------------------------- |
| `e_TEXT`     | a text box on slide 0           |
| `e_SHAPE`    | a rounded rectangle on slide 0  |
| `e_PICTURE`  | a picture on slide 0            |
| `e_TABLE`    | a 2x2 table on slide 0          |
| `e_CHART`    | a bar chart on slide 0          |
| `e_LINE`     | a straight connector on slide 0 |
| `e_GROUP`    | a group on slide 0              |
| `e_CHILD`    | a direct child of that group    |
| `SECTION_ID` | an existing section GUID        |

The example deck has exactly two slides (indices 0 and 1; durable ids `s_1`
and `s_2`).

## Writing rules

- English only. Example text content (titles, labels, categories) is English too.
- Units are document-space EMU unless a field name says otherwise (`...Pt`,
  `...Pct`, `...Deg`). Say so in the field table.
- Field tables list every field the op's `validate` / `apply` reads, not only
  the ones in the signature.
- "Common mistakes" entries describe a frequent misuse and the corrective
  action, in the same voice as the guided errors.
- Keep each group file under roughly 8 KB; split a group into sub-files rather
  than exceeding that.
