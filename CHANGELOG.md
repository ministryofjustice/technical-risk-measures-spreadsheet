# Changelog

Breaking changes which require manual edits to an existing spreadsheet before
running the new version of the script will be recorded here, with instructions
for how to make those edits.

Breaking changes are those which add, remove or reorder columns or rows.
Updates to the wording or descriptions of criteria are not considered to be
breaking changes for the purposes of this script, although they may mean you
need to review your existing data to see whether it needs updating for the new
definitions.

All changes will be documented in commit messages, so look there for detail on
other changes which are not included here.

## 2.0

### Merging general dependencies and security patches into one measure

These two adjacent columns are replaced with one, reducing the total number of
columns to 29.

If you have data in either of these two columns that you want to use for the
replacement measure ("When was the oldest unapplied version of any dependency
released?"), move it into **column X** ("When were general dependency updates
last applied?"). This may mean that you want to keep the older of the two dates
in each row, or use some other strategy to decide what to keep. When you've done
this, delete **column Y** ("When were the oldest unapplied security patches
released?").

Or if you don't want to keep any data from these columns, just delete column Y.

Then run the script.

## 1.0

Supports the [documentation as of 2019-07-03](https://github.com/ministryofjustice/technical-risk-measures/tree/37f6aa33969a28f8890c5404d3e29fbeca0f424a). One
non-breaking wording change made since then (renaming "atrophy" to "decay") is
also supported.

Does everything needed to set up a new spreadsheet from scratch.
