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

Supports the [documentation as of 2020-02-03](https://github.com/ministryofjustice/technical-risk-measures/tree/2ed7eb551c40365603a425e6cf11655ec175396b).

This version contains two breaking changes. Follow the migration instructions
for both of them in order before running the script.

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

### Merging 4 measures about ease of making changes into 2

Two columns are removed, and two others are replaced with two new ones, reducing
the total number of columns further to 27.

Delete **column R** ("Code base can be easily changed?") and **column J**
("Team can deploy in working hours?").

The affected criteria have changed enough that existing data from the other two
columns which will be replaced ("Team who owns the app own the complete
deployment?" and "Can deploy multiple times a day?) is unlikely to be accurate
for the new definitions, so it's best to delete the data from these two columns
(now **columns L and M**).

Then run the script.

## 1.0

Supports the [documentation as of 2019-07-03](https://github.com/ministryofjustice/technical-risk-measures/tree/37f6aa33969a28f8890c5404d3e29fbeca0f424a). One
non-breaking wording change made since then (renaming "atrophy" to "decay") is
also supported.

Does everything needed to set up a new spreadsheet from scratch.
