# Keeping each AP's box on one page

This explains how the generator stops an AP's box from being cut in half by a page break,
and why it does **not** need to break the table into a separate table per page.

Written for someone who has never used Microsoft Word. Nothing here needs to be done by hand —
the program already does all of it. This is just what it is doing and why.

---

## 1. Some words first

**Page break.** A Word document is one long ribbon of content. Word decides where each printed
page ends and the next begins. That invisible boundary is called a *page break*. When the content
reaches the bottom of a page, whatever comes next is pushed onto a new page.

**Table.** A grid of boxes. Ours is 4 columns wide and very long.

**Row.** One horizontal line of boxes in that grid.

**AP block.** In section 3.2 (Signal Coverage Test), each access point takes exactly **5 rows**:

```
Row 1   AP ID              <- one AP's block
Row 2   Test Device 1
Row 3   SSID – Band / BSSID | Expected | Recorded | Result
Row 4   Wireless@SGx – 2.4 GHz
Row 5   Wireless@SGx – 5 GHz
------------------------------------------- (next AP starts here)
Row 6   AP ID              <- next AP's block
Row 7   Test Device 1
...
```

The program repeats those 5 rows once per AP, so 12 APs means 60 rows in one long table.

---

## 2. What was going wrong

Word was free to end a page **anywhere between two rows**. It fills a page to the bottom, and
whatever no longer fits drops to the next page — it does not know or care that our rows come in
groups of five.

So on a full page you would get this:

```
        ...bottom of page 4...
Row 51  AP ID                     <- AP 11 starts
Row 52  Test Device 1
======= page break =======
        ...top of page 5...
Row 53  SSID – Band / BSSID       <- the rest of AP 11
Row 54  Wireless@SGx – 2.4 GHz
Row 55  Wireless@SGx – 5 GHz
```

AP 11's box is torn in two. The heading and the device line sit alone at the bottom of page 4,
and the measurements that belong to them are on page 5.

---

## 3. The fix, in one sentence

Word has a setting that means **"do not let a page break happen right after this row — drag the
next row along with it."** The program switches that setting on for the first four rows of every
AP block, and deliberately leaves it **off** on the fifth.

That single difference is what does all the work.

---

## 4. Why leaving it off on row 5 matters

Think of the setting as glue between one row and the next.

```
Row 1  AP ID                 glued to next
Row 2  Test Device 1         glued to next
Row 3  SSID – Band           glued to next
Row 4  2.4 GHz               glued to next
Row 5  5 GHz                 NOT glued   <- a page may end here
------------------------------------------
Row 6  AP ID                 glued to next     (next AP, same pattern)
```

Rows 1 to 5 are chained into one unbreakable lump. Word will not split a lump, so if the whole
5-row block does not fit in the space left on the page, Word gives up on that space and moves the
**entire block** to the top of the next page. The block always arrives intact.

Row 5 has no glue, so a page is still *allowed* to end there. That gap between one AP and the next
is the only place a page can break — which is exactly where we want it.

If row 5 were glued too, every row in the table would be chained to every other row, the whole
table would become one gigantic lump that cannot fit on any page, and Word would be forced to
ignore the instruction and break wherever it liked. **The gap is the point.** A rule with no
escape hatch is the same as no rule at all.

There is a second, smaller setting applied to every row as well: *do not let a single row break
across pages*. That covers the rare case of one row being so tall (a very long address, say) that
Word would try to split that one row down the middle.

---

## 5. Why the table is not split into separate tables

The "UAT Field Template (Full)" document solves this a different way: it uses a fresh table for
each page, with 5 APs per table.

That works, but it hard-codes an assumption — that exactly 5 AP blocks fit on a page. That is
only true while nothing else changes. A longer venue address, a bigger font, a wider margin, an
extra line anywhere above section 3.2, or one AP block growing by a row, and 5 no longer fit.
The tables would then break in the wrong places, and someone would have to go back and re-do the
grouping by hand.

The glue method makes no assumption at all. It never states how many APs go on a page. It only
states that **an AP block must not be cut**, and lets Word work out how many fit each time it lays
the document out. If the template changes tomorrow and only 4 blocks fit, the document is still
correct — Word simply puts 4 on that page. Nothing needs re-doing.

It also keeps section 3.2 as **one continuous table**, which is what the section is meant to be.

---

## 6. Where this lives in the code

In `uat_document.py`:

```python
def keep_ap_blocks_together(table):
    """Hold each AP's rows on one page, so a whole block moves down rather than being split."""
    for index, row in enumerate(table.rows):
        set_cant_split(row)
        is_last_row_of_block = (index + 1) % ROWS_PER_AP == 0
        set_row_keep_with_next(row, not is_last_row_of_block)
```

Reading it in plain English: walk down every row of the table. Tell each row not to break in half.
Then work out whether this row is the last of its group of five (`ROWS_PER_AP` is 5) — if it is,
leave the glue off; otherwise, glue it to the row below.

It runs at the "Setting the page breaks..." stage, after all the AP blocks have been added.

The two settings it uses are the same ones a person can set by hand in Word:

- Glue: select the rows, then **Home → the small arrow at the corner of the Paragraph group →
  Line and Page Breaks tab → tick "Keep with next"**.
- No splitting a row: select the table, then **Table Layout → Properties → Row tab → untick
  "Allow row to break across pages"**.

The program writes those same settings into the file directly, so the result is identical to
having done it by hand — just done perfectly every time, for hundreds of rows.

---

## 7. How this was checked

Documents were generated for 1, 7, 11, 12, 17, 23, 40 and 97 APs. Each one was opened in Word
through automation, and the real page number of every single row was read back from Word's own
layout engine, then compared block by block.

- Before the fix, a 23-AP document split APs 11, 16 and 21 across pages.
- After the fix, every AP block in every one of those documents sits on a single page.
- The 23-AP document is still 9 pages long, so nothing is wasted — the blocks that move down fill
  space that the following blocks would have used anyway.
