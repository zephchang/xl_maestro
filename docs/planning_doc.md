MAPPING OUT SPACE.

Ok so our goal is to create a meaning dictionary for all the tables in a given sheet.

STEP 1: Create a semantic map of just one table. Args: (cols_range, rows_range, cells_range, title_cell) (1) specify what are the anchor axes. For the collumn axis you may need to have an additional cell which contains the context data (it's a header). [done]

STEP 1B: During this proccess you need to create a list of all cells which need to be checked. [skip]

Note: need to decide how we are going to store cells. Probably as numbers directly or even coordinate pair makes the most sense (4,3) is D3. It needs to have the sheet name in it as well [done]

STEP 1c: test on kayla's shee

STEP 2o: construct dictionary objects for 2 tables in Kayla's sheet

STEP 2: Create a semantic map of all tables on a sheet (probably manually coded is realistic). Semantic maps should store meaning as (cols, rows, table title)

STEP 3: For each cell, we are going to do a check. Pull the formula. We need to deal with some specific cases:

> 1 Generic find cell. Semantic meaning dictionary needs to have sheet name, and then cell. If the cell is just 2 letters, need to add the current sheet name in front of it

> 2 XLOOKUP kind of thing, MATCH, INDEX< >
