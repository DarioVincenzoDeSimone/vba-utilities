# vba-utilities

Utilities for VBA in Excel. Most of them are a kind of trascode of SQL sintax.

## select_personal

Use `select_personal` function to get the value of a cell in a table passing a specific column with some condition.

>SQL style: `SELECT what FROM table WHERE nameOfColumn1 = value1 AND nameOfColumn2 = value2`.

#### Input fields

`fieldToSearch `  name of column where is the value to get.

`sheet` name of sheet where is table

`firstFieldCondition` name of column of first condition

`firstCondition` value of first condition

`secondFieldCondition` name of column of second condition

`secondCondition` value of first condition

#### Output
Value of first cell.


## selectCount

Use `selectCount` function to count number of cells with a specific value in a specific column.
>SQL style: `SELECT COUNT(field) FROM table WHERE nameOfColumn1 = value1`.

#### Input fields

`countfield`  name of column where to count field.

`whereCondition`  the value to count.

`sheet` name of sheet where is table.

#### Output
the count.
