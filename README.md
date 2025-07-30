### 2 languages found
[![en](https://img.shields.io/badge/lang-en-red.svg)](https://github.com/gbetorre/ExcelToSQL/blob/master/README.md)
[![it](https://img.shields.io/badge/lang-it-yellow.svg)](https://github.com/gbetorre/ExcelToSQL/blob/master/README.it.md)

---

![MIT license](https://img.shields.io/badge/license-MIT-blue)

This is a porting of a simple, but interesting, project originally written in Python 2.X by
[Jerry Fengwei Zhang](https://github.com/JerryFZhang), to make it compliant with Python 3.X


# ExcelToSQL

This is a program written in Python that takes an xls file as an input, produces csv code - as an intermediate format - for each sheet stored, and finally generates SQL insertion code for an SQL database.<br> Some of the cool features including

- checking if the file exist or not and then create or wipe the file if necessary;
- check the datatype and parse string and int differently.

# State of the art

Since people started using spreadsheets to store data, there has been a need to access these files to extract
information and turn it into structured data.

This need has generated a tremendous amount of tools that, over the years, have become more and more refined and
efficient.

Those who, like me, are quite old, might remember the bridges from Excel “to database worlds” that, in the years around
the end of the last millennium (alas!), presented as “revolutionary” solutions the transformation connectors from an
Excel spreadsheet to a Microsoft Access database (how usable the latter -given its aberrant SQL dialect- I couldn't
tell).

Nowadays there are countless Open Source macros, utilities, tools, software and repositories that often guarantee a
great result. ETLs have become progressively more user-friendly and can do the transformation work very well. For those
interested in code that can be integrated into other software, on GitHub alone, without even committing, the topic
[excel-to-sql](https://github.com/topics/excel-to-sql?o=desc&s=updated) presents at least 18 public repositories...

# The problem

Why, then, one should undertake to port a repository, whose syntax conforms to Python 2.X, by making it conform to
Python 3.X?

First of all, it should be considered that the operation of transforming Excel into a database can serve many purposes:

- there are those who want to have, by applying a script, a ready-made, self-contained database;
- those who want to have the tables corresponding to the generated spreadsheets in an existing database, whose access
  keys they provide;
- those who need to generate a schema or database on disk...<br> and so on.

In my case, it is enough for me that the script insisting on the Excel spreadsheet does not make the database, nor the
transformations, but that it produces only the queries, pure and simple, written in the purest and simplest ANSI SQL
standard.

# The solution

Database: a layer you must trust!!<br> If there is a fixed point, a cornerstone of an application, a rock, this is
represented by the database.<br> Therefore, I would find it incongruous to rely on tools and software that directly
generate the db: I like to send creation and populating scripts to the database only after having checked them or, at
least, having checked the process by which they were produced!

Perhaps, in this historical moment characterized by "vibe coding" and AI agents that seem destined to supplant us
programmers, this position is a bit anachronistic, a somewhat old fashioned way of proceeding, but it is nevertheless
the only mode I feel comfortable.

This is the bottom line here!<br> By the ExcelToSQL script, I can get exactly that; the script reads an Excel folder,
transforms each of its spreadsheet into a CSV file (Comma Separated Values, which are text files) then produces, for
each CSV file, a text file having:

- as the name the same name as the CSV file, and thus as the Excel sheet;
- as the extension .sql
- as content the text contained in the Excel sheet turned into input query.

# How to use

- There are three ways to download the file:
  1. Download the repo as a zip.
  2. Download the <code>run.py</code> file only.
  3. Or clone the repo by typing the following command into your terminal:

```
git clone https://github.com/gbetorre/ExcelToSQL
```

Install the require modules including:

- xlrd
- pandas

Put the spreadsheet file in the folder and name it as data.xls, make sure it ends with .xls not .xlsx.

- There are two ways to run the file:
  1. Run the file IDLE or other IDEs.
  2. Set the directory in the terminal and then type the following command into the terminal

```
python run.py
```

## Examples

Given the following data, originally stored in an Excel spreadsheet:

| CompanyID | CompanyName | CompanyIndustry |
| --------- | ----------- | --------------- |
| C001      | SDFESDF     | IT              |
| C002      | DAWR        | Electronics     |
| C003      | SDFD        | IT              |
| C004      | F           | IT              |
| C005      | DFEF        | IT              |

**_Table.1 - Example of data stored in the Excel sheet named “Company”_**

become:

```SQL
INSERT INTO  Company
VALUES ('C001', 'SDFESDF', 'IT');
INSERT INTO  Company
VALUES ('C002', 'DAWR', 'Electronics');
INSERT INTO  Company
VALUES ('C003', 'SDFD', 'IT');
INSERT INTO  Company
VALUES ('C004', 'F', 'IT');
INSERT INTO  Company
VALUES ('C005', 'DFEF', 'IT');
```

**_List.1 - SQL code generated by the script from the data in the Excel_**

Note that, in this example, the input queries generated contain only values of type String. In fact, each value is
enclosed in single quotes.

This is not a fixed behavior; to check this, simply add an ID field and enter numbers in it. Nor is it necessary for
Excel column cells to be formatted as type “Numeric”. It is sufficient that, in the field, numbers are present. These
will become floating-point numbers in the CSV file:

| ID  | CompanyID | CompanyName | CompanyIndustry |
| --- | --------- | ----------- | --------------- |
| 1.0 | C001      | SDFESDF     | IT              |
| 2.0 | C002      | DAWR        | Electronics     |
| 3.0 | C003      | SDFD        | IT              |
| 4.0 | C004      | F           | IT              |
| 5.0 | C005      | DFEF        | IT              |

**_Table.2 - Data stored in the intermediate CSV format with the addition of the ID field_**

Finally, they will be treated as data of type INTEGER in query generation:

```SQL
INSERT INTO  Company
VALUES (1, 'C001', 'SDFESDF', 'IT');
INSERT INTO  Company
VALUES (2, 'C002', 'DAWR', 'Electronics');
INSERT INTO  Company
VALUES (3, 'C003', 'SDFD', 'IT');
INSERT INTO  Company
VALUES (4, 'C004', 'F', 'IT');
INSERT INTO  Company
VALUES (5, 'C005', 'DFEF', 'IT');
```

**_List.2 - SQL code generated, having automatically numeric values_**

What happens, now, if various type of values are present in a same field of a spreadsheet? The script will try to detect
whether the values are numeric or string and generate a formally correct but, likely, conceptually incongruent SQL.

For instance, the following values in the Excel sheet:

| ID       | CompanyID | CompanyName | CompanyIndustry |
| -------- | --------- | ----------- | --------------- |
| 1        | C001      | SDFESDF     | IT              |
| 2.02     | C002      | DAWR        | Electronics     |
| foo'     | C003      | SDFD        | IT              |
| mosquito | C004      | F           | IT              |
| 5        | C005      | DFEF        | IT              |

**_Table.3 - The data in the ID field of the Excel sheet are heterogeneous in type._**

will generate the following SQL code:

```SQL
INSERT INTO  Company
VALUES (1, 'C001', 'SDFESDF', 'IT');
INSERT INTO  Company
VALUES (2, 'C002', 'DAWR', 'Electronics');
INSERT INTO  Company
VALUES ('foo’', 'C003', 'SDFD', 'IT');
INSERT INTO  Company
VALUES ('mosquito', 'C004', 'F', 'IT');
INSERT INTO  Company
VALUES (5, 'C005', 'DFEF', 'IT');
```

**_List.3 - The generated SQL code does not make sense according to a traditional database perspective, but the script
respects the type it finds in the Excel._**


# License 

MIT, see LICENSE.

# Authors

* [zhang96](https://github.com/JerryFZhang) (original creator) 
* [gbetorre](https://github.com/gbetorre) (porting to CPython 3.X)


