---
layout: default
title: Glossary
nav_order: 5
---


# Merging two tables 
{: .no_toc }

This instruction will guide you on how to merge information scattered in two or more tables into a single table. 
{: .fs-6 .fw-300 }

## Table of contents
{: .no_toc .text-delta }

1. TOC
{:toc}

---


## Purpose of this instruction
When structuring information within tables it is necessary to keep related data together. Sometimes it is necessary to combine this information into a Power Query. This will provide an insightful summary of data scattered in several tables. This instruction set will follow the procedure using Power Query.  

## Instructions

### Saving a Table as a Connection
1. Open Excel spreadsheet with the tables you wish to merge.
2. Select [Data] from the large toolbar at the top.
![Step 2](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/Step 2.png?raw=true)
3. Click on a cell in any of tables you wish to merge.
4. Select **From Table/Range** in the _Get & Transform_ section of the [Data] menu.
A popup called Power Query Editor will appear 
5. Click [File] from the large toolbar of the Power Query Editor.
6. Select **Save & Load To** from the dropdown menu
A box with several save options will appear on your screen
7. Select the **Only Create Connection** option.
8. Click **OK**
Progress Check 1: A menu should pop out on the right side of your screen called Queries & Connections and you should see your table there.
9. Repeat Steps 3 through 8 for all tables you wish to merge.

### Merging Connections into a Single Table
1. Select [Get Data] from the [Data] menu.
2. Select **Combine Queries** from the dropdown.
3. Click **Merge**.
4. Select the two tables you wish to merge from the dropdown bars.
5. Click on the single column you wish to match the tables by on both tables.
6. Select **Left Join** from the **Join Kind** dropdown menu.
7. Click **OK**
A Power Query editor should appear on our screen.
8. Click on the two arrows pointing in opposite directions above the table of labels.
A Box of options should appear.
9. Unselect **Select all columns**.
10. Click the column you want to display.
11. Click **OK**
12. Click [File] from the menu at the top.
13. Save in the format you wish.
 
