---
layout: default
title: Merging Two Or More Tables
nav_order: 4
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
Open Excel spreadsheet with the tables you wish to merge.

1. Select [Data] from the large toolbar at the top.

2. Click on a cell in any of tables you wish to merge.

3. Select **From Table/Range** in the _Get & Transform_ section of the [Data] menu.
   ![Step 1-3](https://github.com/KevinSCLin/Microsoft-Excel-Useful-Procedures/blob/gh-pages/assets/images/Steps 3+4.PNG?raw=true)
   
   A popup called Power Query Editor will appear 

5. Click [File] from the large toolbar of the Power Query Editor.
   ![Step 5](https://github.com/KevinSCLin/Microsoft-Excel-Useful-Procedures/blob/gh-pages/assets/images/Step5.PNG?raw=true)

6. Select **Save & Load To** from the dropdown menu

   ![Step 6](https://github.com/KevinSCLin/Microsoft-Excel-Useful-Procedures/blob/gh-pages/assets/images/Step6.PNG?raw=true)

   A box with several save options will appear on your screen

7. Select the **Only Create Connection** option.

8. Click **OK**

   ![Step 7](https://github.com/KevinSCLin/Microsoft-Excel-Useful-Procedures/blob/gh-pages/assets/images/Step7.PNG?raw=true)

   Progress Check 1: A menu should pop out on the right side of your screen called Queries & Connections and you should see your table    there.
   
9. Repeat Steps 3 through 8 for all tables you wish to merge.

### Merging Connections into a Single Table

9. Select [Get Data] from the [Data] menu.

10. Select **Combine Queries** from the dropdown.

11. Click **Merge**.

  ![Steps 1,2,3](https://github.com/KevinSCLin/Microsoft-Excel-Useful-Procedures/blob/gh-pages/assets/images/Part2Steps1.PNG?raw=true)
12. Select the two tables you wish to merge from the dropdown bars.

  ![Step 4](https://github.com/KevinSCLin/Microsoft-Excel-Useful-Procedures/blob/gh-pages/assets/images/Part2Steps2.PNG?raw=true)
13. Click on the single column you wish to match the tables by on both tables.
14. Select **Left Join** from the **Join Kind** dropdown menu.

15. Click **OK**

   ![Steps 5,6,7](https://github.com/KevinSCLin/Microsoft-Excel-Useful-Procedures/blob/gh-pages/assets/images/Part2Steps3.PNG?raw=true)

   A Power Query editor should appear on your screen.
   
16. Click on the two arrows pointing in opposite directions above the table of labels.

   ![Steps 8](https://github.com/KevinSCLin/Microsoft-Excel-Useful-Procedures/blob/gh-pages/assets/images/Part2Steps4.PNG?raw=true)
   
   A Box of options should appear.
   
17. Unselect **Select all columns**.

18. Click the column you want to display.

19. Click **OK**

   ![Step 9](https://github.com/KevinSCLin/Microsoft-Excel-Useful-Procedures/blob/gh-pages/assets/images/Part2Steps5.PNG?raw=true)

20. Click [File] from the menu at the top.

21. Save in the format you wish.

## Well Done!
If you followed this guide you can now make queries for individual tables and merge them. This will allow you to quickly create powerful and informative tables while keeping your source information in separate tables.
 
