---
layout: default
title: Build a Gantt chart in Excel
nav_order: 2
---

# Gantt chart in Excel
{: .no_toc }

This instruction will guide you to build a Gantt chart in MS Excel using existing data
{: .fs-6 .fw-300 }

## Table of contents
{: .no_toc .text-delta }

1. TOC
{:toc}

---

## What is a Gantt chart
A Gantt chart is a staggered-layered horizontal bar chart.
Its serves as a visualization of a horizontal timeline demonstrating the start and end dates of multiple tasks.
A Gantt chart provides a viewer a quick, generalized breakdown of time-based weights and sequences of tasks.

## Getting Started
You will need the tasks and their start date or duration entered in a tabular form.
Here is an example.

![Gantt Chart sample table](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/gantt_chart_sample_tasks.png?raw=true)

---
## Insert Graph and Add Data

### Adding a graph to Excel spreadsheet
1. Go to [Insert] > [Charts] > [2-D bar] > **Stacked Bar**.
    ![2D stacked bar button](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/2D_stacked_bar.png?raw=true)
    
    A blank graph appears.
    
### Adding starting dates to all tasks 
2. Right-click on the graph and click **Select Data**.

    The _Select Data Source_ window appears.

3. Click the **Add** button.

    ![series name select_title](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/select_data_source_add_button.png?raw=true)

    Click on the Up Arrow adjacent to _Series name_ textbox to add a title for the tasks.
    
    The _Edit Series_ box transforms to show only a text box for select or entering the range of the title.
    
    ![series name up arrow](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_series_name_up_button.png?raw=true)

4. Click on the title for tasks.

    ![series name select_title](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_select_title.png?raw=true)
   
   Next, click the Down Arrow adjacent to _Series name_ textbox to complete selection of title.
   
   The _Edit Series_ box transforms back to its original form.
   
5. Click on the Up Arrow adjacent to _Series values_ text box to include the task names in the chart.

   The _Edit Series_ box transforms to show only a text box for select or entering the range of cells storing the task starting dates.
   
   ![series values up arrow](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_series_values_up_button.png?raw=true)
   
   Left-click and select the cells containing the task names. Click the Down Arrow to complete selection.
   
   ![series values up arrow](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_select_values.png?raw=true)
   
   The _Edit Series_ box transforms back to its original form.
   
6. Click **OK** to save the data selected in the above steps.