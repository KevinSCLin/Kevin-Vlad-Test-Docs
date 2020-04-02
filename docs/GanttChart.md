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
    
### Adding starting dates to tasks 
1. Right-click on the graph and click **Select Data**.

    The _Select Data Source_ window appears.

2. Click the **Add** button.

    ![add button](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/select_data_source_add_button.png?raw=true)
    
    The _Select Data Source_ transforms to _Edit Series_ window.
    
    Click on the Up Arrow adjacent to _Series name_ text box to add a title for the starting dates.
    
    The _Edit Series_ box transforms to show only a text box for select or entering the range of the title.
    
    ![series name up arrow](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_series_name_up_button.png?raw=true)

3. Click on the title for starting date.

    ![series name select_title](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_select_title.png?raw=true)
   
   Next, click the Down Arrow adjacent to _Series name_ textbox to complete selection of title.
   
   The _Edit Series_ box transforms back to its original form.
   
4. Click on the Up Arrow adjacent to _Series values_ text box to include the task names in the chart.

   ![series values up arrow](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_series_values_up_button.PNG?raw=true)

   The _Edit Series_ box transforms to show only a text box for select or entering the range of cells storing the task starting dates.
      
   Left-click and drag to select the cells containing the task start dates. Click the Down Arrow to complete selection.
   
   ![series values up arrow](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_select_values.png?raw=true)
   
   The _Edit Series_ box transforms back to its original form.
   Both text boxes should contain the cell ranges containing the required information.
   
5. Click **OK** to save the data selected in the above steps.

### Adding names to tasks

1. Click on the **Edit** button in _Select Data Source_ window.

    ![edit button](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/select_data_source_edit_button.png?raw=true)
    
    The _Select Data Source_ window transforms into _Axis Labels_ window.
    
    ![Note][NOTE] NOTE: Even though the **Edit** button's caption says _Horizontal (Category) Axis Labels_, 
    the bar chart is horizontal, effectively the conventional x and y axes placement swapped.
    
2. Click on the Up Arrow button adjacent to the _Axis label range_ text box.
    
    ![axis labels up button](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/axis_labels_up_button.PNG?raw=true)
    
3. Left-click and drag to select the cells containing the task names.

    ![select title range](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/axis_labels_select_range.png?raw=true)

    The _Axis Labels_ text box now contains the range storing the names of tasks.
    
    Click on the Down Arrow to return to the initial _Axis Labels_ window.
    
    ![axis labels up button](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/axis_labels_textbox_filled.png?raw=true)

### Adding duration to tasks

The following steps repeat those in [Adding starting dates to tasks](#adding-starting-dates-to-tasks) above,
with the exception of selecting different ranges of cells. Skip to [Modifying graph layout](#modifying-graph-layout)
if you understand how to add new data to a graph.

1. Click the **Add** button in the _Select Data Source_ window.

    ![add button2](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/select_data_source_add_button_2.PNG?raw=true)

    The _Select Data Source_ window transforms into _Axis Labels_ window.
    
    Click on the Up Arrow adjacent to _Series name_ text box to add a title for the starting dates.
        
    The _Edit Series_ box transforms to show only a text box for select or entering the range of the title.
        
    ![series name up arrow](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_series_name_up_button.png?raw=true)
    
2. Click on the title for task duration.

    ![series name select_title](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_select_title2.png?raw=true)

   Next, click the Down Arrow adjacent to _Series name_ textbox to complete selection of title.
   
   The _Edit Series_ box transforms back to its original form.

3. Click on the Up Arrow adjacent to _Series values_ text box to include the task names in the chart.
   
      The _Edit Series_ box transforms to show only a text box for select or entering the range of cells storing the task durations.

## Modifying graph layout