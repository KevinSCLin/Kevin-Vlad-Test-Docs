---
layout: default
title: Building a Gantt chart in Excel
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

    ![add button](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/select_data_source_add_button.PNG?raw=true)
    
    The _Select Data Source_ transforms to _Edit Series_ window.
    
    Click on the Up Arrow adjacent to _Series name_ text box to add a title for the starting dates.
    
    The _Edit Series_ box transforms to show only a text box for select or entering the range of the title.
    
    ![series name up arrow](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_series_name_up_button.png?raw=true)

3. Click on the title for starting date.

    ![series name select_title](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_select_title.png?raw=true)
   
   Next, click the Down Arrow adjacent to _Series name_ textbox to complete selection of title.
   
   The _Edit Series_ box transforms back to its original form.
   
4. Click on the Up Arrow adjacent to _Series values_ text box allow selection of task names in the chart.

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

3. Click on the Up Arrow adjacent to _Series values_ text box to allow selection of task durations in the chart.
   
   The _Edit Series_ box transforms to show only a text box for select or entering the range of cells storing the task durations.
   
3. Left-click and drag to select the cells containing the task names. Click the Down Arrow to complete selection.

    ![series name select_title](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/edit_series_select_values2.png?raw=true)     
   
      The _Edit Series_ box transforms back to its original form.
      Both text boxes should contain the cell ranges containing the required information.
      
5. Click **OK** to save the data selected in the above steps.

---

## Modifying graph layout

The graph produced from the steps above is not a Gantt chart.
We need to modify the bars, axes range, axes labels, and title to convert it into a Gantt chart.

![initial chart](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/initial_chart.png?raw=true)

### Remove starting date as independent bars

1. Right-click on the chart and left-click on **Format Chart Area**.

    A _Format Chart Area_ pane appears on the right side of the Excel application window.

2. Left-click on any of the blue bars to activate _Format Data Series_. Click on **No fill** for both _Fill_ and _Color_.
    
    The blue bars become invisible in the graph.
    
    ![format data series](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/format_data_series_no_fill.png?raw=true)

3. Invert the y-axis order first by clicking on any task name on the y-axis to activate _Format Axis_ window.
    Click on the bar chart icon to switch to _Axis Options_. Check the box labeled **Categories in reverse order**.
    
    ![format data series](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/format_axis_reverse_order.png?raw=true)

4. Edit the x-axis first by click on any date label to activate _Format Axis_. Click on the bar chart icon same as the
    image to open _Axis Options_. Under _Bounds_ incrementally add the Minimum and subtract the Maximum values until the bars
    in the graph achieve a better use of plot area.
    
    ![Note][NOTE]NOTE: Excel's date format is an integer starting at 1, and beginning from January 1, 1900.
    April 1, 2020's equivalent integer date value in Excel is 43922, meaning 43921 days have passed since
    January 1, 1900.
    
    ![format data series](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/format_axis_bounds.png?raw=true)

5. Modify the number format by expanding the _Number_ tab at the bottom of _Format Axis_ and click on the **Type** dropdown.
   Select a format to your liking and save the file. 

6. Add a chart title by clicking on the graph to trigger the **Chart Element** button.
    Edit the text in the chart title.
    
    You now have a completed Gantt chart.
    
    ![gantt chart finished](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/gantt_chart_finished.png?raw=true)

[NOTE]: https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/note_icon.png?raw=true

## Benefits of a Gantt chart

A Gantt chart provides a fast visual breakdown of tasks to do in series, which helps the viewer to reorder 
and prioritize tasks.