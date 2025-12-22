import os
import numpy as np
import pandas as pd
import textwrap as tw
import matplotlib as mpl
import win32com.client as win32
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from itertools import groupby
from datetime import datetime
from operator import itemgetter
from sqlalchemy import create_engine
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
mpl.rcParams.update({'font.size': 18})

############################
#Plots run and save, can create pdf, just need to sort email and HP and SP to review.
#Then need to work out power automate.
##########################SEE NOTE AT END



base_path = "G:/PerfInfo/Performance Management/OR Team/Emily Projects/General Analysis/SPC Charts/"
symbol_path = "G:/PerfInfo/Performance Management/OR Team/Emily Projects/General Analysis/SPC Charts/Symbols/"
charts_path = "G:/PerfInfo/Performance Management/OR Team/Emily Projects/General Analysis/SPC Charts/SPC Charts/"
pdfs_path = "G:/PerfInfo/Performance Management/OR Team/Emily Projects/General Analysis/SPC Charts/pdfs/"

#########################Read in data
#create connection
sdmart_engine = create_engine('mssql+pyodbc://@SDMartDataLive2/InfoDB?'\
                           'trusted_connection=yes&driver=ODBC+Driver+17'\
                               '+for+SQL+Server')

data_sql = """SELECT [WeekEndDate],
                     [metric_id],
                     [metric_desc],
                     [data],
                     [GoodDirection]
                     FROM [InfoDB].[dbo].[ceo_flashcard]
                     ORDER BY metric_id, WeekEndDate"""

full_data = pd.read_sql(data_sql, sdmart_engine)

sdmart_engine.dispose()


############################SPC Chart function
#Function defining SPC chart without using a window. Using moving range for 
#lower plot
def spcChartIndiv(data, date, title_text, good_dir = 'down', target = None):
    #####Initial calculations
    #Find the mean and SD of this data
    data_mean = np.nanmean(data)
    #Use ddof=1 so that N-1 is used as the denominator
    data_SD = np.nanstd(data, ddof = 1)
    #Also find the moving range (max-min for each period). Have to add two as the
    #second value in square brackets is not included.
    mov_R = [np.ptp(data[i:i+2]) for i in range(len(data) - 1)]
    mov_R = np.array(mov_R, dtype = float)
    mov_R_mean = mov_R.mean()
    #Find the action lines and warning lines for the mean plot
    u_mean_al =  data_mean + 2.66 * mov_R_mean
    l_mean_al = max(data_mean - 2.66 * mov_R_mean, 0)
    u_mean_wl =  data_mean + (2/3)*2.66 * mov_R_mean
    l_mean_wl = max(data_mean - (2/3)*2.66 * mov_R_mean, 0)
    
    ######Set up Plot
    fig, ax1 = plt.subplots(1,1, figsize=(15, 4))
    ax1.set_title('\n'.join(tw.wrap(title_text, 50)))
    #Set up the colours for each direction
    if good_dir == 'Down' or good_dir == 'Neutral':
        c_between_u = 'darkorange'
        c_between_l = 'blue'#'limegreen'
        c_above_3s = 'darkorange'#'red'
        c_below_3s = 'blue'#'lime'
        c_run_above = 'darkorange'#'gold'
        c_run_below = 'blue'#'forestgreen'
    elif good_dir == 'Up':
        c_between_u = 'blue'#'limegreen'
        c_between_l = 'darkorange'
        c_above_3s = 'blue'#'lime'
        c_below_3s = 'darkorange'#'red'
        c_run_above = 'blue'#'forestgreen'
        c_run_below = 'darkorange'#'gold'

    ######Plot the data
    #Plot the data as small points
    ax1.plot(date, data, 'o-', color = 'slategrey', markersize = 4, zorder = 10)
    #Add mean and upper and lower lines
    ax1.plot(date, data_mean*np.ones(len(data)),'k--')
    ax1.plot(date, u_mean_al*np.ones(len(data)), 'r:')
    ax1.plot(date, l_mean_al*np.ones(len(data)), 'r:')

    #######Set xlabel and x ticks
    ax1.set_xlabel('Date')
    #Find the optimal interval for x-axis labels, we want roughly 5 ticks
    num_months = (date[-1] - date[0]).days / 30
    intvl = int(num_months // 5)
    #If this gives a number less than one, then need to look at days on the 
    #x-axis instead
    if intvl <= 1:
        if num_months <= 4:
            intvl = 1
        else:
            intvl = 2
        ax1.xaxis.set_major_locator(mdates.MonthLocator(bymonthday = 1,
                                                        interval = intvl))
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%y'))
        for label in ax1.get_xticklabels(which='major'):
            label.set(rotation=0, horizontalalignment='center')
    else:
        ax1.xaxis.set_major_locator(mdates.MonthLocator(bymonthday = 1,
                                                        interval = intvl))
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%y'))
        for label in ax1.get_xticklabels(which='major'):
            label.set(rotation=0, horizontalalignment='center')

    ######Implement rules
    #Rules used here - out of control if:
    #    any point is outside of either 3sigma line,
    #    or, (N-1) points consecutively between 2sigma and 3sigma line 
    #       (on the same side of the mean line) '
    #    or, 2.5N consecutive points all rising or falling ,
    #    or, if any point is outside lines on SD plot
    above_3s, = np.where((data >= u_mean_al))
    below_3s, = np.where((data <= l_mean_al))
    between_limit = 5
    between_lines_upper, =  np.where((data <= u_mean_al) 
                                   & (data >= u_mean_wl))
    between_lines_lower, =  np.where((data >= l_mean_al)
                                   & (data <= l_mean_wl))
    run_limit = 8
    run_above, = np.where(data > data_mean)
    run_below, = np.where(data < data_mean)
    #Search for consecutive points in between_lines and run_above/below
    def consecutiveNums(dataset, min_run):
        output = []
        for k, g in groupby(enumerate(dataset), lambda ix : ix[0] - ix[1]):
            row = list(map(itemgetter(1), g))
            if len(row) >= min_run: output.append(row)
        return output
    run_above_points = consecutiveNums(run_above, run_limit)
    run_below_points = consecutiveNums(run_below, run_limit)
    between_upper_points = consecutiveNums(between_lines_upper, between_limit)
    between_lower_points = consecutiveNums(between_lines_lower, between_limit)

    ######Stability test
    #Test whether the most recent 50 data points are consistent with a system
    #in control
    test_start = len(data) - 50
    if (len(data) < 50) and (len(data) > 29):
        test_start = 0
    stable = True
    #First, look at 3 sigma
    if (any(x > test_start for x in above_3s)
        or any(x > test_start for x in below_3s)):
            stable = False
    #Looking at the maximum run for each of the remaining rules. If the last 
    #run is not in the last 50 points, then no runs will be
    else:
        breaches = 0
        #filter(None,...) removes the empty values from the list
        for run_data in filter(None, [run_above_points, run_below_points, 
                                      between_upper_points, between_lower_points]):
            if any(x > test_start for x in max(run_data)):
                breaches += 1
        #If there are any breaches, print unsuitability warning
        if breaches > 0:
            stable = False
    
    #Add a note to the chart on stability of the system
    stable_note = 'This process is stable' if stable else 'This process is not stable'
    ax1.text(0.5, -0.3, stable_note, ha='center', fontsize=14, transform=ax1.transAxes)
    
    ######Symbols
    #To find which symbol to use to describe the recent changes
    #Assume no changes and re-assign if changes found
    state = 'no_change'
    symbol_test_start = len(data) - 7
    if ((any(x > symbol_test_start for x in above_3s))
        and (any(x > symbol_test_start for x in below_3s))):
        state = 'warning'
    elif (any(x > symbol_test_start for x in above_3s)):
        state = 'Up'
    elif (any(x > symbol_test_start for x in below_3s)):
        state = 'Down'
    else:
        for run_data in filter(None, [run_above_points,between_upper_points]):
            if (any(x>symbol_test_start for x in max(run_data))):
                state = 'Up'
            else:
                for run_data in filter(None, [run_below_points,between_lower_points]):
                    if (any(x>symbol_test_start for x in max(run_data))):
                        state = 'Down'
                    
    #Then, find if this is a good or bad change and assign the symbol accordingly
    if state == good_dir:
        if good_dir == 'Up':
            symbol_file = "ImprovementUpwardsSymbol.png"
        elif good_dir == 'Down':
            symbol_file = "ImprovementDownwardsSymbol.png"
    elif state == 'no_change':
        symbol_file = "NoChangeSymbol.png"
    else:
        symbol_file = "ConcerningChangeSymbol.png"

    with mpl.cbook.get_sample_data(symbol_path + symbol_file) as file:
        symbol = plt.imread(file, format='png')
    
    #Add in the symbol defined above on an inset axis
    axin = ax1.inset_axes([0.97, 0.35, 0.2, 0.2], zorder = 15)
    axin.imshow(symbol)
    axin.axis('off')

    ######Plot extrenuous points in a different colour
    #Add these points to the plot
    #Loop through the size of each of the lists and plot the points on the graph
    for i in range(len(run_above_points)):
        #Add the window to the x-values so that the points are in the right place
        ax1.plot([date[x] for x in run_above_points[i]],
                   data[run_above_points[i]],
                   color = c_run_above, marker = 'o', markersize = 4,
                   zorder = 11, linestyle = 'None')

    #Do the same for the run below
    for i in range(len(run_below_points)):
        #Add the window to the x-values so that the points are in the right place
        ax1.plot([date[x] for x in run_below_points[i]],
                   data[run_below_points[i]],
                   color = c_run_below, marker = 'o', markersize = 4,
                   zorder = 11, linestyle = 'None')

    #Repeat for between points
    #Loop through the size of each of the lists and plot the points on the graph
    for i in range(len(between_upper_points)):
        #Add the window to the x-values so that the points are in the right place
        ax1.plot([date[x] for x in between_upper_points[i]],
                   data[between_upper_points[i]],
                   color = c_between_u, marker = 'o', markersize = 4,
                   zorder = 12, linestyle = 'None')

    #Do the same for the run below
    for i in range(len(between_lower_points)):
        #Add the window to the x-values so that the points are in the right place
        ax1.plot([date[x] for x in between_lower_points[i]],
                   data[between_lower_points[i]],
                   color = c_between_l, marker = 'o', markersize = 4,
                   zorder = 12, linestyle = 'None')
        
    #Finally, add in the points above or below 3 sigma lines
    for i in range(len(above_3s)):
        ax1.plot(date[above_3s[i]],
                   data[above_3s[i]],
                   color = c_above_3s, marker = 'o', markersize = 4,
                   zorder = 13, linestyle = 'None')
        
    for i in range(len(below_3s)):
        ax1.plot(date[below_3s[i]],
                   data[below_3s[i]],
                   color = c_below_3s, marker = 'o', markersize = 4,
                   zorder = 13, linestyle = 'None')
    
    ######Target Value
    #Work out process capability if a target was specified and the process is stable
    #Following formula in https://www.itl.nist.gov/div898/handbook/pmc/section1/pmc16.htm
    #for one-sided specifications and altering for which way is good
    if (target is not None) & (stable):
        if good_dir == 'down':
            cp = (target - data_mean) / (3* data_SD)
        elif good_dir == 'up':
            cp = (data_mean - target) / (3* data_SD)
                
    #######Save the figure
    fig.savefig(charts_path + title_text + ".png", format='png', bbox_inches="tight")
    plt.close()


######Loop over each metric and create it's SPC Chart
#List of unique metrics
metrics = (full_data[['metric_id', 'metric_desc', 'GoodDirection']]
           .drop_duplicates().sort_values(by='metric_id').values.tolist())

for id, metric, good_dir in metrics:
    #Filter data to that metric, ensure it's sorted
    sub_data = full_data.loc[full_data['metric_id'] == id].copy()
    #Create file title (replacing invalid symbols)
    title = f'{id} {metric.replace('<', 'lt').replace('>', 'gt').replace('/', ' or ')}'
    #Create and save the SPC Chart
    spcChartIndiv(sub_data['data'].to_numpy(),  sub_data['WeekEndDate'].to_numpy(), title, good_dir)


######Create PDF
#Create a pdf at the desired path
pdf_filepath = pdfs_path + f'SPC Charts {datetime.today().strftime('%Y-%m-%d')}.pdf'
c = canvas.Canvas(pdf_filepath, pagesize=letter)
#Set key variables
images_per_page = 4
cols = 1
rows = 4
width, height = letter
margin = 20
img_width = (width - margin * 3) / cols
img_height = (height - margin * 3) / rows

#Loop over each image, start a new page if required
for i, img_path in enumerate(os.listdir(charts_path)):
    if i > 0 and i % images_per_page == 0:
        c.showPage()  # New page
    
    #Work out image position
    position = i % images_per_page
    col = position % cols
    row = position // cols
    x = margin + col * (img_width + margin)
    y = height - margin - (row + 1) * (img_height + margin)
    #add image to page
    c.drawImage(charts_path+img_path, x, y, width=img_width, height=img_height, 
                preserveAspectRatio=True, anchor='c')
#save pdf
c.save()

######Send Email


#############ERROR ModuleNotFoundError - pywintypes
#need to get to the bottom to be able to automate email sending (360 review emails work, maybe this needs its own venv?)
###########################################################





#send email with latest flagged output attatched    
# Create Outlook application object and mail item
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)    
# Set email properties
mail.To = open(base_path + 'emails.txt', 'r').read()
mail.Subject = 'TEST - SPC Charts'
mail.Body = """Hi,\n
Please find attatched the test SPC file \n
Emily"""
#Attatch the flagged file
mail.Attachments.Add(pdf_filepath)
# Send email
mail.Send()
print(f"Email sent successfully")
