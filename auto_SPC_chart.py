print('Importing libraries...')
import os
import time
import pyodbc
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
from reportlab.lib.pagesizes import A4
mpl.rcParams.update({'font.size': 18})


###################################################################################################
                                ####Define various filepaths####
###################################################################################################
print('Libraries imported. Running script...')
base_path = "G:/PerfInfo/Performance Management/OR Team/Emily Projects/General Analysis/SPC Charts"
symbol_path = base_path + "/Symbols/"
charts_path = base_path + "/SPC Charts/"
pdfs_path = base_path + "/pdfs/"
email_path = base_path + "/email files/"
#List to store the file name and final pdf position of all the plots created
plot_files = []
###################################################################################################
                                        ####Read in data####
###################################################################################################
#check which driver is installed (Simon doesn't have 17, so he needs *special code* to check and
#use whatever driver he does have).
driver = [driver for driver in pyodbc.drivers() if 'ODBC Driver' in driver][0]
sdmart_engine = create_engine('mssql+pyodbc://@SDMartDataLive2/InfoDB?'\
                              'trusted_connection=yes&driver='
                              + '+'.join(driver.split(' ')))

data_sql = """SELECT *
              FROM [InfoDB].[dbo].[ceo_flashcard]
              ORDER BY metric_id, WeekEndDate"""

full_data = pd.read_sql(data_sql, sdmart_engine)
sdmart_engine.dispose()
print('Data read sucessfully.  Beginning producing plots...')
###################################################################################################
                                ####Functions to create charts####
###################################################################################################

#######Line Charts
def Line_Chart(data):
    #pivot into one row perdate
    plot = data.pivot(index='WeekEndDate', columns='metric_desc', values='data')
    #Get the actual and plan column names to use
    actuals = [met for met in plot.columns if 'actuals' in met.lower()][0]
    plan = [met for met in plot.columns if 'plan' in met.lower()][0]
    #Get the overall name of the plot
    name = plan.replace('plan', '').replace('Plan', '')
    save_name = f'{id2} & {id2} {name.replace('<', 'lt').replace('>', 'gt').replace('/', ' or ')}'
    #plot
    fig, ax1 = plt.subplots(1,1, figsize=(15, 4))
    ax1.set_title('\n'.join(tw.wrap(name, 50)))
    plot[plan].plot(ax=ax1, label='Plan', color='darkorange', style='--', lw=3)
    plot[actuals].plot(ax=ax1, label='Actual', color='blue', lw=3)
    ax1.legend()
    fig.savefig(charts_path + save_name + ".png", format='png', bbox_inches="tight")
    plt.close()
    #record the file name and it's plot order
    plot_files.append((data['SortOrder'].mean().round(), save_name))

#######Stacked Bar Charts
def stacked_bar_chart(data):
    save_name = data['metric_desc'].iloc[0].split(' - ')[0].replace('/', ' or ')
    order = data['SortOrder'].mean().round()
    #pivot to put data in correct positions
    data = data.pivot(index='WeekEndDate', columns='metric_desc', values='data')
    #rename columns and get the data for each part of the stacked columns
    data.columns = ['P0', 'P1', 'P2', 'P3']
    x = data.index
    P0 = data['P0']
    P1 = data['P1']
    P2 = data['P2']
    P3 = data['P3']
    #create plot
    fig, ax1 = plt.subplots(1,1, figsize=(15, 4))
    ax1.set_title('\n'.join(tw.wrap(save_name, 50)))
    ax1.bar(x, P0, label='P0', width=5, color='blue')
    ax1.bar(x, P1, bottom=P0, label='P1', width=5, color='darkorange')
    ax1.bar(x, P2, bottom=P0+P1, label='P2', width=5, color='green')
    ax1.bar(x, P3, bottom=P0+P1+P2, label='P3', width=5, color='red')
    ax1.legend()
    #######Save the figure
    fig.savefig(charts_path + save_name + ".png", format='png', bbox_inches="tight")
    plt.close()
    #######record the file name and it's plot order
    plot_files.append((order, save_name))

#######SPC Charts
def SPC_Chart(sub_data, id, metric, good_dir = 'down', target=''):
    #####Initial calculations
    data_len = len(sub_data['data'])
    date = sub_data['WeekEndDate'].to_numpy()
    save_name = f'{id} {metric.replace('<', 'lt').replace('>', 'gt').replace('/', ' or ')}'
    recalc = sub_data['ReCalc'].str.contains('Y').any()
    #If no re-calculations defined, then we just calculate the thresholds on all the data
    if not recalc:
        #Find the mean and SD of this data (Use ddof=1 so that N-1 is the denominator)
        sub_data['data_mean'] = np.nanmean(sub_data['data'])
        sub_data['data_std'] = np.nanstd(sub_data['data'], ddof = 1)
        #Also find the moving range (max-min for each period). Have to add two as the
        #second value in square brackets is not included.
        sub_data['mov_R_mean'] = np.array([np.ptp(sub_data['data'].iloc[i:i+2])
                                           for i in range(data_len - 1)],
                                           dtype = float).mean()

    #if there are recalculation flags, then we calculate for individual sections
    else:
        #filter data to get the rows where recalc flags, and create recalc string
        recalcs = sub_data.loc[sub_data['ReCalc'] == 'Y'].copy()
        recalc_string = f'Limits have been recalculated {len(recalcs)} times due to: {', '.join(recalcs['Reason'])}.'
        #Get a list of where the bounaries of each recalc section begin and end
        boundary_list = [0] + recalcs.index.tolist() + [len(sub_data)]
        #empty list to store results
        split_calculations = []
        #Loop over each section of the data and perform calculations
        for i in range(len(boundary_list)-1):
            #select only the data for that section and calculate mean, std and moving range
            data_part = sub_data['data'].iloc[boundary_list[i] : boundary_list[i+1]]
            data_mean = np.nanmean(data_part)
            data_std = np.nanstd(data_part)
            mov_R = np.array([np.ptp(data_part[i:i+2]) for i in range(len(data_part) - 1)], dtype = float).mean()
            #add tuple of mean, std and moving range, to output list (repeated for the number of data points)
            split_calculations += [(data_mean, data_std, mov_R)] * len(data_part)
        #Create the columns on sub data for the changing mean, std and moving range
        sub_data[['data_mean', 'data_std', 'mov_R_mean']] = split_calculations

    #Find the action lines and warning lines for the mean plot (3sigma and 2 sigma)
    sub_data['u_mean_al'] =  sub_data['data_mean'] + 2.66 * sub_data['mov_R_mean']
    sub_data['l_mean_al'] = (sub_data['data_mean'] - 2.66 * sub_data['mov_R_mean']).clip(lower=0)
    sub_data['u_mean_wl'] =  sub_data['data_mean'] + (2/3)*2.66 * sub_data['mov_R_mean']
    sub_data['l_mean_wl'] = (sub_data['data_mean'] - (2/3)*2.66 * sub_data['mov_R_mean']).clip(lower=0)
 
    ######Set up Plot
    fig, ax1 = plt.subplots(1,1, figsize=(15, 4))
    ax1.set_title('\n'.join(tw.wrap(metric, 50)))
    #Set up the colours for each direction
    if good_dir == 'Down' or good_dir == 'Neutral':
        c_between_u = c_above_3s = c_run_above = 'darkorange'
        c_between_l = c_below_3s = c_run_below = 'blue'
    elif good_dir == 'Up':
        c_between_u = c_above_3s = c_run_above = 'blue'
        c_between_l = c_below_3s = c_run_below = 'darkorange'

    ######Plot the data
    #Plot the data as points
    markersize = 15
    ax1.plot(date, sub_data['data'], 'o-', color = 'slategrey', markersize = markersize, zorder = 10)
    #Add mean and upper and lower lines
    ax1.plot(date, sub_data['data_mean'], 'k--')
    ax1.plot(date, sub_data['u_mean_al'], 'r:')
    ax1.plot(date, sub_data['l_mean_al'], 'r:')

    #######Set xlabel and x ticks
    ax1.set_xlabel('Date')
    #Find the optimal interval for x-axis labels, we want roughly 5 ticks
    num_months = (date[-1] - date[0]).days / 30
    intvl = int(num_months // 5)
    #If this gives a number less than one, then need to look at days on the 
    #x-axis instead
    if intvl <= 1:
        intvl = 1 if num_months <=4 else 2
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
    #Search for consecutive points in between_lines and run_above/below
    def consecutiveNums(dataset, min_run):
        output = []
        for k, g in groupby(enumerate(dataset), lambda ix : ix[0] - ix[1]):
            row = list(map(itemgetter(1), g))
            if len(row) >= min_run: output.append(row)
        return output
    
    #Rules used here - out of control if:
    #N consecutive points above or below the mean (rising/falling)
    run_limit = 8
    run_above = sub_data.loc[sub_data['data'] > sub_data['data_mean']].index.tolist()
    run_below = sub_data.loc[sub_data['data'] < sub_data['data_mean']].index.tolist()
    run_above_points = consecutiveNums(run_above, run_limit)
    run_below_points = consecutiveNums(run_below, run_limit)

    #N points consecutively between 2sigma and 3sigma line (on the same side of the mean line)
    between_limit = 5
    between_lines_upper = sub_data.loc[(sub_data['data'] <= sub_data['u_mean_al'])
                                     & (sub_data['data'] >= sub_data['u_mean_wl'])].index.tolist()
    between_lines_lower = sub_data.loc[(sub_data['data'] >= sub_data['l_mean_al'])
                                     & (sub_data['data'] <= sub_data['l_mean_wl'])].index.tolist()
    between_upper_points = consecutiveNums(between_lines_upper, between_limit)
    between_lower_points = consecutiveNums(between_lines_lower, between_limit)
    
    #any point is outside of the 3sigma line
    above_3s = sub_data.loc[sub_data['data'] >= sub_data['u_mean_al']].index.tolist()
    below_3s = sub_data.loc[sub_data['data'] <= sub_data['l_mean_al']].index.tolist()

    ######Stability test
    #Test whether the most recent 50 data points are consistent with a system
    #in control
    test_start = data_len - 50
    if (data_len < 50) and (data_len > 29):
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
    stable_note = 'This process is stable.' if stable else 'This process is not stable.'
    if recalc:
        stable_note = '\n'.join(tw.wrap(recalc_string + '   ' + stable_note, 120))
    ax1.text(0.5, -0.3, stable_note, ha='center', fontsize=14, transform=ax1.transAxes)
    
    ######Symbols
    #To find which symbol to use to describe the recent changes
    #Assume no changes and re-assign if changes found
    state = 'no_change'
    symbol_test_start = data_len - 7
    if ((any(x > symbol_test_start for x in above_3s))
        and (any(x > symbol_test_start for x in below_3s))):
        state = 'warning'
    elif (any(x > symbol_test_start for x in above_3s)):
        state = 'Up'
    elif (any(x > symbol_test_start for x in below_3s)):
        state = 'Down'
    else:
        for run_data in filter(None, [run_above_points, between_upper_points]):
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
    def plot_extrenuous(test_list,  colour, order, window):
        for i in range(len(test_list)):
            #Add the window to the x-values so that the points are in the right place (if required)
            dates = date[test_list[i]] if window else [date[x] for x in test_list[i]]
            #Plot the different coloured points
            ax1.plot(dates,
                    sub_data.loc[test_list[i], 'data'],
                    color = colour, marker = 'o', markersize = markersize,
                    zorder = order, linestyle = 'None')

    #Highlight the points that were above the mean for N consecutive dates
    plot_extrenuous(run_above_points, c_run_above, 11, False)
    plot_extrenuous(run_below_points, c_run_below, 11, False)
    #Highlight the points that were between the 2 and 3 sigma lines for N consecutive dates.
    plot_extrenuous(between_upper_points, c_between_u, 12, False)
    plot_extrenuous(between_lower_points, c_between_l, 12, False)
    #Highlight the points that fell above or below the 3 sigma line
    plot_extrenuous(above_3s, c_above_3s, 13, True)
    plot_extrenuous(below_3s, c_below_3s, 13, True)
                
    #######Save the figure
    fig.savefig(charts_path + save_name + ".png", format='png', bbox_inches="tight")
    plt.close()
    #######record the file name and it's plot order
    plot_files.append((sub_data['SortOrder'].mean(), save_name))

###################################################################################################
                                ####Create Actual/Plan plots####
###################################################################################################
#Get the list of pairs usign thier ids (these are consecutively numbered which is helpful)
metric_id_pairs = [(i, i+1) for i in range(1038, 1050, 2)]

for id1, id2 in metric_id_pairs:
    #for each plan/actual pair, get the data and create the line graph.
    data = full_data.loc[(full_data['metric_id'] == id1) | (full_data['metric_id'] == id2)].copy()
    Line_Chart(data)

###################################################################################################
                                ####Stacked Bar Chart Plot####
###################################################################################################
#Get the list of metrics to create the stacked bar plot for discharges
metric_ids = [i for i in range(1024, 1028)]
#filter to data for these ids
data = full_data.loc[full_data['metric_id'].isin(metric_ids)].copy()
stacked_bar_chart(data)

###################################################################################################
                                  ####Create SPC Charts####
###################################################################################################
######Loop over each metric and create it's SPC Chart
#List of unique metrics
complete_ids = metric_ids + list(sum(metric_id_pairs, ()))
metrics = (full_data.loc[~full_data['metric_id'].isin(complete_ids),
                         ['metric_id', 'metric_desc', 'GoodDirection']]
                         .drop_duplicates().values.tolist())
#Loop over each metric and create it's SPC
for id, metric, good_dir in metrics:
    #Filter data to that metric, ensure it's sorted
    sub_data = full_data.loc[full_data['metric_id'] == id].copy().reset_index()
    #Create and save the SPC Chart
    SPC_Chart(sub_data, id, metric, good_dir)

print(f'All charts created and saved in {charts_path}')
time.sleep(10)

###################################################################################################
                                    ####Create pdf file####
###################################################################################################
#Create a pdf at the desired path
pdf_filepath = pdfs_path + f'SPC Charts {datetime.today().strftime('%Y-%m-%d')}.pdf'
c = canvas.Canvas(pdf_filepath)
#Set key variables
images_per_page = 4
cols = 1
rows = 4
width, height = A4
margin = 2
img_width = (width - margin * 3) / cols
img_height = (height - margin * 3) / rows

ordered_file_list = (pd.DataFrame(plot_files, columns=['order', 'file name'])
                     .sort_values(by='order')['file name'].values.tolist())
#Loop over each image, start a new page if required
for i, img_name in enumerate(ordered_file_list):
    #Create a new page if 4 images have already been added
    if i > 0 and i % images_per_page == 0:
        c.showPage()  # New page
    #Work out image position
    position = i % images_per_page
    col = position % cols
    row = position // cols
    x = margin + col * (img_width + margin)
    y = height - margin - (row + 1) * (img_height + margin)
    #add image to page
    c.drawImage(charts_path+img_name+'.png', x, y, width=img_width, height=img_height, 
                preserveAspectRatio=True, anchor='c')
#save pdf
c.save()
print(f'new pdf created in {pdfs_path}')

###################################################################################################
                                        ####Send email####
###################################################################################################
#send email with latest flagged output attatched 
# Create Outlook application object and mail item
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)    
# Set email properties
mail.To = open(email_path + 'email addresses.txt', 'r').read()
#mail.To = 'e.obrien6@nhs.net'
mail.Subject = 'Metric Charts'
#HTML of email content with signature footer
email_content = ("""<p>Hi,</p>
                    <p>Please find attatched the outputted metric charts file.</p>"""
                    + open(email_path + 'email signature.txt', 'r').read())
#Add email signature to the end of the html
attachment = mail.Attachments.Add(email_path + 'email footer.png')
attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
mail.HTMLBody = email_content + "<html><body><img src=""cid:MyId1""></body></html>"
#Attatch the flagged file
mail.Attachments.Add(pdf_filepath)
# Send email
mail.Send()
print(f"Email sent successfully, process complete.")
