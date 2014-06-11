# -*- coding: utf-8 -*-

# Coded by evil_inside
# Contact at evil_inside@hotmail.gr

import csv
import numpy as np
import os
import subprocess
import sys
from itertools import ifilter
import datetime
import time

# Functions

def get_shmeio_stats(data, reference_list):
    shmeio_stats = '-'
    if data:
        tally = (data.count(i) for i in reference_list)
        shmeio_stats = '-'.join(map(str, tally))
    else:
        shmeio_stats ='-'
        
    if shmeio_stats == '0-0-0':
        shmeio_stats ='-'
        
    return shmeio_stats

def get_mesos_oros(data):
    data = list(ifilter(lambda c: c != '-', data))
    try:
        if data:
            data = round(np.average(data[-6:]), 2)
        else:
            data = '-'
    except Exception, e:
        data = 'ERROR'
    return data

def current_time():
    now = datetime.datetime.now()
    return ''.join(['[', now.strftime("%H:%M:%S"), ']'])

###############################################################################
# Διαγραφή προηγούμενων αρχείων
print
print current_time(), u'Διαγραφή προηγούμενων αρχείων...',
csv_files = [x for x in os.listdir('.') if x.endswith('.csv')]
for each_csv in csv_files:
  os.remove(each_csv)
  
try:
    os.remove('BET365_stats.xlsb')
except:
    pass
print u'Επιτυχής.'

# Μετατροπή XLSB -> CSV
print current_time(), u'Μετατροπή BET365.xlsb σε BET365.csv...',
subprocess.call("cscript ExcelToCsv.vbs BET365.xlsb BET365.csv",
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=False) # Supress any messages
print u'Επιτυχής.'

# Επεξεργασία δεδομένων
print current_time(), u'Επεξεργασία δεδομένων:',
shmeia_list = ['1', 'x', '2']
with open('BET365.csv', 'rb') as f:
    bet365_data = csv.reader(f)
    bet365_matches = list(bet365_data)[1:]

stats_dict = {}
all_data = []
with open('BET365_stats.csv', 'wb') as f:
    bet365_stats = csv.writer(f)

    headers = ['id', 'Protathlima', 'Xronia', 'Hmerominia', 'Home', '1', 'x', '2', \
                'Away', 'Skor', 'Skor 1', 'Skor 2', 'Simeio', 'Favori', 'Under-over', \
                'Home 1', 'Home x', 'Home 2', 'Away 1', 'Away x', 'Away 2', \
                'Home record all years', 'Away record all years', \
                'Forma home last 6 home', 'Forma away last 6 away', \
                'Forma home last 6 home-away', 'Forma away last 6 home-away', \
                'Akrivis protatlima', 'Akrivis genika', 'Mesos oros gkol home last 6 at home', \
                'Mesos oros gkol away last 6 at away']
    
    bet365_stats.writerow(headers)
    
    total = len(bet365_matches)
    start = time.time()
    for index, each_match in enumerate(bet365_matches):
        id = index       
        protathlima, xronia, match_date, home, odd_1, odd_x, odd_2, away, score, score_1, score_2, simeio, favori, under_over = each_match
        
        match_stats = [id, protathlima, xronia, match_date, home,
                        odd_1, odd_x, odd_2, away, score, score_1,
                        score_2, simeio, favori, under_over]
        
        stats_keys = [(home, 'entos', '1', odd_1), (home, 'entos', 'x', odd_x), (home, 'entos', '2', odd_2), ## home_1, home_x, home_2
                      (away, 'ektos', '1', odd_1), (away, 'ektos', 'x', odd_x), (away, 'ektos', '2', odd_2), ## away_1, away_x, away_2
                      (home, 'entos', protathlima), (away, 'ektos', protathlima)] ## home_all_yrs_protathlima, away_all_yrs_protathlima

        stats_last_6_keys = [(home, 'entos', protathlima, xronia), (away, 'ektos', protathlima, xronia), ## home_forma_last_6_home, away_forma_last_6_away
                             (home, 'entos-ektoss', protathlima, xronia), (away, 'entos-ektoss', protathlima, xronia)] # home_forma_last_6_home_away, away_forma_last_6_home_away]

        stats_keys_odds = [(odd_1, odd_x, odd_2, protathlima), (odd_1, odd_x, odd_2)]
        
        stats_goal_keys = [(home, 'entos', 'goals', protathlima, xronia), (away, 'ektos', 'goals', protathlima, xronia)] # mesos_oros_goal_home_last_6_home, mesos_oros_goal_away_last_6_away

        for stats_key in stats_keys:
            try:
                stats_dict[stats_key].append(simeio)
            except:
                stats_dict[stats_key] = [simeio]        
            stats_key_values = stats_dict[stats_key][:-1]
            match_stats.append(get_shmeio_stats(stats_key_values, shmeia_list))
            
        for stats_last_6_key in stats_last_6_keys:
            try:
                stats_dict[stats_last_6_key].append(simeio)
            except:
                stats_dict[stats_last_6_key] = [simeio]
            stats_last_6_key_values = stats_dict[stats_last_6_key][:-1][-6:]
            match_stats.append(get_shmeio_stats(stats_last_6_key_values, shmeia_list))
            
        for stats_key_odds in stats_keys_odds:
            try:
                stats_dict[stats_key_odds].append(simeio)
            except:
                stats_dict[stats_key_odds] = [simeio]
            stats_key_odds_values = stats_dict[stats_key_odds][:-1]
            match_stats.append(get_shmeio_stats(stats_key_odds_values, shmeia_list))
 
        for stats_goal_key in stats_goal_keys:
            if stats_goal_key[1] == 'entos':
                try:
                    goal_value = float(score_1)
                except: 
                    goal_value = '-'
            elif stats_goal_key[1] == 'ektos':
                try:
                    goal_value = float(score_2)
                except: 
                    goal_value = '-'
            try:
                stats_dict[stats_goal_key].append(goal_value)
            except:
                stats_dict[stats_goal_key] = [goal_value]        
            stats_goal_key_values = stats_dict[stats_goal_key][:-1]
            match_stats.append(get_mesos_oros(stats_goal_key_values))           
               
               
        all_data.append(match_stats)

        message = ' '.join([current_time(),  u'Επεξεργασία δεδομένων:'])
        progress = str(round(index/float(total), 3) * 100.)
        sys.stdout.write("\r" + message + " {0}/{1}".format(index+1, total) + ' ' + progress + '%')
        sys.stdout.flush()
        
    print 
    print current_time(), u'Αποθήκευση δεδομένων ως BET365_stats.csv...',
    bet365_stats.writerows(all_data)
    print u'Επιτυχής.'

print current_time(), u'Η διαδικασία ολοληρώθηκε.'
