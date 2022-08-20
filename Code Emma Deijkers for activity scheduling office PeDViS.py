# -*- coding: utf-8 -*-
"""
Created on Fri Jun 10 12:10:40 2022

@author: Emma
"""
import random
import pandas as pd
import numpy as np
import math
import xlsxwriter
from datetime import datetime
import pytz



#INPUT HERE WHICH SCENARIO (1-18)

scenario = 14

#INPUT HERE HOMW MUCH SCHEDULES YOU WANT TO MAKE/ HOW MANY EMPLOYEES YOU HAVE

nr_of_schedules = 40



if scenario == 1:
    mr_meetings_day = 0
    office_meetings_day = 0
    nr_people_meetingroom = 0

elif scenario == 2:
    mr_meetings_day = 3
    office_meetings_day = 4
    nr_people_meetingroom = 7
    
elif scenario == 3:
    mr_meetings_day = 8
    office_meetings_day = 4
    nr_people_meetingroom = 7

elif scenario == 4:
    mr_meetings_day = 1
    office_meetings_day = 4
    nr_people_meetingroom = 7
    
elif scenario == 5:
    mr_meetings_day = 3
    office_meetings_day = 4
    nr_people_meetingroom = 14
    
elif scenario == 6:
    mr_meetings_day = 1
    office_meetings_day = 4
    nr_people_meetingroom = 14

elif scenario == 7:
    mr_meetings_day = 3
    office_meetings_day = 4
    nr_people_meetingroom = 3
    
elif scenario == 8:
    mr_meetings_day = 8
    office_meetings_day = 4
    nr_people_meetingroom = 3
    
elif scenario == 9:
    mr_meetings_day = 1
    office_meetings_day = 4
    nr_people_meetingroom = 3

elif scenario == 10:
    mr_meetings_day = 3
    office_meetings_day = 1
    nr_people_meetingroom = 7
    
elif scenario == 11:
    mr_meetings_day = 8
    office_meetings_day = 1
    nr_people_meetingroom = 7
    
elif scenario == 12:
    mr_meetings_day = 1
    office_meetings_day = 1
    nr_people_meetingroom = 7
    
elif scenario == 13:
    mr_meetings_day = 3
    office_meetings_day = 1
    nr_people_meetingroom = 14
    
elif scenario == 14:
    mr_meetings_day = 8
    office_meetings_day = 1
    nr_people_meetingroom = 14
    
elif scenario == 15:
    mr_meetings_day = 1
    office_meetings_day = 1
    nr_people_meetingroom = 14
    
elif scenario == 16:
    mr_meetings_day = 3
    office_meetings_day = 1
    nr_people_meetingroom = 3
    
elif scenario == 17:
    mr_meetings_day = 8
    office_meetings_day = 1
    nr_people_meetingroom = 3

elif scenario == 18:
    mr_meetings_day = 1
    office_meetings_day = 1
    nr_people_meetingroom = 3






dict_mr1 = {1: 'mr1_1', 2 : 'mr1_2' , 3: 'mr1_3', 4: 'mr1_4', 5: 'mr1_5', 6: 'mr1_6', 7: 'mr1_7', 8: 'mr1_8', 9: 'mr1_9', 10: 'mr1_10', 11: 'mr1_11', 12: 'mr1_12' , 13: 'mr1_13', 14: 'mr1_14'}

dict_mr2 = {1: 'mr2_1', 2: 'mr2_2', 3: 'mr2_3', 4: 'mr2_4', 5: 'mr2_5', 6: 'mr2_6', 7: 'mr2_7', 8: 'mr2_8', 9: 'mr2_9', 10: 'mr2_10', 11: 'mr2_11', 12: 'mr2_12', 13: 'mr2_13', 14: 'mr2_14'}



dict_8meeting =  {0: dict_mr1, 1: dict_mr2, 2:  dict_mr1, 3: dict_mr2, 4: dict_mr1, 5: dict_mr2, 6: dict_mr1, 7: dict_mr2 }
dict_3meeting =  {0: dict_mr1, 1: dict_mr2, 2: dict_mr1 }
dict_1meeting =  {0: dict_mr1}

dict_of_dict = {8: dict_8meeting, 3: dict_3meeting, 1: dict_1meeting}


dict_8_entertime = {1: 30, 2: 30, 3:145, 4: 145, 5: 260, 6: 260, 7: 375, 8: 375} 
dict_3_entertime = {1: 30, 2: 30, 3: 145}
dict_1_entertime = {1: 145}
dict_nr_entertime = {8:dict_8_entertime, 3: dict_3_entertime, 1: dict_1_entertime} #input is de mr_meeting_per_day


dict_8_leavetime = {1: 85, 2: 80, 3:200, 4: 195, 5: 315, 6: 310, 7:430, 8: 425} 
#hierbij alleen de even getallen verandert zodat mr 2 eerder afloop
dict_3_leavetime = {1: 85, 2: 80, 3: 200}
dict_1_leavetime = {1: 200}
dict_nr_leavetime = {8:dict_8_leavetime, 3: dict_3_leavetime, 1: dict_1_leavetime} #input is de mr_meeting_per_day



#dictionaries chairs per office
officemeeting_1 = {1: 'officemeeting_1_chair1', 2 : 'officemeeting_1_chair2', 3 : 'officemeeting_1_chair3'}
officemeeting_2 = {1: 'officemeeting_2_chair1', 2 : 'officemeeting_2_chair2', 3 : 'officemeeting_2_chair3'}
officemeeting_3 = {1: 'officemeeting_3_chair1', 2 : 'officemeeting_3_chair2', 3 : 'officemeeting_3_chair3'}
officemeeting_4 = {1: 'officemeeting_4_chair1', 2 : 'officemeeting_4_chair2', 3 : 'officemeeting_4_chair3'}
officemeeting_5 = {1: 'officemeeting_5_chair1', 2 : 'officemeeting_5_chair2', 3 : 'officemeeting_5_chair3'}
officemeeting_6 = {1: 'officemeeting_6_chair1', 2 : 'officemeeting_6_chair2', 3 : 'officemeeting_6_chair3'}
officemeeting_7 = {1: 'officemeeting_7_chair1', 2 : 'officemeeting_7_chair2', 3 : 'officemeeting_7_chair3'}
officemeeting_8 = {1: 'officemeeting_8_chair1', 2 : 'officemeeting_8_chair2', 3 : 'officemeeting_8_chair3'}
officemeeting_9 = {1: 'officemeeting_9_chair1', 2 : 'officemeeting_9_chair2', 3 : 'officemeeting_9_chair3'}
officemeeting_10 = {1: 'officemeeting_10_chair1', 2 : 'officemeeting_10_chair2', 3 : 'officemeeting_10_chair3'}
officemeeting_11 = {1: 'officemeeting_11_chair1', 2 : 'officemeeting_11_chair2', 3 : 'officemeeting_11_chair3'}
officemeeting_12 = {1: 'officemeeting_12_chair1', 2 : 'officemeeting_12_chair2', 3 : 'officemeeting_12_chair3'}
officemeeting_13 = {1: 'officemeeting_13_chair1', 2 : 'officemeeting_13_chair2', 3 : 'officemeeting_13_chair3'}
officemeeting_14 = {1: 'officemeeting_14_chair1', 2 : 'officemeeting_14_chair2', 3 : 'officemeeting_14_chair3'}
officemeeting_15 = {1: 'officemeeting_15_chair1', 2 : 'officemeeting_15_chair2', 3 : 'officemeeting_15_chair3'}
officemeeting_16 = {1: 'officemeeting_16_chair1', 2 : 'officemeeting_16_chair2', 3 : 'officemeeting_16_chair3'}
officemeeting_17 = {1: 'officemeeting_17_chair1', 2 : 'officemeeting_17_chair2', 3 : 'officemeeting_17_chair3'}
officemeeting_18 = {1: 'officemeeting_18_chair1', 2 : 'officemeeting_18_chair2', 3 : 'officemeeting_18_chair3'}
officemeeting_19 = {1: 'officemeeting_19_chair1', 2 : 'officemeeting_19_chair2', 3 : 'officemeeting_19_chair3'}
officemeeting_20 = {1: 'officemeeting_20_chair1', 2 : 'officemeeting_20_chair2', 3 : 'officemeeting_20_chair3'}
officemeeting_21 = {1: 'officemeeting_21_chair1', 2 : 'officemeeting_21_chair2', 3 : 'officemeeting_21_chair3'}
officemeeting_22 = {1: 'officemeeting_22_chair1', 2 : 'officemeeting_22_chair2', 3 : 'officemeeting_22_chair3'}
officemeeting_23 = {1: 'officemeeting_23_chair1', 2 : 'officemeeting_23_chair2', 3 : 'officemeeting_23_chair3'}
officemeeting_24 = {1: 'officemeeting_24_chair1', 2 : 'officemeeting_24_chair2', 3 : 'officemeeting_24_chair3'}
officemeeting_25 = {1: 'officemeeting_25_chair1', 2 : 'officemeeting_25_chair2', 3 : 'officemeeting_25_chair3'}
officemeeting_26 = {1: 'officemeeting_26_chair1', 2 : 'officemeeting_26_chair2', 3 : 'officemeeting_26_chair3'}
officemeeting_27 = {1: 'officemeeting_27_chair1', 2 : 'officemeeting_27_chair2', 3 : 'officemeeting_27_chair3'}
officemeeting_28 = {1: 'officemeeting_28_chair1', 2 : 'officemeeting_28_chair2', 3 : 'officemeeting_28_chair3'}
officemeeting_29 = {1: 'officemeeting_29_chair1', 2 : 'officemeeting_29_chair2', 3 : 'officemeeting_29_chair3'}
officemeeting_30 = {1: 'officemeeting_30_chair1', 2 : 'officemeeting_30_chair2', 3 : 'officemeeting_30_chair3'}
officemeeting_31 = {1: 'officemeeting_31_chair1', 2 : 'officemeeting_31_chair2', 3 : 'officemeeting_31_chair3'}
officemeeting_32 = {1: 'officemeeting_32_chair1', 2 : 'officemeeting_32_chair2', 3 : 'officemeeting_32_chair3'}
officemeeting_33 = {1: 'officemeeting_33_chair1', 2 : 'officemeeting_33_chair2', 3 : 'officemeeting_33_chair3'}
officemeeting_34 = {1: 'officemeeting_34_chair1', 2 : 'officemeeting_34_chair2', 3 : 'officemeeting_34_chair3'}
officemeeting_35 = {1: 'officemeeting_35_chair1', 2 : 'officemeeting_35_chair2', 3 : 'officemeeting_35_chair3'}
officemeeting_36 = {1: 'officemeeting_36_chair1', 2 : 'officemeeting_36_chair2', 3 : 'officemeeting_36_chair3'}
officemeeting_37 = {1: 'officemeeting_37_chair1', 2 : 'officemeeting_37_chair2', 3 : 'officemeeting_37_chair3'}
officemeeting_38 = {1: 'officemeeting_38_chair1', 2 : 'officemeeting_38_chair2', 3 : 'officemeeting_38_chair3'}
officemeeting_39 = {1: 'officemeeting_39_chair1', 2 : 'officemeeting_39_chair2', 3 : 'officemeeting_39_chair3'}
officemeeting_40 = {1: 'officemeeting_40_chair1', 2 : 'officemeeting_40_chair2', 3 : 'officemeeting_40_chair3'}



officemeeting_offices = {1:officemeeting_1, 2: officemeeting_2, 3: officemeeting_3, 4: officemeeting_4, 5: officemeeting_5, 6: officemeeting_6, 7: officemeeting_7, 8: officemeeting_8, 9: officemeeting_9, 10: officemeeting_10, 11: officemeeting_11, 12 : officemeeting_12, 13: officemeeting_13, 14: officemeeting_14, 15: officemeeting_15, 16: officemeeting_16, 17: officemeeting_17, 18: officemeeting_18, 19: officemeeting_19, 20: officemeeting_20, 21: officemeeting_21, 22: officemeeting_22, 23: officemeeting_23, 24: officemeeting_24, 25: officemeeting_25, 26: officemeeting_26, 27: officemeeting_27, 28: officemeeting_28, 29: officemeeting_29, 30: officemeeting_30, 31: officemeeting_31, 32: officemeeting_32, 33: officemeeting_33, 34: officemeeting_34, 35: officemeeting_35, 36: officemeeting_36, 37: officemeeting_37, 38: officemeeting_38, 39: officemeeting_39, 40: officemeeting_40}


begin_5officemeeting=  {0:90, 1: 205, 2: 320, 3:435, 4: 470} #not used anymore
begin_4officemeeting = {0:90, 1: 205, 2: 320, 3 : 435}
begin_1officemeeting  = {0 : 320}

leave_5officemeeting=  {0: 135, 1: 250, 2 : 365, 3: 480, 4:515} #not used anymore
leave_4officemeeting = {0: 135, 1: 250, 2 : 365, 3: 480}
leave_1officemeeting  = {0:365}


begintimes_office = {5: begin_5officemeeting, 4: begin_4officemeeting, 1: begin_1officemeeting}
leavetimes_office = {5: leave_5officemeeting, 4: leave_4officemeeting, 1: leave_1officemeeting}



schedules= {}

for i in range(nr_of_schedules):
        schedules["schedule" + str(i)] = pd.DataFrame(index=np.arange(540), columns=np.arange(1))


    
transitionmatrix = pd.DataFrame({'Personal office work area':[0, 0.7, 0.75, 0.67, 0.5, 1], 'Coffee':[0.5, 0.15, 0, 0.33, 0.45, 0], 'toilet': [0.5, 0.15, 0.25, 0, 0.05, 0]})

transitionmatrix_less = pd.DataFrame ({ 'Coffee':[0.5, 0.5, 0, 1,  0.9, 0.5], 'toilet': [0.5, 0.5, 1, 0, 0.1, 0.5]})

transitionmatrix_toilet_office = pd.DataFrame({'Personal office work area':[0, 0.82, 0.75, 1, 0.725, 1], 'toilet': [1, 0.18, 0.25, 0, 0.275, 0]})

transitionmatrix_coffee_office = pd.DataFrame({'Personal office work area':[0, 0.82, 1, 0.67, 0.5025, 1], 'Coffee':[1, 0.18, 0, 0.33, 0.4525, 0]})



def add_meetings(nr_people_meetingroom, mr_meetings_day,office_meetings_day):
    j = 0
    meeting_slot = 1
    persons_shift1  = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39] #CHANGE
    persons_shift2  = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39] #CHANGE
    persons_shift3  = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39] #CHANGE
    persons_shift4  = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39]#CHANGE
    dict_meetingrooms = dict_of_dict[mr_meetings_day]
    begin_meeting_options = dict_nr_entertime[mr_meetings_day]
    leave_meeting_options = dict_nr_leavetime[mr_meetings_day]
    for item in range (mr_meetings_day):
        i = 0
        chairs = [1,2,3,4,5,6,7,8,9,10,11,12,13,14]
        while i < nr_people_meetingroom: 
           # print(f'meetingslot = {meeting_slot}')
            if meeting_slot == 1:
                chosen_person = random.choice(persons_shift1)
                my_schedule = schedules[f'schedule{chosen_person}']
                persons_shift1.remove(chosen_person)
            elif meeting_slot == 2:
                chosen_person = random.choice(persons_shift1)
                my_schedule = schedules[f'schedule{chosen_person}']
                persons_shift1.remove(chosen_person)
            elif meeting_slot == 3:
                chosen_person = random.choice(persons_shift2)
                my_schedule = schedules[f'schedule{chosen_person}']
                persons_shift2.remove(chosen_person)
            elif meeting_slot == 4:
                chosen_person = random.choice(persons_shift2)
                my_schedule = schedules[f'schedule{chosen_person}']
                persons_shift2.remove(chosen_person)
            elif meeting_slot == 5:
                chosen_person = random.choice(persons_shift3)
                my_schedule = schedules[f'schedule{chosen_person}']
                persons_shift3.remove(chosen_person)
            elif meeting_slot ==  6:
                chosen_person = random.choice(persons_shift3)
                my_schedule = schedules[f'schedule{chosen_person}']
                persons_shift3.remove(chosen_person)
            elif meeting_slot == 7:
                chosen_person = random.choice(persons_shift4)
                my_schedule = schedules[f'schedule{chosen_person}']
                persons_shift4.remove(chosen_person)
            elif meeting_slot == 8:
                chosen_person = random.choice(persons_shift4)
                my_schedule = schedules[f'schedule{chosen_person}']
                persons_shift4.remove(chosen_person)

            meeting_room = dict_meetingrooms[j]
            chair = random.choice(chairs)
            chairs.remove(chair) #delete chair from options
            begin_meeting = begin_meeting_options[meeting_slot] 
            leave_meeting = leave_meeting_options[meeting_slot]
            pers_meeting = meeting_room[chair]
            vroegtijd = (leave_meeting -1)


            if meeting_slot in [1,3,5,7]:
                my_schedule[(begin_meeting) : (leave_meeting+1)] = pers_meeting
            elif meeting_slot in [2,4,6,8]:
                #my_schedule[(begin_meeting) : (leave_meeting-1)] = pers_meeting
                my_schedule[(begin_meeting) : vroegtijd] = pers_meeting
                
            i += 1
        j += 1
        meeting_slot +=1

        
        
    k = 0
    m = 0
    begin_officemeeting_options = begintimes_office[office_meetings_day]
    leave_officemeeting_options = leavetimes_office[office_meetings_day]
    while k < office_meetings_day:
        late = []
        chosen_offices = []
        offices  = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40] #CHANGE
        visitors = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39]   #CHANGE
        for item in range(13): #CHANGE
            chosen_own_office = random.choice(offices)
            offices.remove(chosen_own_office)
            chosen_offices.append(chosen_own_office)
            prob_late = random.choice([1,2])
            if prob_late ==1:
                late.append(chosen_own_office)
                #print("is 1")
            person = chosen_own_office - 1
            visitors.remove(person)
            my_schedule = schedules[f'schedule{person}']
            pers_meeting_office = officemeeting_offices[chosen_own_office]
            pers_meeting_chair = pers_meeting_office[1]
            begin_meeting_office = begin_officemeeting_options[k] 
            leave_meeting_office = leave_officemeeting_options[k]
            if chosen_own_office in late:
                #print("ja")
                my_schedule[(begin_meeting_office+10) : (leave_meeting_office+10)] = pers_meeting_chair
            else:
               # print('nee')
                my_schedule[begin_meeting_office : (leave_meeting_office)] = pers_meeting_chair
        #other_offices = [1, 2, 3, 
        for item in range(13): #CHANGE
            chosen_visitor_1 = random.choice(visitors)
            visitors.remove(chosen_visitor_1)
            chosen_visitor_2 = random.choice(visitors)
            visitors.remove(chosen_visitor_2)
            chosen_other_office = random.choice(chosen_offices)
            chosen_offices.remove(chosen_other_office)
            my_schedule_1 = schedules[f'schedule{chosen_visitor_1}']
            my_schedule_2 = schedules[f'schedule{chosen_visitor_2}']
            pers_meeting_office = officemeeting_offices[chosen_other_office]
            chosen_chair_1 = 2
            chosen_chair_2 = 3
            pers_meeting_chair_office_1 = pers_meeting_office[chosen_chair_1]
            pers_meeting_chair_office_2 = pers_meeting_office[chosen_chair_2]
            #tijden deel           
            
            begin_meeting_office = begin_officemeeting_options[k] 
            leave_meeting_office = leave_officemeeting_options[k]
           # print(leave_meeting_office)
            if chosen_other_office in late:
                my_schedule_1[(begin_meeting_office+10) : (leave_meeting_office+10)] = pers_meeting_chair_office_1
                my_schedule_2[(begin_meeting_office+10) : (leave_meeting_office+10)] = pers_meeting_chair_office_2
            else:
                my_schedule_1[begin_meeting_office : (leave_meeting_office)] = pers_meeting_chair_office_1
                my_schedule_2[begin_meeting_office : (leave_meeting_office)] = pers_meeting_chair_office_2
            
        k += 1    
        #print(k)
            
            
            

if scenario != 1:
    add_meetings(nr_people_meetingroom, mr_meetings_day, office_meetings_day)

duration_coffee = 2
duration_toilet = 1
duration_sink = 2
entertimes_list = []

for s in range(nr_of_schedules):
    my_schedule = schedules[f'schedule{s}']
    enter_time =  random.randint(1, 30)
    print(f'employee is {s}')
    tick = enter_time
    my_schedule.iloc[0: enter_time, 0] = 'nothing'
    entertimes_list.append(enter_time)
    leave_time = random.randint(510, 539) # Draw leave time between 17.00 and 17.30
    my_schedule.iloc[leave_time:540,0] = 'nothing' 
    office_activity = f'office_{s}'


    while tick < 539: 
        
        #while type(my_schedule.iloc[(tick-1),0]) == float:
         #   print(tick)
          #  print("heeft nan")
         #   tick -= 1
           # print(tick)
                                                            
        if type(my_schedule.iloc[(tick + 1),0]) == float:
            #print(my_schedule.iloc[(tick-1), 0])
            #print(tick)

            if my_schedule.iloc[tick-1,0] == 'nothing':
                row = 5 
            #elif type(my_schedule.iloc[tick-1,0]) == float: #omdat alleen na 1 min office er een nan tussen zit, dus komen dan uit ofifce
             #   row = 0
            elif 'meeting' in my_schedule.iloc[(tick-1),0]:# == True:   
                row = 4
            elif 'mr' in my_schedule.iloc[(tick-1), 0]:# == True:       
                row = 1
            elif 'office' in (my_schedule.iloc[(tick-1), 0]): #== True:
                row = 0
            elif 'coffee' in (my_schedule.iloc[(tick-1), 0]):# ==True:
                row = 2
            elif 'sink' in (my_schedule.iloc[(tick-1), 0]):
                row = 3
            else:
                print("error determining row")
                print(f'tick is {tick}')


            
      
                
            if my_schedule.iloc[(tick+1):(tick+121), 0].isnull().all():
                #print("2 uur over")
                coffee_list = list(my_schedule.iloc[(tick-5):(tick-3),0].str.contains('coffee'))
                toilet_list = list(my_schedule.iloc[(tick-5):(tick-3),0].str.contains('sink'))
                
                if True in coffee_list:
                    next_space = random.choices(['office', 'restroom'],transitionmatrix_toilet_office.iloc[row,:])#kies tussen toilet and office
                    

                        
                elif True in toilet_list:
                    next_space = random.choices(['office', 'coffee'],transitionmatrix_coffee_office.iloc[row,:])#kies tussen wc en kantoor
                    #print("chosen trans")
                    
                
                else:          
                    next_space = random.choices(['office', 'coffee', 'restroom'], transitionmatrix.iloc[row,:])
                
                #print(next_space)



                        
                if next_space == ['coffee']:
                    bezet_coffee_1 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_1_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_1'))
                        if True in coffee_occupied_1_list:
                            #print(i)
                            bezet_coffee_1 = True
                            
                    bezet_coffee_2 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_2_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_2'))
                        if True in coffee_occupied_2_list:
                            #print(i)
                            bezet_coffee_2 = True
                            
                    bezet_coffee_3 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_3_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_3'))
                        if True in coffee_occupied_3_list:
                            #print(i)
                            bezet_coffee_3 = True
                            
                    bezet_coffee_4 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_4_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_4'))
                        if True in coffee_occupied_4_list:
                            #print(i)
                            bezet_coffee_4 = True
                            
                            
                    bezet_coffee_5 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_5_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_5'))
                        if True in coffee_occupied_5_list:
                            #print(i)
                            bezet_coffee_5 = True
                            
                    bezet_coffee_6 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_6_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_6'))
                        if True in coffee_occupied_6_list:
                            #print(i)
                            bezet_coffee_6 = True
                            
                    
                    if bezet_coffee_1 == False:
                        chosen_coffee = 1
                    elif bezet_coffee_1 == True:
                         if bezet_coffee_2 == False:
                             chosen_coffee = 2
                         elif bezet_coffee_2 == True:
                             if bezet_coffee_3 == False:
                                 chosen_coffee= 3
                             elif bezet_coffee_3 == True:
                                 if bezet_coffee_4 == False:
                                     chosen_coffee = 4
                                 elif bezet_coffee_4 == True:
                                     if bezet_coffee_5 == False:
                                         chosen_coffee = 5
                                     elif bezet_coffee_5 == True:
                                         if bezet_coffee_6 == False:
                                             chosen_coffee = 6
                                         elif bezet_coffee_6 == True:
                                             chosen_coffee = 7
                                             print("alle koffie al bezet 2 uur")
                                             print(tick)
                 
                    
                    if chosen_coffee == 1:
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_1'
                        tick += duration_coffee
                    elif chosen_coffee == 2:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_2'
                        tick += duration_coffee
                    elif chosen_coffee == 3:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_3'
                        tick += duration_coffee
                    elif chosen_coffee == 4:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_4'
                        tick += duration_coffee
                    elif chosen_coffee == 5:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_5'
                        tick += duration_coffee
                    elif chosen_coffee == 6:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_6'
                        tick += duration_coffee
                    elif chosen_coffee == 7:
                        next_space = ['office']  
                        print("maar naar kantoor 2 uur")
                        print(tick)

                 
                    
                    
                    
                    
                elif next_space == ['restroom']:
                   bezet_toilet_1 = False
                   for i in range(nr_of_schedules):
                       other_schedule = schedules[f'schedule{i}']
                       toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_1'))
                       if True in toilet_occupied_list:
                           bezet_toilet_1 = True
                           
                   bezet_toilet_2 = False
                   for i in range(nr_of_schedules):
                       other_schedule = schedules[f'schedule{i}']
                       toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_2'))
                       if True in toilet_occupied_list:
                           bezet_toilet_2 = True
                           
                   bezet_toilet_3 = False
                   for i in range(nr_of_schedules):
                       other_schedule = schedules[f'schedule{i}']
                       toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_3'))
                       if True in toilet_occupied_list:
                           bezet_toilet_3 = True
                           
                   bezet_toilet_4 = False
                   for i in range(nr_of_schedules):
                       other_schedule = schedules[f'schedule{i}']
                       toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_4'))
                       if True in toilet_occupied_list:
                           bezet_toilet_4 = True
                           
                   
                   if bezet_toilet_1 == True:
                       if bezet_toilet_2 == False:
                           if bezet_toilet_3 == False:
                               if bezet_toilet_4 == False:
                                   chosen_toilet = random.choice([2,3,4])
                               elif bezet_toilet_4 == True:
                                   chosen_toilet = random.choice([2,3])
                           elif bezet_toilet_3 == True:
                               if bezet_toilet_4 == False:
                                   chosen_toilet = random.choice([2,4])
                               elif bezet_toilet_4 == True:
                                   chosen_toilet = 2
                       elif bezet_toilet_2 == True:
                           if bezet_toilet_3 == False:
                               chosen_toilet = 3
                           elif bezet_toilet_3 == True:
                               if bezet_toilet_4 == False:
                                   chosen_toilet = 4
                               elif bezet_toilet_4 == True:
                                       chosen_toilet = random.choice([1,2,3,4])         
                                       
                   elif bezet_toilet_2 ==True:
                       if bezet_toilet_3 == False:
                           chosen_toilet = random.choice([1,3])
                       elif bezet_toilet_3 == True:
                           chosen_toilet = random.choice([1,4])
                   elif bezet_toilet_3 == True:
                       if bezet_toilet_4 == True:
                           chosen_toilet = random.choice([1,2])
                       elif bezet_toilet_4 == False:
                           chosen_toilet = random.choice([1,2,4])
                   elif bezet_toilet_4 == True:
                       chosen_toilet = random.choice([1,2,3])
                   else:         
                       chosen_toilet = random.choice([1,2,3,4])
                                
                               
                          
                               
                           
                  #print(bezet_toilet_1, bezet_toilet_2, bezet_toilet_3, bezet_toilet_4)
                   if chosen_toilet == 1:
                       my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_1'
                       tick += duration_toilet
                       my_schedule[tick:(tick+duration_sink)] = 'sink_1'
                       tick += duration_sink
                   elif chosen_toilet == 2:
                       my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_2'
                       tick += duration_toilet
                       my_schedule[tick:(tick+duration_sink)] = 'sink_2'
                       tick += duration_sink 
                   elif chosen_toilet == 3:
                       my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_3'
                       tick += duration_toilet
                       my_schedule[tick:(tick+duration_sink)] = 'sink_3'
                       tick += duration_sink  
                   elif chosen_toilet == 4:
                       my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_4'
                       tick += duration_toilet
                       my_schedule[tick:(tick+duration_sink)] = 'sink_4'
                       tick += duration_sink 
                    
                    
                elif next_space == ['office']:
                    office_time = random.choice([30,60,120])
                    
                    if office_time == 30:
                        my_schedule.iloc[tick:(tick+30),0] = office_activity+'_30'
                        tick += 30   #hoeft niet + 31 want laatste ding van : word tniet genomen

                        
                        
                        
                        
                    elif office_time == 60:
                        new_tick = tick + 60
                        my_schedule.iloc[tick:(tick+60),0] = office_activity+'_60'
                        tick += 60

                        
                        
                        
                    elif office_time == 120:
                        new_tick = tick + 120
                        my_schedule.iloc[tick:(tick+120),0] = office_activity+'_120'
                        tick += 120                    
     
                    
            elif my_schedule.iloc[(tick+1):(tick+61), 0].isnull().all():
                #print("uur over")
                coffee_list = list(my_schedule.iloc[(tick-5):(tick-3),0].str.contains('coffee'))
                toilet_list = list(my_schedule.iloc[(tick-5):(tick-3),0].str.contains('sink'))

                if True in coffee_list:
                    next_space = random.choices(['office', 'restroom'],transitionmatrix_toilet_office.iloc[row,:])#kies tussen toilet and office
                    

                       
                elif True in toilet_list:
                   next_space = random.choices(['office', 'coffee'],transitionmatrix_coffee_office.iloc[row,:])#kies tussen wc en kantoor
                   
               
                else:
                    next_space = random.choices(['office', 'coffee', 'restroom'], transitionmatrix.iloc[row,:])

                
                
                        
                        
                if next_space == ['coffee']:
                    bezet_coffee_1 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_1_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_1'))
                        if True in coffee_occupied_1_list:
                            bezet_coffee_1 = True
                            
                    bezet_coffee_2 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_2_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_2'))
                        if True in coffee_occupied_2_list:
                            bezet_coffee_2 = True
                            
                    bezet_coffee_3 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_3_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_3'))
                        if True in coffee_occupied_3_list:
                            bezet_coffee_3 = True
                            
                    bezet_coffee_4 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_4_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_4'))
                        if True in coffee_occupied_4_list:
                            bezet_coffee_4 = True
                            
                            
                    bezet_coffee_5 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_5_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_5'))
                        if True in coffee_occupied_5_list:
                            bezet_coffee_5 = True
                            
                    bezet_coffee_6 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_6_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_6'))
                        if True in coffee_occupied_6_list:
                            bezet_coffee_6 = True
                            
                    
                    if bezet_coffee_1 == False:
                        chosen_coffee = 1
                    elif bezet_coffee_1 == True:
                         if bezet_coffee_2 == False:
                             chosen_coffee = 2
                         elif bezet_coffee_2 == True:
                             if bezet_coffee_3 == False:
                                 chosen_coffee= 3
                             elif bezet_coffee_3 == True:
                                 if bezet_coffee_4 == False:
                                     chosen_coffee = 4
                                 elif bezet_coffee_4 == True:
                                     if bezet_coffee_5 == False:
                                         chosen_coffee = 5
                                     elif bezet_coffee_5 == True:
                                         if bezet_coffee_6 == False:
                                             chosen_coffee = 6
                                         elif bezet_coffee_6 == True:
                                             chosen_coffee = 7
                                             print("alle koffie al bezet 1 uur")
                                             print(tick)
                 
                    
                    if chosen_coffee == 1:
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_1'
                        tick += duration_coffee
                    elif chosen_coffee == 2:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_2'
                        tick += duration_coffee
                    elif chosen_coffee == 3:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_3'
                        tick += duration_coffee
                    elif chosen_coffee == 4:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_4'
                        tick += duration_coffee
                    elif chosen_coffee == 5:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_5'
                        tick += duration_coffee
                    elif chosen_coffee == 6:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_6'
                        tick += duration_coffee
                    elif chosen_coffee == 7:
                        next_space = ['office']  
                        print("maar naar kantoor 1 uur")
                        print(tick)

                    
                elif next_space == ['restroom']:
                   # print("toilet 1 uur")
                   bezet_toilet_1 = False
                   for i in range(nr_of_schedules):
                       other_schedule = schedules[f'schedule{i}']
                       toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_1'))
                       if True in toilet_occupied_list:
                           bezet_toilet_1 = True
                           
                   bezet_toilet_2 = False
                   for i in range(nr_of_schedules):
                       other_schedule = schedules[f'schedule{i}']
                       toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_2'))
                       if True in toilet_occupied_list:
                           bezet_toilet_2 = True
                           
                   bezet_toilet_3 = False
                   for i in range(nr_of_schedules):
                       other_schedule = schedules[f'schedule{i}']
                       toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_3'))
                       if True in toilet_occupied_list:
                           bezet_toilet_3 = True
                           
                   bezet_toilet_4 = False
                   for i in range(nr_of_schedules):
                       other_schedule = schedules[f'schedule{i}']
                       toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_4'))
                       if True in toilet_occupied_list:
                           bezet_toilet_4 = True
                           
                   
                   if bezet_toilet_1 == True:
                       if bezet_toilet_2 == False:
                           if bezet_toilet_3 == False:
                               if bezet_toilet_4 == False:
                                   chosen_toilet = random.choice([2,3,4])
                               elif bezet_toilet_4 == True:
                                   chosen_toilet = random.choice([2,3])
                           elif bezet_toilet_3 == True:
                               if bezet_toilet_4 == False:
                                   chosen_toilet = random.choice([2,4])
                               elif bezet_toilet_4 == True:
                                   chosen_toilet = 2
                       elif bezet_toilet_2 == True:
                           if bezet_toilet_3 == False:
                               chosen_toilet = 3
                           elif bezet_toilet_3 == True:
                               if bezet_toilet_4 == False:
                                   chosen_toilet = 4
                               elif bezet_toilet_4 == True:
                                       chosen_toilet = random.choice([1,2,3,4])         
                                       
                   elif bezet_toilet_2 ==True:
                       if bezet_toilet_3 == False:
                           chosen_toilet = random.choice([1,3])
                       elif bezet_toilet_3 == True:
                           chosen_toilet = random.choice([1,4])
                   elif bezet_toilet_3 == True:
                       if bezet_toilet_4 == True:
                           chosen_toilet = random.choice([1,2])
                       elif bezet_toilet_4 == False:
                           chosen_toilet = random.choice([1,2,4])
                   elif bezet_toilet_4 == True:
                       chosen_toilet = random.choice([1,2,3])
                   else:         
                       chosen_toilet = random.choice([1,2,3,4])
                               
                           
                               
                           
                   
                   if chosen_toilet == 1:
                       my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_1'
                       tick += duration_toilet
                       my_schedule[tick:(tick+duration_sink)] = 'sink_1'
                       tick += duration_sink
                   elif chosen_toilet == 2:
                       my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_2'
                       tick += duration_toilet
                       my_schedule[tick:(tick+duration_sink)] = 'sink_2'
                       tick += duration_sink 
                   elif chosen_toilet == 3:
                       my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_3'
                       tick += duration_toilet
                       my_schedule[tick:(tick+duration_sink)] = 'sink_3'
                       tick += duration_sink  
                   elif chosen_toilet == 4:
                       my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_4'
                       tick += duration_toilet
                       my_schedule[tick:(tick+duration_sink)] = 'sink_4'
                       tick += duration_sink 
                    
      
                elif next_space == ['office']:
                    office_time = random.choice([30,60])
                    
                    if office_time == 30:
                        my_schedule.iloc[tick:(tick+30),0] = office_activity+'_30'
                        tick += 30
                        
                        
                        
                    elif office_time == 60:
                        my_schedule.iloc[tick:(tick+60),0] = office_activity+'_60'
                        tick += 60    
                    
            elif my_schedule.iloc[(tick+1):(tick+31), 0].isnull().all():
                #print("30 min over")
                coffee_list = list(my_schedule.iloc[(tick-5):(tick-3),0].str.contains('coffee'))
                toilet_list = list(my_schedule.iloc[(tick-5):(tick-3),0].str.contains('sink'))
                if True in coffee_list:
                    next_space = random.choices(['office', 'restroom'],transitionmatrix_toilet_office.iloc[row,:])#kies tussen toilet and office
                    

                        
                elif True in toilet_list:
                    next_space = random.choices(['office', 'coffee'],transitionmatrix_coffee_office.iloc[row,:])#kies tussen wc en kantoor
                    #print("chosen trans")
                
                else:
                    next_space = random.choices(['office', 'coffee', 'restroom'], transitionmatrix.iloc[row,:])
                #print(next_space)
                #print(tick)
                
                

                        
                if next_space == ['coffee']:
                    bezet_coffee_1 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_1_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_1'))
                        if True in coffee_occupied_1_list:
                            #print(i)
                            bezet_coffee_1 = True
                            
                    bezet_coffee_2 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_2_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_2'))
                        if True in coffee_occupied_2_list:
                            #print(i)
                            bezet_coffee_2 = True
                            
                    bezet_coffee_3 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_3_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_3'))
                        if True in coffee_occupied_3_list:
                            #print(i)
                            bezet_coffee_3 = True
                            
                    bezet_coffee_4 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_4_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_4'))
                        if True in coffee_occupied_4_list:
                            #print(i)
                            bezet_coffee_4 = True
                            
                    bezet_coffee_5 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_5_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_5'))
                        if True in coffee_occupied_5_list:
                            #print(i)
                            bezet_coffee_5 = True
                            
                    bezet_coffee_6 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        coffee_occupied_6_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_6'))
                        if True in coffee_occupied_6_list:
                            #print(i)
                            bezet_coffee_6 = True
                            
                    
                    if bezet_coffee_1 == False:
                        chosen_coffee = 1
                    elif bezet_coffee_1 == True:
                         if bezet_coffee_2 == False:
                             chosen_coffee = 2
                         elif bezet_coffee_2 == True:
                             if bezet_coffee_3 == False:
                                 chosen_coffee= 3
                             elif bezet_coffee_3 == True:
                                 if bezet_coffee_4 == False:
                                     chosen_coffee = 4
                                 elif bezet_coffee_4 == True:
                                     if bezet_coffee_5 == False:
                                         chosen_coffee = 5
                                     elif bezet_coffee_5 == True:
                                         if bezet_coffee_6 == False:
                                             chosen_coffee = 6
                                         elif bezet_coffee_6 == True:
                                             chosen_coffee = 7 
                                             print("alle koffie al bezet 30 min")
                                             print(tick)
                 
                    
                    if chosen_coffee == 1:
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_1'
                        tick += duration_coffee
                    elif chosen_coffee == 2:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_2'
                        tick += duration_coffee
                    elif chosen_coffee == 3:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_3'
                        tick += duration_coffee
                    elif chosen_coffee == 4:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_4'
                        tick += duration_coffee
                    elif chosen_coffee == 5:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_5'
                        tick += duration_coffee
                    elif chosen_coffee == 6:    
                        my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_6'
                        tick += duration_coffee
                    elif chosen_coffee == 7:
                        next_space = ['office']  
                        print("maar naar kantoor 30 min")
                        print(tick)
 
                    
                elif next_space == ['restroom']:
                    bezet_toilet_1 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_1'))
                        if True in toilet_occupied_list:
                            #print(i)
                            bezet_toilet_1 = True
                            
                    bezet_toilet_2 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_2'))
                        if True in toilet_occupied_list:
                            #print(i)
                            bezet_toilet_2 = True
                            
                    bezet_toilet_3 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_3'))
                        if True in toilet_occupied_list:
                            #print(i)
                            bezet_toilet_3 = True
                            
                    bezet_toilet_4 = False
                    for i in range(nr_of_schedules):
                        other_schedule = schedules[f'schedule{i}']
                        toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_4'))
                        if True in toilet_occupied_list:
                            #print(i)
                            bezet_toilet_4 = True
                            
                    
                    if bezet_toilet_1 == True:
                        if bezet_toilet_2 == False:
                            if bezet_toilet_3 == False:
                                if bezet_toilet_4 == False:
                                    chosen_toilet = random.choice([2,3,4])
                                elif bezet_toilet_4 == True:
                                    chosen_toilet = random.choice([2,3])
                            elif bezet_toilet_3 == True:
                                if bezet_toilet_4 == False:
                                    chosen_toilet = random.choice([2,4])
                                elif bezet_toilet_4 == True:
                                    chosen_toilet = 2
                        elif bezet_toilet_2 == True:
                            if bezet_toilet_3 == False:
                                chosen_toilet = 3
                            elif bezet_toilet_3 == True:
                                if bezet_toilet_4 == False:
                                    chosen_toilet = 4
                                elif bezet_toilet_4 == True:
                                        chosen_toilet = random.choice([1,2,3,4])         
                                        
                    elif bezet_toilet_2 ==True:
                        if bezet_toilet_3 == False:
                            chosen_toilet = random.choice([1,3])
                        elif bezet_toilet_3 == True:
                            chosen_toilet = random.choice([1,4])
                    elif bezet_toilet_3 == True:
                        if bezet_toilet_4 == True:
                            chosen_toilet = random.choice([1,2])
                        elif bezet_toilet_4 == False:
                            chosen_toilet = random.choice([1,2,4])
                    elif bezet_toilet_4 == True:
                        chosen_toilet = random.choice([1,2,3])
                    else:         
                        chosen_toilet = random.choice([1,2,3,4])
                                
                                
                            
                                
                            
                    
                    if chosen_toilet == 1:
                        my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_1'
                        tick += duration_toilet
                        my_schedule[tick:(tick+duration_sink)] = 'sink_1'
                        tick += duration_sink
                    elif chosen_toilet == 2:
                        my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_2'
                        tick += duration_toilet
                        my_schedule[tick:(tick+duration_sink)] = 'sink_2'
                        tick += duration_sink 
                    elif chosen_toilet == 3:
                        my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_3'
                        tick += duration_toilet
                        my_schedule[tick:(tick+duration_sink)] = 'sink_3'
                        tick += duration_sink  
                    elif chosen_toilet == 4:
                        my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_4'
                        tick += duration_toilet
                        my_schedule[tick:(tick+duration_sink)] = 'sink_4'
                        tick += duration_sink 
                        
                elif next_space == ['office']:
                    my_schedule.iloc[tick:(tick+30),0] = office_activity+'_30'
                    tick += 30
                        

            
                 #als 3 minuten over 
        
            elif tick < 536:
                #print(my_schedule.iloc[(tick+3), 0])
                #print(tick)
                if type(my_schedule.iloc[(tick+3), 0]) == str:
                    
                    #print("3min over")

                    if type(my_schedule.iloc[(tick+2), 0]) == float:
                        coffee_list = list(my_schedule.iloc[(tick-5):(tick-3),0].str.contains('coffee'))
                        toilet_list = list(my_schedule.iloc[(tick-5):(tick-3),0].str.contains('sink'))
                        if True in coffee_list:
                            next_space =  ['restroom']
                            #print("restroom sws")
 
                                
                        elif True in toilet_list:
                            next_space = ['coffee']
                            #print("nextspace coffee gelijk")
                         #   my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee'
                          #  tick += duration_coffee 
                        
                                
                                
                        else:        
                            next_space = random.choices(['coffee', 'restroom'],transitionmatrix_less.iloc[row,:])
                            
                       # print(next_space)
                        
                        if next_space == ['coffee']:
                            bezet_coffee_1 = False
                            for i in range(nr_of_schedules):
                                other_schedule = schedules[f'schedule{i}']
                                coffee_occupied_1_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_1'))
                                if True in coffee_occupied_1_list:
                                    #print(i)
                                    bezet_coffee_1 = True
                                    
                            bezet_coffee_2 = False
                            for i in range(nr_of_schedules):
                                other_schedule = schedules[f'schedule{i}']
                                coffee_occupied_2_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_2'))
                                if True in coffee_occupied_2_list:
                                    #print(i)
                                    bezet_coffee_2 = True
                                    
                            bezet_coffee_3 = False
                            for i in range(nr_of_schedules):
                                other_schedule = schedules[f'schedule{i}']
                                coffee_occupied_3_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_3'))
                                if True in coffee_occupied_3_list:
                                    #print(i)
                                    bezet_coffee_3 = True
                                    
                            bezet_coffee_4 = False
                            for i in range(nr_of_schedules):
                                other_schedule = schedules[f'schedule{i}']
                                coffee_occupied_4_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_4'))
                                if True in coffee_occupied_4_list:
                                    #print(i)
                                    bezet_coffee_4 = True
                                    
                            bezet_coffee_5 = False
                            for i in range(nr_of_schedules):
                                other_schedule = schedules[f'schedule{i}']
                                coffee_occupied_5_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_5'))
                                if True in coffee_occupied_5_list:
                                    #print(i)
                                    bezet_coffee_5 = True
                                    
                            bezet_coffee_6 = False
                            for i in range(nr_of_schedules):
                                other_schedule = schedules[f'schedule{i}']
                                coffee_occupied_6_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_6'))
                                if True in coffee_occupied_6_list:
                                    #print(i)
                                    bezet_coffee_6 = True
                                    
                            
                            if bezet_coffee_1 == False:
                                chosen_coffee = 1
                            elif bezet_coffee_1 == True:
                                 if bezet_coffee_2 == False:
                                     chosen_coffee = 2
                                 elif bezet_coffee_2 == True:
                                     if bezet_coffee_3 == False:
                                         chosen_coffee= 3
                                     elif bezet_coffee_3 == True:
                                         if bezet_coffee_4 == False:
                                             chosen_coffee = 4
                                         elif bezet_coffee_4 == True:
                                             if bezet_coffee_5 == False:
                                                 chosen_coffee = 5
                                             elif bezet_coffee_5 == True:
                                                 if bezet_coffee_6 == False:
                                                     chosen_coffee = 6
                                                 elif bezet_coffee_6 == True:
                                                     chosen_coffee = 7 
                                                     print("coffee bezet 3 min")
                                                     print(tick)
                         
                            
                            if chosen_coffee == 1:
                                my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_1'
                                tick += duration_coffee
                            elif chosen_coffee == 2:    
                                my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_2'
                                tick += duration_coffee
                            elif chosen_coffee == 3:    
                                my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_3'
                                tick += duration_coffee
                            elif chosen_coffee == 4:    
                                my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_4'
                                tick += duration_coffee 
                            elif chosen_coffee == 5:    
                                my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_5'
                                tick += duration_coffee
                            elif chosen_coffee == 6:    
                                my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_6'
                                tick += duration_coffee 
                            elif chosen_coffee ==7:
                                my_schedule[tick:(tick+1)] = office_activity+'_1m'
                                tick += 2

                        
                        elif next_space == ['restroom']:
                            bezet_toilet_1 = False
                            for i in range(nr_of_schedules):
                                other_schedule = schedules[f'schedule{i}']
                                toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_1'))
                                if True in toilet_occupied_list:
                                    #print(i)
                                    bezet_toilet_1 = True
                                    
                            bezet_toilet_2 = False
                            for i in range(nr_of_schedules):
                                other_schedule = schedules[f'schedule{i}']
                                toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_2'))
                                if True in toilet_occupied_list:
                                    #print(i)
                                    bezet_toilet_2 = True
                                    
                            bezet_toilet_3 = False
                            for i in range(nr_of_schedules):
                                other_schedule = schedules[f'schedule{i}']
                                toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_3'))
                                if True in toilet_occupied_list:
                                    #print(i)
                                    bezet_toilet_3 = True
                                    
                            bezet_toilet_4 = False
                            for i in range(nr_of_schedules):
                                other_schedule = schedules[f'schedule{i}']
                                toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_4'))
                                if True in toilet_occupied_list:
                                    #print(i)
                                    bezet_toilet_4 = True
                                    
                            
                            if bezet_toilet_1 == True:
                                if bezet_toilet_2 == False:
                                    if bezet_toilet_3 == False:
                                        if bezet_toilet_4 == False:
                                            chosen_toilet = random.choice([2,3,4])
                                        elif bezet_toilet_4 == True:
                                            chosen_toilet = random.choice([2,3])
                                    elif bezet_toilet_3 == True:
                                        if bezet_toilet_4 == False:
                                            chosen_toilet = random.choice([2,4])
                                        elif bezet_toilet_4 == True:
                                            chosen_toilet = 2
                                elif bezet_toilet_2 == True:
                                    if bezet_toilet_3 == False:
                                        chosen_toilet = 3
                                    elif bezet_toilet_3 == True:
                                        if bezet_toilet_4 == False:
                                            chosen_toilet = 4
                                        elif bezet_toilet_4 == True:
                                                chosen_toilet = random.choice([1,2,3,4])         
                                                
                            elif bezet_toilet_2 ==True:
                                if bezet_toilet_3 == False:
                                    chosen_toilet = random.choice([1,3])
                                elif bezet_toilet_3 == True:
                                    chosen_toilet = random.choice([1,4])
                            elif bezet_toilet_3 == True:
                                if bezet_toilet_4 == True:
                                    chosen_toilet = random.choice([1,2])
                                elif bezet_toilet_4 == False:
                                    chosen_toilet = random.choice([1,2,4])
                            elif bezet_toilet_4 == True:
                                chosen_toilet = random.choice([1,2,3])
                            else:         
                                chosen_toilet = random.choice([1,2,3,4])
                                        
                                    
                                        
                            #print(bezet_toilet_1, bezet_toilet_2, bezet_toilet_3, bezet_toilet_4)         
                            
                            if chosen_toilet == 1:
                                my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_1'
                                tick += duration_toilet
                                my_schedule[tick:(tick+duration_sink)] = 'sink_1'
                                tick += duration_sink
                            elif chosen_toilet == 2:
                                my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_2'
                                tick += duration_toilet
                                my_schedule[tick:(tick+duration_sink)] = 'sink_2'
                                tick += duration_sink 
                            elif chosen_toilet == 3:
                                my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_3'
                                tick += duration_toilet
                                my_schedule[tick:(tick+duration_sink)] = 'sink_3'
                                tick += duration_sink  
                            elif chosen_toilet == 4:
                                my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_4'
                                tick += duration_toilet
                                my_schedule[tick:(tick+duration_sink)] = 'sink_4'
                                tick += duration_sink 
                        
                #als minder dan 2 minuten over, blijf 1 minuut in office, en 1 minuut overlaten voor lopen        
                    else:
                        my_schedule[tick:(tick+1)] = office_activity+'_1m'
                        tick += 2
                        #print("797")
                    
                    
                
                #als tussen de 3 en 30 min over, zin binnen 3 min over ding, want die optie moet eerst afgesloten
                else:
                    #print("3 tot 30")
                    #print(tick)
                #print(my_schedule.iloc[(tick-4):tick,0])
                    coffee_list = list(my_schedule.iloc[(tick-5):(tick-3),0].str.contains('coffee'))
                    toilet_list = list(my_schedule.iloc[(tick-5):(tick-3),0].str.contains('sink')) #verandering duration
                #print(coffee_list)
                    if True in coffee_list:      
                        #print("geen coffee")
                        next_space = random.choices(['office', 'restroom'],transitionmatrix_toilet_office.iloc[row,:])#kies tussen toilet and office
                        #print(next_space)
               
                    
                    elif True in toilet_list:
                        #print("geen toilet")
                        next_space = random.choices(['office', 'coffee'],transitionmatrix_coffee_office.iloc[row,:])#kies tussen wc en kantoor
                        #print(next_space)
                    
                    else:    
                        #print("else")
                        next_space = random.choices(['office', 'coffee', 'restroom'],transitionmatrix.iloc[row,:])

                
                    if next_space == ['coffee']:
                        bezet_coffee_1 = False
                        for i in range(nr_of_schedules):
                            other_schedule = schedules[f'schedule{i}']
                            coffee_occupied_1_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_1'))
                            if True in coffee_occupied_1_list:
                                #print(i)
                                bezet_coffee_1 = True
                                
                        bezet_coffee_2 = False
                        for i in range(nr_of_schedules):
                            other_schedule = schedules[f'schedule{i}']
                            coffee_occupied_2_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_2'))
                            if True in coffee_occupied_2_list:
                                #print(i)
                                bezet_coffee_2 = True
                                
                        bezet_coffee_3 = False
                        for i in range(nr_of_schedules):
                            other_schedule = schedules[f'schedule{i}']
                            coffee_occupied_3_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_3'))
                            if True in coffee_occupied_3_list:
                                #print(i)
                                bezet_coffee_3 = True
                                
                        bezet_coffee_4 = False
                        for i in range(nr_of_schedules):
                            other_schedule = schedules[f'schedule{i}']
                            coffee_occupied_4_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_4'))
                            if True in coffee_occupied_4_list:
                                #print(i)
                                bezet_coffee_4 = True
                                
                        bezet_coffee_5 = False
                        for i in range(nr_of_schedules):
                            other_schedule = schedules[f'schedule{i}']
                            coffee_occupied_5_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_5'))
                            if True in coffee_occupied_5_list:
                                #print(i)
                                bezet_coffee_5 = True
                                
                        bezet_coffee_6 = False
                        for i in range(nr_of_schedules):
                            other_schedule = schedules[f'schedule{i}']
                            coffee_occupied_6_list = list(other_schedule.iloc[(tick-2):(tick+1),0].str.contains('coffee_6'))
                            if True in coffee_occupied_6_list:
                                #print(i)
                                bezet_coffee_6 = True
                                
                        
                        if bezet_coffee_1 == False:
                            chosen_coffee = 1
                        elif bezet_coffee_1 == True:
                             if bezet_coffee_2 == False:
                                 chosen_coffee = 2
                             elif bezet_coffee_2 == True:
                                 if bezet_coffee_3 == False:
                                     chosen_coffee= 3
                                 elif bezet_coffee_3 == True:
                                     if bezet_coffee_4 == False:
                                         chosen_coffee = 4
                                     elif bezet_coffee_4 == True:
                                         if bezet_coffee_5 == False:
                                             chosen_coffee = 5
                                         elif bezet_coffee_5 == True:
                                             if bezet_coffee_6 == False:
                                                 chosen_coffee = 6
                                             elif bezet_coffee_6 == True:
                                                  chosen_coffee = 7
                                                  
                     
                        
                        if chosen_coffee == 1:
                            my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_1'
                            tick += duration_coffee
                        elif chosen_coffee == 2:    
                            my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_2'
                            tick += duration_coffee
                        elif chosen_coffee == 3:    
                            my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_3'
                            tick += duration_coffee
                        elif chosen_coffee == 4:    
                            my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_4'
                            tick += duration_coffee 
                        elif chosen_coffee == 5:    
                            my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_5'
                            tick += duration_coffee
                        elif chosen_coffee == 6:    
                            my_schedule.iloc[tick:(tick+duration_coffee),0] = 'coffee_6'
                            tick += duration_coffee
                        elif chosen_coffee ==7:
                            next_space = ['office']  
                            print("maar naar kantoor 3-30")
                            print(tick)
                    
                        
                    elif next_space == ['restroom']:
                        bezet_toilet_1 = False
                        for i in range(nr_of_schedules):
                            other_schedule = schedules[f'schedule{i}']
                            toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_1'))
                            if True in toilet_occupied_list:
                                #print(i)
                                bezet_toilet_1 = True
                                
                        bezet_toilet_2 = False
                        for i in range(nr_of_schedules):
                            other_schedule = schedules[f'schedule{i}']
                            toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_2'))
                            if True in toilet_occupied_list:
                                #print(i)
                                bezet_toilet_2 = True
                                
                        bezet_toilet_3 = False
                        for i in range(nr_of_schedules):
                            other_schedule = schedules[f'schedule{i}']
                            toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_3'))
                            if True in toilet_occupied_list:
                                #print(i)
                                bezet_toilet_3 = True
                                
                        bezet_toilet_4 = False
                        for i in range(nr_of_schedules):
                            other_schedule = schedules[f'schedule{i}']
                            toilet_occupied_list = list(other_schedule.iloc[(tick-1):(tick+1),0].str.contains('toilet_4'))
                            if True in toilet_occupied_list:
                                #print(i)
                                bezet_toilet_4 = True
                                
                        
                        if bezet_toilet_1 == True:
                            if bezet_toilet_2 == False:
                                if bezet_toilet_3 == False:
                                    if bezet_toilet_4 == False:
                                        chosen_toilet = random.choice([2,3,4])
                                    elif bezet_toilet_4 == True:
                                        chosen_toilet = random.choice([2,3])
                                elif bezet_toilet_3 == True:
                                    if bezet_toilet_4 == False:
                                        chosen_toilet = random.choice([2,4])
                                    elif bezet_toilet_4 == True:
                                        chosen_toilet = 2
                            elif bezet_toilet_2 == True:
                                if bezet_toilet_3 == False:
                                    chosen_toilet = 3
                                elif bezet_toilet_3 == True:
                                    if bezet_toilet_4 == False:
                                        chosen_toilet = 4
                                    elif bezet_toilet_4 == True:
                                            chosen_toilet = random.choice([1,2,3,4])         
                                            
                        elif bezet_toilet_2 ==True:
                            if bezet_toilet_3 == False:
                                chosen_toilet = random.choice([1,3])
                            elif bezet_toilet_3 == True:
                                chosen_toilet = random.choice([1,4])
                        elif bezet_toilet_3 == True:
                            if bezet_toilet_4 == True:
                                chosen_toilet = random.choice([1,2])
                            elif bezet_toilet_4 == False:
                                chosen_toilet = random.choice([1,2,4])
                        elif bezet_toilet_4 == True:
                            chosen_toilet = random.choice([1,2,3])
                        else:         
                            chosen_toilet = random.choice([1,2,3,4])
                                    
                                
                                    
                                
                        
                        if chosen_toilet == 1:
                            my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_1'
                            tick += duration_toilet
                            my_schedule[tick:(tick+duration_sink)] = 'sink_1'
                            tick += duration_sink
                        elif chosen_toilet == 2:
                            my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_2'
                            tick += duration_toilet
                            my_schedule[tick:(tick+duration_sink)] = 'sink_2'
                            tick += duration_sink 
                        elif chosen_toilet == 3:
                            my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_3'
                            tick += duration_toilet
                            my_schedule[tick:(tick+duration_sink)] = 'sink_3'
                            tick += duration_sink  
                        elif chosen_toilet == 4:
                            my_schedule.iloc[tick:(tick+duration_toilet),0] = 'toilet_4'
                            tick += duration_toilet
                            my_schedule[tick:(tick+duration_sink)] = 'sink_4'
                            tick += duration_sink 
                         
                     
                    
                    elif next_space == ['office']:
                        while type(my_schedule.iloc[(tick + 1),0]) == float :
                            my_schedule.iloc[tick,0] = office_activity+'_1m'
                            tick += 1
                            
                            
                            
                        if 'mr1' in my_schedule.iloc[(tick+1), 0]: 
                           tick += 57 #moet niet lager want dan meer nan
                           
                        elif 'mr2' in my_schedule.iloc[(tick+1), 0]:  
                           tick += 50  #moet niet lager want dan meer nan     
                        
                      
            
                        
                        elif 'meeting' in my_schedule.iloc[(tick+1), 0]:
                            #print("861")
                            tick += 46

                        
              #als het al 536 is            
            else:
                tick = 540
                
        
           #als wel een waarde in volgende cel        
        else:
            if my_schedule.iloc[(tick+1), 0] =='nothing':
                tick = 541
                #print("nothing true")
                #print(tick)
                
            #when their next location is a meeting
            elif 'mr1' in my_schedule.iloc[(tick+1), 0]:  
                if type(my_schedule.iloc[tick,0]) != float:
                    tick+= 56
                else:
                    tick += 57 #moet niet lager want dan meer nan
                
            elif 'mr2' in my_schedule.iloc[(tick+1), 0]:  
                if type(my_schedule.iloc[tick,0]) != float:
                    tick += 49 #moet niet lager want dan meer nan
                else: 
                    tick+= 50
                
            
                
            elif "meeting" in my_schedule.iloc[(tick+1), 0]:
                if type(my_schedule.iloc[tick,0]) != float:
                    #print(tick)
                    #print("891")
                    tick += 45
                    #print(tick)
                else:
                    tick+= 46
                #print("was meeting")

            else:
                print("error regel 1551")
     
activitypatterns_list = [] 
new_line = '\n'
tab = '\t'
                   
for s in range(nr_of_schedules):
    k = s + 1
    my_schedule = schedules[f'schedule{s}']
    if s  == 0:
        activitypattern = f'<listActivityPattern ID="employee{k}">{new_line}{tab}{tab}{tab}{tab}<IDlist>(' #begin met empty string
    elif s > 0:
        activitypattern = activitypattern +f'{new_line}{tab}{tab}{tab}{tab}<listActivityPattern ID="employee{k}">{new_line}{tab}{tab}{tab}{tab}<IDlist>('   
    for i in range(539):
        activity_now = my_schedule.iloc[i, 0]
        activity_before = my_schedule.iloc[(i-1), 0]
        if '1m' in str(activity_now):
            activitypattern = activitypattern + str(activity_now)
            activitypattern = activitypattern + ','            
        #maak uitzondering voor de 1 minuut office loop
        elif activity_now != activity_before:
            if type(activity_now) != float:
                if activity_now != 'nothing':
                    activitypattern = activitypattern + str(activity_now)
                    activitypattern = activitypattern + ','
    activitypattern = activitypattern + f"leave)</IDlist>{new_line}{tab}{tab}{tab}</listActivityPattern>"
    
    
    

for d in range(nr_of_schedules):
    p = d + 1
    activitypatternID = f'employee{p}'
    discrete_pattern = entertimes_list[d]
    if d == 0:
        demandpattern = f'<discreteDemandPattern ID="demand_discrete{p}">{new_line}{tab}{tab}{tab}{tab}<sourceID>entrance38_ingang</sourceID>{new_line}{tab}{tab}{tab}{tab}<activityPatternID>{activitypatternID}</activityPatternID>{new_line}{tab}{tab}{tab}{tab}<parameterSetID>base</parameterSetID>{new_line}{tab}{tab}{tab}<discretePattern>(({entertimes_list[d]}*60,1))</discretePattern></discreteDemandPattern>'
    elif d > 0:
        
         demandpattern = demandpattern + f'{new_line}{tab}{tab}{tab}<discreteDemandPattern ID="demand_discrete{p}">{new_line}{tab}{tab}{tab}{tab}<sourceID>entrance38_ingang</sourceID>{new_line}{tab}{tab}{tab}{tab}<activityPatternID>{activitypatternID}</activityPatternID>{new_line}{tab}{tab}{tab}{tab}<parameterSetID>base</parameterSetID>{new_line}{tab}{tab}{tab}<discretePattern>(({entertimes_list[d]}*60,1))</discretePattern></discreteDemandPattern>'



            
scenarioinput =  f'{activitypattern} {new_line} </activityPatterns> {new_line} <demandPatterns> {new_line} {demandpattern}{new_line}</demandPatterns> '


now2 = datetime.now()
now_str=str(now2)

now_kort = now_str[0:16]
now_deel1 = now_str[0:13]
now_deel2 = now_str[14:16]
now_goed = now_deel1 + now_deel2

now = datetime.now(pytz.timezone('Europe/Amsterdam'))

file = open(f"scenarioinput_{scenario}_{now_goed}.txt", "w+")
file.write(scenarioinput)
file.close()







wb = xlsxwriter.Workbook(f'schedules-scenario{scenario}_{now_goed}.xlsx')
ws = wb.add_worksheet('schemas')
ws.write_row(0, 0, ['employee1', 'employee2', 'employee3', 'employee4', 'employee5', 'employee6', 'employee7', 'employee8', 'employee9', 'employee10', 'employee11', 'employee12', 'employee13', 'employee14', 'employee15', 'employee16', 'employee17', 'employee18', 'employee19', 'employee20', 'employee21', 'employee22', 'employee23', 'employee24', 'employee25', 'employee26', 'employee27', 'employee28', 'employee29', 'employee30', 'employee31', 'employee32', 'employee33', 'employee34', 'employee35', 'employee36', 'employee37', 'employee38', 'employee39', 'employee40'])
for i in range(40):
    #data = repr(schedules[f'schedule{i}'])
    data2 = (schedules[f'schedule{i}'])
    for j in range (540):
        tijdstap = data2.iloc[j,0]
        if type(tijdstap) == str:
            ws.write(j+1,i,tijdstap)
           # print(tijdstap)
        elif type(tijdstap)== float:
            ws.write(j+1,i,'niks')
            #print('niks')
        else:
            print("iets anders")
    ws.write(550,0,now_kort)
        


wb.close()



#pd.options.display.max_rows = 999
 