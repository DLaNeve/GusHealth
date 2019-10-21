#!/bin/env python

import calendar
import datetime

now = datetime.datetime.now()
cal = calendar.Calendar()

total_wk_days = len([x for x in cal.itermonthdays2(now.year, now.month) if x[0] !=0 and x[1] < 5])
current_month_goal = 500
daily_goal=current_month_goal/total_wk_days
current_month_ttl = 250
completed_wk_days = 5

a=current_month_ttl/current_month_goal
b=completed_wk_days/total_wk_days

print(a,b)


progress = completed_wk_days*daily_goal

percent_days = completed_wk_days/total_wk_days


progress=percent_days*current_month_goal/current_month_goal

print (progress)

